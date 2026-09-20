#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실습학생 현장실습평가표 자동 기입 스크립트
=========================================
입력:  (1) 실습학생 성적산출표 .xlsx   (2) 학교에서 온 빈 평가표 PDF(스캔본, 학생 1명 = 1페이지)
출력:  동그라미·항목합계·출석·종합평가(등급/실습수석)·총점이 기입된 PDF

사용 예)
  python fill_eval.py --xlsx "성적산출표.xlsx" --pdf "학생성적_영상_평가전.pdf" \
                      --out "학생성적_영상_평가완료.pdf" --school 대구보건대학교

규칙(2026 하계 기준)
  - 총점      = 산출표 '성적' 열(79~100)
  - 출석(40)  = 산출표 '출결(40)' 열
  - 12개 항목 = (총점 - 출석)점을 5점 만점 항목 12개에 '골고루' 감점하여 배분
                (같은 항목에 2점 이상 몰리지 않게, 4개 행에 균등, 학생마다 시작 항목을 회전)
  - 종합평가  = A+(95~) / A(90~94) / B+(85~89) / B(80~84) / C+(~79),
                해당 학교 1등(성적 최고점)은 '실습수석'
  - 평가자 서명/도장 칸은 비워 둠 (직접 날인)

필요 패키지:  pip install openpyxl pypdf reportlab pypdfium2 pillow numpy
필요 폰트:    한글 TrueType/OpenType 폰트 (Windows: C:/Windows/Fonts/malgunbd.ttf 자동 탐색)
"""
import argparse, io, os, random, subprocess, sys, json
import numpy as np
import openpyxl
from PIL import Image, ImageFont, ImageDraw
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

# ---------------------------------------------------------------------------
# 양식 레이아웃 (대구보건대 '현장실습평가표', A4 596x842pt, 기준 100dpi 픽셀 좌표)
# 양식이 바뀌면 이 값만 고치거나 --layout layout.json 으로 덮어쓴다.
# ---------------------------------------------------------------------------
DEFAULT_LAYOUT = {
    "dpi": 100,
    "row_y": [400, 479, 570, 659],                       # 4개 '평가' 행의 숫자 중심 y
    "cols": [[204, 231, 259, 287, 314],                  # 항목 1열: 5,4,3,2,1 의 x
             [360, 388, 415, 443, 470],                  # 항목 2열
             [518, 546, 573, 601, 629]],                 # 항목 3열
    "sum_x": 690, "sum_y": [390, 470, 560, 650],         # 행별 합계 칸
    "att": [690, 728],                                   # 출석평가(40) 칸
    "overall": [340, 798],                               # 종합평가 칸
    "total": [702, 798],                                 # 총점 칸
    "align_region": [250, 850, 60, 770],                 # 페이지 정렬용 영역 y0,y1,x0,x1
    "circle_r": 7.0,                                     # 동그라미 반지름(pt)
    "ink": [0.05, 0.10, 0.55]                            # 잉크색 (파란 펜)
}

ORDER = [1, 4, 7, 10, 2, 5, 8, 11, 0, 3, 6, 9]           # 감점 순서(행을 번갈아 가며)


def grade_letter(score):
    return 'A+' if score >= 95 else 'A' if score >= 90 else 'B+' if score >= 85 else 'B' if score >= 80 else 'C+'


def distribute(idx, need):
    """12개 항목(각 5점 만점)에 need점이 되도록 골고루 감점."""
    v = [5] * 12
    d = 60 - need
    if not 0 <= d <= 48:
        raise ValueError(f'항목 합계 {need}는 12~60 범위를 벗어남')
    order = ORDER[idx % 12:] + ORDER[:idx % 12]
    k = 0
    while d > 0:
        v[order[k % 12]] -= 1
        d -= 1
        k += 1
    return v


def load_students(xlsx, school, sheet=None):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb[sheet] if sheet else wb.active
    rows = list(ws.iter_rows(values_only=True))
    # 헤더 행 찾기
    hdr_i = next(i for i, r in enumerate(rows) if r and '이름' in r and '학교' in r)
    hdr = list(rows[hdr_i])
    ci = {name: hdr.index(name) for name in hdr if name}
    att_col = next(c for c in ci if str(c).startswith('출결'))
    students = []
    for r in rows[hdr_i + 1:]:
        if r[ci['학교']] == school and r[ci['이름']]:
            students.append(dict(
                no=r[ci['번호']],
                name=str(r[ci['이름']]).replace('(조장)', '').strip(),
                att=int(round(r[ci[att_col]])),
                score=int(round(r[ci['성적']])),
            ))
    return students


def render_pages(pdf, dpi):
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(pdf)
    return [p.render(scale=dpi / 72).to_pil().convert('L') for p in doc]


def profile(img, reg):
    y0, y1, x0, x1 = reg
    a = (np.array(img) < 190)[y0:y1, x0:x1]
    return a.mean(axis=1), a.mean(axis=0)


def best_shift(a, b, maxs=15):
    best = (0, -1)
    for s in range(-maxs, maxs + 1):
        c = np.dot(a[s:], b[:len(b) - s]) if s >= 0 else np.dot(a[:s], b[-s:])
        if c > best[1]:
            best = (s, c)
    return best[0]


def find_korean_font():
    cands = ['C:/Windows/Fonts/malgunbd.ttf', 'C:/Windows/Fonts/malgun.ttf',
             '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc',
             '/System/Library/Fonts/AppleSDGothicNeo.ttc']
    for c in cands:
        if os.path.exists(c):
            return c
    raise FileNotFoundError('한글 폰트를 찾지 못했습니다. --font 로 지정하세요.')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--xlsx', required=True)
    ap.add_argument('--pdf', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--school', default='대구보건대학교')
    ap.add_argument('--sheet', default=None)
    ap.add_argument('--layout', default=None, help='레이아웃 JSON (DEFAULT_LAYOUT 키 덮어쓰기)')
    ap.add_argument('--font', default=None, help='한글 폰트 파일')
    ap.add_argument('--no-suseok', action='store_true', help="1등에게 '실습수석' 대신 등급 기입")
    ap.add_argument('--seed', type=int, default=2026)
    a = ap.parse_args()

    L = dict(DEFAULT_LAYOUT)
    if a.layout:
        L.update(json.load(open(a.layout, encoding='utf-8')))
    K = 72 / L['dpi']

    students = load_students(a.xlsx, a.school, a.sheet)
    reader = PdfReader(a.pdf)
    if len(students) != len(reader.pages):
        sys.exit(f'[중단] {a.school} 학생 {len(students)}명 vs PDF {len(reader.pages)}페이지 — 인원이 다릅니다.')
    best = max(students, key=lambda s: s['score'])

    fontfile = a.font or find_korean_font()
    pil_font = ImageFont.truetype(fontfile, 120, index=1 if fontfile.endswith('.ttc') else 0)
    def kr_image(text):
        im = Image.new('RGBA', (900, 180), (0, 0, 0, 0))
        ImageDraw.Draw(im).text((10, 10), text, font=pil_font,
                                fill=tuple(int(c * 255) for c in L['ink']) + (255,))
        return im.crop(im.getbbox())
    # 숫자/영문은 reportlab 내장 Helvetica-Bold (폰트 임베딩 문제 없음)
    NUMFONT = 'Helvetica-Bold'

    pages = render_pages(a.pdf, L['dpi'])
    ref = profile(pages[0], L['align_region'])
    rng = random.Random(a.seed)
    writer = PdfWriter()
    r = L['circle_r']

    print(f"{'No':>3} {'이름':<6} 출석  항목점수                              행합계          총점 종합")
    for i, (pg, img, st) in enumerate(zip(reader.pages, pages, students)):
        p = profile(img, L['align_region'])
        dy, dx = best_shift(ref[0], p[0]), best_shift(ref[1], p[1])
        def P(x, y):
            return ((x - dx) * K, 842 - (y - dy) * K)
        v = distribute(i, st['score'] - st['att'])
        sums = [sum(v[k * 3:k * 3 + 3]) for k in range(4)]
        assert sum(sums) + st['att'] == st['score']

        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(596, 842))
        c.setStrokeColorRGB(*L['ink']); c.setFillColorRGB(*L['ink']); c.setLineWidth(1.1)
        for row in range(4):
            for col in range(3):
                x, y = P(L['cols'][col][5 - v[row * 3 + col]], L['row_y'][row])
                j = lambda: rng.uniform(-.5, .5)
                c.ellipse(x - r + j(), y - r + 1 + j(), x + r + j(), y + r - 1 + j())
            c.setFont(NUMFONT, 15)
            x, y = P(L['sum_x'], L['sum_y'][row]); c.drawCentredString(x, y - 5, str(sums[row]))
        c.setFont(NUMFONT, 16); x, y = P(*L['att']); c.drawCentredString(x, y - 5, str(st['att']))
        x, y = P(*L['overall'])
        if st is best and not a.no_suseok:
            im = kr_image('실습수석'); H = 14; W = im.size[0] * H / im.size[1]
            c.drawImage(ImageReader(im), x - W / 2, y - H / 2, W, H, mask='auto'); overall = '실습수석'
        else:
            overall = grade_letter(st['score']); c.setFont(NUMFONT, 15); c.drawCentredString(x, y - 5, overall)
        c.setFont(NUMFONT, 17); x, y = P(*L['total']); c.drawCentredString(x, y - 6, str(st['score']))
        c.save(); buf.seek(0)
        pg.merge_page(PdfReader(buf).pages[0]); writer.add_page(pg)
        print(f"{st['no']:>3} {st['name']:<6} {st['att']:>3}  {v}  {sums}  {st['score']:>3} {overall}")

    with open(a.out, 'wb') as f:
        writer.write(f)
    print('\n완료:', a.out)
    print('※ PDF 페이지 순서가 산출표 번호 순서와 같은지(학번·이름) 반드시 눈으로 확인하세요.')


if __name__ == '__main__':
    main()
