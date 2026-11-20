# 첨부 발표자료(디자인 유지) 수정: python3 edit_attached.py <in.pptx> <out.pptx>
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

src, out = sys.argv[1], sys.argv[2]
p = Presentation(src)
EMU = 914400
def shp(s, sid): return [x for x in s.shapes if x.shape_id == sid][0]

# 1) 소제목(eyebrow) 줄바꿈으로 제목과 겹치는 문제: 폭 넓히고 줄바꿈 끔
for s in p.slides:
    for sh in s.shapes:
        if sh.has_text_frame and abs(sh.top / EMU - 0.71) < 0.03 and sh.height / EMU < 0.25 and sh.text_frame.text.strip():
            sh.width = Inches(8); sh.text_frame.word_wrap = False

# 2) 작은 글씨 키우기(본문 영역, 슬라이드별 하한)
bump = {5: {9.5: 11}, 7: {11.0: 13, 11.5: 13}, 10: {9.5: 11}, 11: {10.0: 11.5}, 13: {10.0: 11.5},
        16: {11.0: 13, 9.5: 11}, 17: {11.0: 12.5}}
for si, rule in bump.items():
    for sh in p.slides[si - 1].shapes:
        if not sh.has_text_frame or not (1.6 <= sh.top / EMU <= 6.9): continue
        for para in sh.text_frame.paragraphs:
            for r in para.runs:
                if r.font.size and round(r.font.size.pt, 2) in rule:
                    r.font.size = Pt(rule[round(r.font.size.pt, 2)])

# 3) 03 차트 연한 막대 진하게
ch = [x for x in p.slides[2].shapes if x.has_chart][0].chart
for pt_i in range(6):
    pt = ch.plots[0].series[0].points[pt_i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = RGBColor.from_string("2563EB" if pt_i == 5 else "7DA3EE")

# 4) 05 차트: 동별 65세 이상 → 85세 이상 1인세대(부곡동 강조)
s5 = p.slides[4]
ch = [x for x in s5.shapes if x.has_chart][0].chart
cd = CategoryChartData()
cd.categories = ["부곡동", "청계동", "내손1동", "오전동", "내손2동", "고천동"]
cd.add_series("85세 이상 1인세대", (292, 211, 177, 156, 114, 96))
ch.replace_data(cd)
ser = ch.plots[0].series[0]
for i in range(6):
    pt = ser.points[i]; pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = RGBColor.from_string("0284C7" if i == 0 else "6AB4DD")
shp(s5, 18).text_frame.paragraphs[0].runs[0].text = "동별 85세 이상 1인세대 (세대)"
shp(s5, 20).text_frame.paragraphs[0].runs[0].text = "단위: 세대"

# 5) 10 하단 안내 → 발굴 흐름 한 줄
r = shp(p.slides[9], 24).text_frame.paragraphs[0].runs[0]
r.text = "발굴 흐름: 기초연금 명단(85세 이상 1인가구) → 신청자 제외 → 생활지원사 선별 → 동 전담인력 방문 → 직권신청"

# 6) 12 모니터링 표 → 타임라인
s12 = p.slides[11]
for sid in range(35, 48):
    el = shp(s12, sid)._element; el.getparent().remove(el)
ACC, DARK, SUB, LINE = "A24B08", "111827", "475569", "CBD5E1"
def tb(text, x, y, w, h, size, color, bold=False, align=PP_ALIGN.CENTER):
    t = s12.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = t.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for k, line in enumerate(text.split("\n")):
        para = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        para.alignment = align
        rr = para.add_run(); rr.text = line
        rr.font.size = Pt(size); rr.font.bold = bold; rr.font.name = "Pretendard"; rr.font.color.rgb = RGBColor.from_string(color)
    return t
tb("모니터링 주기 · 동 전담인력", 0.9, 4.2, 5, 0.3, 13, DARK, True, PP_ALIGN.LEFT)
ln = s12.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1.2), Inches(5.3), Inches(7.25), Inches(5.3))
ln.line.color.rgb = RGBColor.from_string(LINE); ln.line.width = Pt(3)
nodes = [(1.55, "계획 안내", "개인별지원계획\n확정 통보", SUB), (4.22, "1개월 이내", "최초 모니터링\n서비스 개시 확인(대면·유선)", ACC), (6.9, "3개월마다", "정기 모니터링\n유선, 필요 시 대면", SUB)]
for cx, top, bottom, col in nodes:
    d = 0.34
    c = s12.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(5.3 - d / 2), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string(col); c.line.fill.background()
    tb(top, cx - 1.2, 4.68, 2.4, 0.35, 14, col if col == ACC else DARK, True)
    tb(bottom, cx - 1.25, 5.62, 2.5, 0.7, 11.5, SUB)

# 7) 17 제목: 88%는 신규 3종 포함 금액이므로 표현 정정
shp(p.slides[16], 8).text_frame.paragraphs[0].runs[2].text = ", 의왕살핌·신규서비스에 88%를 씁니다"

# 8) 표지 제목: 디자인 레터링 이미지(글꼴 미설치 PC에서도 동일하게 보임)
import os
cov = p.slides[0]
t = shp(cov, 6)
alt = t.text_frame.text.replace("\n", " ")
t._element.getparent().remove(t._element)
img = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "figures", "표지_제목_레터링.png")
pic = cov.shapes.add_picture(img, Inches(0.83), Inches(2.85), width=Inches(9.0))
shp(cov, 7).top = Inches(4.95)
pic.name = "표지 제목"
pic._element.nvPicPr.cNvPr.set("descr", "2027년도 통합지원 지역계획 — 살던 곳에서 건강하게, 함께 살피는 의왕")

p.save(out); print("saved", out)
