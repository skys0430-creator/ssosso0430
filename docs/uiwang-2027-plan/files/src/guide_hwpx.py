# 제2차 통합지원협의체 회의 안내지(1쪽, HWPX): python3 guide_hwpx.py <out.hwpx>
import sys, warnings
warnings.filterwarnings("ignore")
from hwpx.document import HwpxDocument

doc = HwpxDocument.new()
doc.set_page_setup(paper_size="A4", orientation="portrait", margins_mm={"left": 20, "right": 20, "top": 20, "bottom": 18})
F = "맑은 고딕"; HD = "HY헤드라인M"
S = lambda **k: doc.ensure_run_style(font=k.pop("font", F), **k)
st = {
 "title": S(bold=True, size=19, font=HD),
 "date": S(size=11.5, color="#555555"),
 "h": S(bold=True, size=13.5, color="#1B3A5C"),
 "b": S(size=11.5), "bb": S(bold=True, size=11.5),
 "th": S(bold=True, size=10.5, color="#FFFFFF"),
 "td": S(size=10.5), "tdb": S(bold=True, size=10.5),
 "bx": S(size=11), "bxb": S(bold=True, size=11, color="#1B3A5C"),
 "note": S(size=10, color="#555555"),
}
pc = doc.styles.ensure_paragraph_format(alignment="CENTER", line_spacing_percent=150)
pb = doc.styles.ensure_paragraph_format(line_spacing_percent=155)
pi = doc.styles.ensure_paragraph_format(line_spacing_percent=150, margins={"left": 1400})
pt = doc.styles.ensure_paragraph_format(alignment="CENTER", line_spacing_percent=130)
pl = doc.styles.ensure_paragraph_format(line_spacing_percent=140)

first = [True]
def para(runs, pp=pb):
    if first[0]: p = doc.paragraphs[0]; first[0] = False
    else: p = doc.add_paragraph("")
    p.para_pr_id_ref = pp; p.clear_text()
    for t, s in runs: p.add_run(t, char_pr_id_ref=st[s])
    return p
def cell(t, r, c, runs_list, pp=pt):
    ce = t.cell(r, c)
    for k, runs in enumerate(runs_list):
        p = ce.paragraphs[0] if k == 0 else ce.add_paragraph("")
        p.para_pr_id_ref = pp; p.clear_text()
        for tx, s in runs: p.add_run(tx, char_pr_id_ref=st[s])

# 제목
para([("2026년 제2차 의왕시 통합지원협의체 회의", "title")], pc)
para([("2026. 11. 26.(목) 15:00  |  의왕시청 중회의실", "date")], pc)
para([("", "b")])

# □ 회의 개요
para([("□ 회의 개요", "h")])
for k, v in [("일    시", "2026. 11. 26.(목) 15:00 ~ 16:00"), ("장    소", "의왕시청 중회의실"),
             ("참석대상", "통합지원협의체 위원 18명(위원장: 의왕시장)"),
             ("안    건", "【보고】 2026년 통합돌봄 추진사항"),
             ("", "【심의】 2027년도 의왕시 통합지원 지역계획(안)")]:
    if k: para([("  ❍ " + k + " : ", "bb"), (v, "b")])
    else: para([("                  " + v, "b")])
para([("", "b")])

# □ 진행 순서
para([("□ 진행 순서", "h")])
rows = [("15:00 ~ 15:02", "2'", "개회", "통합돌봄팀장"),
        ("15:02 ~ 15:07", "5'", "참석자 소개", "통합돌봄팀장"),
        ("15:07 ~ 15:12", "5'", "인사 말씀", "위원장"),
        ("15:12 ~ 15:15", "3'", "성원 보고", "통합돌봄팀장"),
        ("15:15 ~ 15:55", "40'", "안건 보고·심의 (발표 및 질의·의결)", "위원장 진행"),
        ("15:55 ~ 16:00", "5'", "폐회", "통합돌봄팀장")]
t = doc.add_table(len(rows) + 1, 4, width=48500)
t.set_column_widths([24, 10, 44, 22])
for j, h in enumerate(["시  간", "소요", "주 요 내 용", "비  고"]):
    t.set_cell_shading(0, j, "#1B3A5C"); cell(t, 0, j, [[(h, "th")]])
for i, (a, b_, c, d) in enumerate(rows, start=1):
    hi = i == 5
    for j, v in enumerate([a, b_, c, d]):
        if hi: t.set_cell_shading(i, j, "#FCE4F1")
        cell(t, i, j, [[(v, "tdb" if (hi or j == 2) else "td")]])
para([("", "b")])

# □ 심의 안건 주요내용
para([("□ 심의 안건 주요내용 : 2027년도 의왕시 통합지원 지역계획(안)", "h")])
box = doc.add_table(1, 1, width=48500)
box.set_cell_shading(0, 0, "#F6F8FB")
lines = [
 [("▪ 비전  ", "bxb"), ("살던 곳에서 건강하게, 함께 살피는 의왕", "bx")],
 [("▪ 3대 목표", "bxb")],
 [("   - 먼저 찾고 끝까지 살피는 대상자 관리체계 확립", "bx")],
 [("   - 재가 의료 이용 확대와 의왕살핌 고도화로 재가돌봄 공백 해소", "bx")],
 [("   - 보건·복지·민간이 함께하는 수요자 중심 운영기반 강화", "bx")],
 [("▪ 추진체계  ", "bxb"), ("4대 과업 · 12개 핵심사업(신규 5, 확대 4, 개선 3)", "bx")],
 [("▪ 주요 내용", "bxb")],
 [("   - 85세 이상 초고령 가구 다경로 선제발굴, 퇴원환자 의뢰 경로 이원화", "bx")],
 [("   - 방문의료 이용 확대, 의왕살핌 이용자 30% 확대", "bx")],
 [("   - 신규서비스 도입: 이동살핌, AI 돌봄로봇, 회복살핌 꾸러미", "bx")],
 [("▪ 목    표  ", "bxb"), ("통합지원 대상자 541명, 직권·자체발굴 발굴률 20% 이상", "bx")],
 [("▪ 예    산  ", "bxb"), ("706,636천원", "bx")],
]
cell(box, 0, 0, lines, pl)
para([("", "b")])
para([("※ 계획서 전문은 배부자료를 참고하여 주시기 바랍니다.", "note")])

doc.save_to_path(sys.argv[1]); print("saved", sys.argv[1])
