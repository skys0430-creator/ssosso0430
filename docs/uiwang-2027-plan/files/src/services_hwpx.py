# 부록 서비스 한눈에 보기 표(HWPX) 생성: python3 services_hwpx.py <out.hwpx>
import sys
from hwpx.document import HwpxDocument

AREAS = ["보건의료", "건강관리", "장기요양", "일상생활돌봄", "주거복지"]
COLS = [("지역특화서비스", "12개 · 복지정책과 통합돌봄팀"),
        ("연계서비스(국가돌봄)", "17개 · 중앙 제도·공단·보건소"),
        ("연계서비스(도·지자체)", "13개 · 경기도·시 자체·민간자원")]
D = {
 (0,0): [("1-1-1","건강살핌–방문의료","의사·한의사 방문진료 본인부담 지원(의원 14개소)"),
         ("1-1-2","건강살핌–가정간호","퇴원 후 치료·간호 필요자(서울홈케어내과)"),
         ("1-1-3","건강살핌–방문재활","재활 필요 거동불편자(물리치료사협회)"),
         ("1-1-4","재택의료센터 운영지원","참여 의료기관 운영비 지원")],
 (0,1): [("2-1-1","장기요양 재택의료센터","다학제팀 방문진료·간호, 재가 수급자"),
         ("2-1-2","일차의료 방문진료 시범사업","거동불편자 방문진료(참여의원 14개소)"),
         ("2-1-3","방문건강관리","건강취약계층 방문 관리·상담(보건소)")],
 (0,2): [],
 (1,0): [("1-2-1","동고동락(방문형 치매프로그램)","치매진단·예방약 복용자, 주 2회 14회기"),
         ("1-2-2","건강돌봄가","방문 스트레칭·건강·영양 모니터링, 복약 지원")],
 (1,1): [("2-2-1","다제약물관리","10종 이상 복용자 약물 점검(건보공단)"),
         ("2-2-2","만성질환 건강지원서비스","고혈압·당뇨 측정기 대여·상담(건보공단)"),
         ("2-2-3","AI·IoT 어르신 건강관리","기기 연동 비대면 건강관리(보건소)"),
         ("2-2-4","정신건강복지센터 연계","상담·등록 사례관리·위기개입"),
         ("2-2-5","치매환자 지원","치매·인지저하자 맞춤 서비스(치매안심센터)"),
         ("2-2-6","재가의료급여(의료)","퇴원 의료급여수급자 케어플랜(생활보장팀)")],
 (1,2): [("3-2-1","다제약물 복용자 방문 약료","5종 이상 3개월 복용자, 약사 2인 방문"),
         ("3-2-2","장애인 대소변 흡수용품 지원","심한 장애인 구매비 50%(월 5만원)")],
 (2,0): [],
 (2,1): [("2-3-1","노인장기요양보험","방문요양·목욕·간호, 주야간보호, 복지용구")],
 (2,2): [],
 (3,0): [("1-4-1","일상살핌–생활돌봄","신체·가사 지원(1일 2~8시간)"),
         ("1-4-2","일상살핌–동행돌봄","병원·관공서 필수 외출 동행"),
         ("1-4-3","일상살핌–식사지원","도시락 배달(1일 1~3식)"),
         ("1-4-4","영양통합지원","맞춤형 영양식 주 2회 배달·안부확인"),
         ("1-4-5","퇴원환자 연계 지원","협약병원 퇴원 전 의뢰 시 연계수당")],
 (3,1): [("2-4-1","장애인 활동지원","6~65세 미만 등록장애인 신체·가사·사회활동"),
         ("2-4-2","장애인 맞춤형 지원사","저소득 장애인 가사·외출, 산후·육아"),
         ("2-4-3","노인맞춤돌봄","기초·차상위·기초연금 노인 안전·일상 지원"),
         ("2-4-4","노인맞춤돌봄(퇴원환자 단기집중)","퇴원 2주 이내 단기 집중돌봄"),
         ("2-4-5","응급안전안심","독거노인·장애인 ICT 모니터링·119 연계"),
         ("2-4-6","재가의료급여(돌봄·이동·식사)","퇴원 의료급여수급자 가사간병·교통비·식사"),
         ("2-4-7","가사간병 방문지원","중위 70% 이하 65세 미만 중증")],
 (3,2): [("3-4-1","누구나돌봄(경기도)","생활·동행·주거안전·식사(연 150만원)"),
         ("3-4-2","저소득 재가노인 식사배달","거동불편 수급·차상위 노인 1일 1식"),
         ("3-4-3","경기도 AI 노인말벗","주 1회 AI 안부전화·위기징후 조치"),
         ("3-4-4","재가노인지원서비스","위기 노인 사례관리 기반 재가지원"),
         ("3-4-5","어르신 안부 우유배달","저소득 독거노인 우유배달·안부확인")],
 (4,0): [("1-5-1","내집살핌–주거환경개선","낙상예방 손잡이·문턱제거 등(생애 200만원)")],
 (4,1): [],
 (4,2): [("3-5-1","굿모닝 홈클린","주택 청소·생활폐기물 수거"),
         ("3-5-2","어르신 안전하우징","기초연금 노인 안전설비(500만원)"),
         ("3-5-3","햇살하우징","중위 50% 이하 에너지효율 개보수(500만원)"),
         ("3-5-4","주거든든","도배·장판·싱크대·수전 교체"),
         ("3-5-5","G-하우징","기부·재능나눔 집수리(1,000만원)"),
         ("3-5-6","깔끄미","청소·소독·방역(독거노인·장애인 우선)")],
}

doc = HwpxDocument.new()
doc.set_page_setup(paper_size="A4", orientation="portrait", margins_mm={"left":15,"right":15,"top":15,"bottom":15},
                   header_margin_mm=0, footer_margin_mm=8)
F = "맑은 고딕"
st_title = doc.ensure_run_style(bold=True, size=15, color="#1B3A5C", font=F)
st_head  = doc.ensure_run_style(bold=True, size=10, color="#FFFFFF", font=F)
st_hsub  = doc.ensure_run_style(size=7.5, color="#FFFFFF", font=F)
st_area  = doc.ensure_run_style(bold=True, size=10, color="#FFFFFF", font=F)
st_no1   = doc.ensure_run_style(bold=True, size=7.5, color="#B8650D", font=F)
st_no    = doc.ensure_run_style(bold=True, size=7.5, color="#4D6680", font=F)
st_nm    = doc.ensure_run_style(bold=True, size=8.5, color="#1B3A5C", font=F)
st_ds    = doc.ensure_run_style(size=7.5, color="#4D6680", font=F)
st_none  = doc.ensure_run_style(size=8, color="#A8B8C8", font=F)
st_note  = doc.ensure_run_style(size=8, color="#4D6680", font=F)
st_noteb = doc.ensure_run_style(bold=True, size=8, color="#1B3A5C", font=F)

pp_item = doc.styles.ensure_paragraph_format(line_spacing_percent=120)
pp_ctr = doc.styles.ensure_paragraph_format(alignment="CENTER", line_spacing_percent=120)

def setp(p, runs, pp=None):
    if pp is not None:
        p.para_pr_id_ref = pp
    p.clear_text()
    for t, s in runs:
        p.add_run(t, char_pr_id_ref=s)

# ---- 간지 ----
st_c1 = doc.ensure_run_style(bold=True, size=30, color="#1B3A5C", font=F)
st_c2 = doc.ensure_run_style(bold=True, size=22, color="#1B3A5C", font=F)
st_c3 = doc.ensure_run_style(size=13, color="#4D6680", font=F)
st_c3b = doc.ensure_run_style(bold=True, size=13, color="#E58A1F", font=F)
pp_cover = doc.styles.ensure_paragraph_format(alignment="CENTER", line_spacing_percent=160)
first_para = doc.paragraphs[0]
setp(first_para, [("", st_c3)], pp_cover)
for _ in range(11):
    setp(doc.add_paragraph(""), [("", st_c3)], pp_cover)
setp(doc.add_paragraph(""), [("부  록", st_c1)], pp_cover)
setp(doc.add_paragraph(""), [("", st_c3)], pp_cover)
setp(doc.add_paragraph(""), [("통합돌봄 서비스 상세 내역", st_c2)], pp_cover)
for _ in range(3):
    setp(doc.add_paragraph(""), [("", st_c3)], pp_cover)
for t1, t2 in [("통합돌봄 서비스 한눈에 보기", ""), ("1. 지역특화서비스", "  12개"),
               ("2. 연계서비스(국가돌봄)", "  17개"), ("3. 연계서비스(도·지자체)", "  13개")]:
    setp(doc.add_paragraph(""), [(t1, st_c3), (t2, st_c3b)], pp_cover)

# ---- 한눈에 보기 ----
pp_break = doc.styles.ensure_paragraph_format(line_spacing_percent=120, break_setting={"pageBreakBefore": True})
p = doc.add_paragraph("")
p.para_pr_id_ref = pp_break
p.element.set("pageBreak", "1")   # 간지 다음 쪽에서 시작
p.section.mark_dirty() if hasattr(p.section, "mark_dirty") else None
p.clear_text()
for _t, _s in [("통합돌봄 서비스 한눈에 보기", st_title), ("   지역특화 12 · 국가돌봄 17 · 도·지자체 13 (총 42개)", st_note)]:
    p.add_run(_t, char_pr_id_ref=_s)

t = doc.add_table(6, 4, width=51000)
t.set_column_widths([8, 31, 31, 30])

def cell_lines(r, c, lines, pp=None):
    cell = t.cell(r, c)
    first = True
    for runs in lines:
        para = cell.paragraphs[0] if first else cell.add_paragraph("")
        setp(para, runs, pp if pp is not None else pp_item); first = False

t.set_cell_shading(0, 0, "#FFFFFF")
cell_lines(0, 0, [[("구분", doc.ensure_run_style(bold=True, size=10, color="#1B3A5C", font=F))]], pp_ctr)
for j, (h, sub) in enumerate(COLS, start=1):
    t.set_cell_shading(0, j, ["#E58A1F", "#24507D", "#5F7F9F"][j-1])
    cell_lines(0, j, [[(h, st_head)], [(sub, st_hsub)]], pp_ctr)
for i, a in enumerate(AREAS, start=1):
    t.set_cell_shading(i, 0, "#1B3A5C")
    cell_lines(i, 0, [[(a, st_area)]], pp_ctr)
    for j in range(3):
        items = D[(i-1, j)]
        t.set_cell_shading(i, j+1, "#FDF1E3" if j == 0 else "#EEF5FB")
        if not items:
            cell_lines(i, j+1, [[("해당 없음", st_none)]], pp_ctr); continue
        lines = []
        for no, nm, ds in items:
            lines.append([(no + "  ", st_no1 if j == 0 else st_no), (nm, st_nm)])
            lines.append([("          " + ds, st_ds)])
        cell_lines(i, j+1, lines)

p = doc.add_paragraph("")
setp(p, [("※ 지역특화서비스 한도 ", st_noteb), ("건강살핌 1인 60만원(방문진료·가정간호·방문재활 합산), 일상살핌 1인 200만원(생활·동행·식사 합산), 내집살핌 생애 200만원", st_note)])
p = doc.add_paragraph("")
setp(p, [("※ 재원 ", st_noteb), ("국 50% · 도 5% · 시 45%, 중위소득별 본인부담 차등   ", st_note), ("※ 연계 원칙 ", st_noteb), ("국가돌봄서비스 우선 연계, 공백은 지역특화서비스로 보완", st_note)])

doc.save_to_path(sys.argv[1])
print("saved", sys.argv[1])
