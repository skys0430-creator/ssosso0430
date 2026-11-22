// 회의장 모니터 대기 화면(v3 본문 틀): NODE_PATH=<node_modules> node build_monitor_v3.js <out.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react"), RDS = require("react-dom/server"), sharp = require("sharp"), L = require("react-icons/lu");
const FIG = path.join(__dirname, "..", "figures");
const OUT = process.argv[2] || "monitor_v3.pptx";
const F = "Pretendard", FT = "Pretendard ExtraBold";
const NAVY = "13277A", INK = "111B3D", GRAY = "5B6785", BLUE = "2F5BEA", PALE = "DCE6FF", BAND = "0E1F5B", MAG = "C8157F";
const W = 13.333, MX = 0.6, CW = W - 2 * MX;
async function icon(n, c) { const svg = RDS.renderToStaticMarkup(React.createElement(L[n], { color: "#" + c, size: "256" })); return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64"); }

(async () => {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE"; p.title = "제2차 의왕시 통합지원협의체 회의";
  const s = p.addSlide();
  const T = (t, x, y, w, h, o = {}) => s.addText(t, { x, y, w, h, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, color: INK, valign: "top", ...o });
  const tab = (t, y) => { s.addShape(p.ShapeType.rect, { x: MX, y, w: 3.4, h: 0.4, fill: { color: NAVY }, line: { type: "none" } }); T(t, MX + 0.18, y, 3.2, 0.4, { fontSize: 15, bold: true, color: "FFFFFF", valign: "middle" }); s.addShape(p.ShapeType.line, { x: MX, y: y + 0.4, w: CW, h: 0, line: { color: NAVY, width: 1.5 } }); };
  const ic = async (n, x, y, d, fg, bg) => { s.addShape(p.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } }); s.addImage({ data: await icon(n, fg), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5 }); };

  s.addImage({ path: path.join(FIG, "v3", "bg_content.png"), x: 0, y: 0, w: W, h: 7.5, altText: "" });
  T([{ text: "제2차  ", options: { bold: true, color: MAG } }, { text: "—  의왕시 통합지원협의체 회의", options: { color: GRAY } }], MX, 0.5, 10, 0.35, { fontSize: 15 });
  T([{ text: "2027년도 통합지원 지역계획(안)을\n", options: { color: INK } }, { text: "심의", options: { color: MAG } }, { text: "하겠습니다.", options: { color: INK } }],
    MX, 0.92, CW, 1.6, { fontFace: FT, fontSize: 40, lineSpacingMultiple: 1.05 });

  tab("회의 개요", 2.85);
  const info = [["LuCalendarCheck", "일시", "2026. 11. 26.(목) 14:00"], ["LuMapPin", "장소", "의왕시청 중회의실"]];
  const cw = (CW - 0.3) / 2;
  for (let i = 0; i < 2; i++) {
    const x = MX + i * (cw + 0.3), y = 3.45;
    s.addShape(p.ShapeType.roundRect, { x, y, w: cw, h: 1.15, fill: { color: "FFFFFF" }, line: { color: PALE, width: 1 }, rectRadius: 0.1 });
    await ic(info[i][0], x + 0.3, y + 0.2, 0.75, i === 0 ? "FFFFFF" : BLUE, i === 0 ? MAG : PALE);
    T(info[i][1], x + 1.3, y + 0.17, 2, 0.35, { fontSize: 15, bold: true, color: GRAY });
    T(info[i][2], x + 1.3, y + 0.5, cw - 1.5, 0.5, { fontFace: FT, fontSize: 26, color: INK, valign: "middle" });
  }

  tab("안건", 4.95);
  const ag = [["보고", "2026년 통합돌봄 추진사항", BLUE], ["심의", "2027년도 의왕시 통합지원 지역계획(안)", MAG]];
  for (let i = 0; i < 2; i++) {
    const x = MX + i * (cw + 0.3), y = 5.55;
    s.addShape(p.ShapeType.roundRect, { x, y, w: cw, h: 0.82, fill: { color: i ? NAVY : "FFFFFF" }, line: i ? { type: "none" } : { color: PALE, width: 1 }, rectRadius: 0.1 });
    s.addShape(p.ShapeType.roundRect, { x: x + 0.25, y: y + 0.24, w: 0.8, h: 0.34, fill: { color: ag[i][2] }, line: { type: "none" }, rectRadius: 0.17 });
    T(ag[i][0], x + 0.25, y + 0.24, 0.8, 0.34, { fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    T(ag[i][1], x + 1.25, y, cw - 1.4, 0.82, { fontSize: 18, bold: true, color: i ? "FFFFFF" : INK, valign: "middle" });
  }

  s.addShape(p.ShapeType.rect, { x: 0, y: 6.82, w: W, h: 0.68, fill: { color: BAND }, line: { type: "none" } });
  T("살던 곳에서 건강하게, 함께 살피는 의왕", MX, 6.82, 8, 0.68, { fontSize: 15, bold: true, color: "FFFFFF", valign: "middle" });
  T("의왕시 복지정책과 통합돌봄팀", W - MX - 4, 6.82, 4, 0.68, { fontSize: 13, color: "C9D7FF", align: "right", valign: "middle" });
  s.addNotes("회의장 모니터 대기 화면(v3 발표자료 틀).");
  await p.writeFile({ fileName: OUT }); console.log("wrote", OUT);
})();
