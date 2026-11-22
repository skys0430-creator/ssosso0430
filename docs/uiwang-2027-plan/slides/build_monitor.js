// 회의장 모니터 화면(1장): NODE_PATH=<node_modules> node build_monitor.js <out.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");
const M = path.join(__dirname, "..", "figures", "monitor");
const OUT = process.argv[2] || "monitor.pptx";
const F = "Pretendard ExtraBold", FB = "Pretendard";
const NAVY = "13277A", MAG = "C8157F", INK = "1F2A44";
const p = new pptxgen(); p.layout = "LAYOUT_WIDE"; p.title = "제2차 의왕시 통합지원협의체 회의";
const W = 13.333;
const s = p.addSlide();
s.addImage({ path: path.join(M, "bg_monitor2.png"), x: 0, y: 0, w: W, h: 7.5, altText: "" });
const T = (t, x, y, w, h, o) => s.addText(t, { x, y, w, h, isTextBox: true, margin: 0, align: "center", valign: "middle", ...o });
T("2027년도 통합지원 지역계획(안) 심의", 0.6, 1.75, W - 1.2, 0.85, { fontFace: F, fontSize: 40, color: NAVY });
T([{ text: "제2차 통합지원협의체 ", options: { color: NAVY } }, { text: "회의", options: { color: MAG } }], 0.6, 2.6, W - 1.2, 1.35, { fontFace: F, fontSize: 72 });
// 일시·장소 알약
const pw = 9.6, px = (W - pw) / 2, py = 4.35;
s.addShape(p.ShapeType.roundRect, { x: px, y: py, w: pw, h: 0.72, fill: { color: "FDE8F3" }, line: { type: "none" }, rectRadius: 0.36 });
const tag = (t, x) => { s.addShape(p.ShapeType.roundRect, { x, y: py + 0.12, w: 0.95, h: 0.48, fill: { color: MAG }, line: { type: "none" }, rectRadius: 0.24 }); T(t, x, py + 0.12, 0.95, 0.48, { fontFace: F, fontSize: 18, color: "FFFFFF" }); };
tag("일시", px + 0.25); T("2026. 11. 26.(목) 14:00", px + 1.35, py, 4.0, 0.72, { fontFace: FB, bold: true, fontSize: 21, color: INK, align: "left" });
s.addShape(p.ShapeType.line, { x: px + 5.4, y: py + 0.18, w: 0, h: 0.36, line: { color: "D9A3C4", width: 1.5 } });
tag("장소", px + 5.65); T("의왕시청 중회의실", px + 6.75, py, 2.75, 0.72, { fontFace: FB, bold: true, fontSize: 21, color: INK, align: "left" });
s.addImage({ path: path.join(M, "logo_uiwang.png"), x: (W - 2.3) / 2, y: 5.4, w: 2.3, h: 2.3 * 228 / 651, altText: "의왕시" });
s.addNotes("회의장 모니터 대기 화면입니다. ");
p.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT));
