// 표지·간지·닫는 슬라이드 레터링(투명 PNG) 생성: node lettering.js <출력폴더>
// 규칙: ① 성격에 맞는 글꼴 ② 대비되는 두 말은 글꼴을 다르게 ③ 강조는 낱말 하나에만
const {chromium}=require('playwright');
const path=require('path');
const OUT=process.argv[2];
const FONTS='<link href="https://fonts.googleapis.com/css2?family=Gothic+A1:wght@700;800;900&family=Gowun+Batang:wght@700&family=Gowun+Dodum&display=swap" rel="stylesheet">';
const BASE=`*{margin:0;padding:0;box-sizing:border-box}html,body{background:transparent}body{display:inline-block;padding:8px 14px 14px 4px;word-break:keep-all;white-space:nowrap}
.sans{font-family:'Gothic A1',sans-serif}.serif{font-family:'Gowun Batang',serif;font-weight:700}.dodum{font-family:'Gowun Dodum',sans-serif}
.hl{position:relative;padding:0 3px}.hl:before{content:"";position:absolute;left:-2px;right:-2px;bottom:3px;height:38%;border-radius:10px;z-index:-1;transform:rotate(-1deg)}`;
const pages={
 cover:`<div class="sans" style="font-weight:900;font-size:80px;letter-spacing:-2.5px;color:#0B1220;line-height:1.05">2027년도 통합지원 지역계획</div>
  <div style="display:flex;align-items:center;gap:16px;margin-top:22px;font-size:42px;line-height:1.1">
   <svg width="48" height="48" viewBox="0 0 48 48" fill="none" stroke="#2563EB" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M6 22 24 7l18 15"/><path d="M11 18v22h26V18"/><path d="M24 34s-7-4.3-7-9a3.8 3.8 0 0 1 7-2.2A3.8 3.8 0 0 1 31 25c0 4.7-7 9-7 9z" fill="#BFDBFE"/></svg>
   <span class="serif" style="color:#1E293B">살던 곳에서 <span class="hl" style="color:#1D4ED8">건강하게</span>,</span>
   <span class="dodum" style="color:#475569">함께 살피는 의왕</span></div>
  <style>.hl:before{background:#BFDBFE}</style>`,
 closing:`<div style="text-align:center"><div class="sans" style="font-weight:800;font-size:44px;letter-spacing:-1.2px;color:#334155">2027년도 통합지원 지역계획(안)을</div>
  <div class="serif" style="font-size:70px;color:#0B1220;margin-top:10px;letter-spacing:-1px"><span class="hl" style="color:#057555">심의</span>해 주십시오</div></div>
  <style>.hl:before{background:#BBF7D0}</style>`,
};
const sec=[["01","2026","성과","","대상자·서비스 연계 실적 · 운영 기반 구축","#2159D4","#BFDBFE"],
 ["02","지역","진단","","돌봄 수요 · 서비스 대응 수준 · 핵심 문제와 추진방향","#0277B3","#BAE6FD"],
 ["03","","목표","·체계","비전과 2027년 목표 · 추진과업 및 핵심사업","#7034D5","#DDD6FE"],
 ["04","세부","추진","","대상자 발굴 · 퇴원환자 · 모니터링 · 서비스 · 운영 · 교육·홍보","#A24B08","#FED7AA"],
 ["05","지표·","예산","","성과지표 · 2027년 예산","#04875E","#BBF7D0"]];
sec.forEach(([no,pre,key,post,items,col,tint])=>{
 pages['div'+no]=`<div style="display:flex;align-items:flex-end;gap:30px;padding-top:10px">
   <div class="sans" style="font-weight:900;font-size:150px;line-height:1;padding-top:6px;letter-spacing:-6px;color:${col};opacity:.22">${no}</div>
   <div class="serif" style="font-size:78px;line-height:1;color:#0B1220;letter-spacing:-1.5px;padding-bottom:4px">${pre?(pre.endsWith('·')?pre:pre+' '):''}<span class="hl" style="color:${col}">${key}</span>${post}</div></div>
  <div class="dodum" style="font-size:28px;color:#475569;margin-top:26px">${items}</div>
  <style>.hl:before{background:${tint}}</style>`;
});
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1800,height:600},deviceScaleFactor:3});
 for(const [k,html] of Object.entries(pages)){
  await p.setContent(`<!doctype html><html lang="ko"><head><meta charset="utf-8">${FONTS}<style>${BASE}</style></head><body>${html}</body></html>`,{waitUntil:'networkidle'});
  await p.evaluate(()=>document.fonts.ready);await p.waitForTimeout(600);
  const el=await p.$('body');await el.screenshot({path:path.join(OUT,`lettering_${k}.png`),omitBackground:true});console.log(k);}
 await b.close();})();
