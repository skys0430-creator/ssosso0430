// v3(레퍼런스형) 배경·레터링 생성: node v3_assets.js <출력폴더>
const {chromium}=require('playwright');const path=require('path');const OUT=process.argv[2];
const FONTS='<link href="https://fonts.googleapis.com/css2?family=Gothic+A1:wght@700;800;900&family=Gowun+Batang:wght@700&display=swap" rel="stylesheet">';
const sphere=(x,y,d,blur=0,op=1)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;opacity:${op};filter:blur(${blur}px);
 background:radial-gradient(circle at 32% 28%,#ffffff 0%,#dfe9ff 18%,#9fb9ff 48%,#5b7cf0 78%,#3a55d6 100%);box-shadow:0 ${d/8}px ${d/4}px rgba(60,90,200,.25)"></div>`;
const ring=(x,y,d,w,op,rot=0)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;border:${w}px solid rgba(120,150,255,${op});transform:rotate(${rot}deg)"></div>`;
const ribbon=(x,y,w,h,rot,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;border-radius:${h}px;transform:rotate(${rot}deg);opacity:${op};
 background:linear-gradient(90deg,rgba(255,255,255,0) 0%,rgba(190,210,255,.9) 35%,rgba(120,150,255,.7) 60%,rgba(255,255,255,0) 100%);filter:blur(2px)"></div>`;
const bgs={
 bg_cover:`<div style="position:absolute;inset:0;background:linear-gradient(120deg,#F4F8FF 0%,#E6EEFF 55%,#D6E3FF 100%)"></div>
  ${ring(1020,-260,1100,90,.10)}${ring(1180,120,760,40,.14)}${ribbon(880,560,1300,170,-24,.85)}${ribbon(1000,180,1100,110,-36,.6)}
  ${sphere(1460,180,170)}${sphere(1290,470,90)}${sphere(1650,640,240,0,.95)}${sphere(1180,860,60,1)}${sphere(1730,90,50,1)}${sphere(980,740,36,2,.8)}`,
 bg_content:`<div style="position:absolute;inset:0;background:linear-gradient(180deg,#EEF3FF 0%,#F6F9FF 45%,#FFFFFF 100%)"></div>
  ${ring(1500,-380,900,50,.06)}${ribbon(-160,10,520,70,-20,.55)}${sphere(-24,-20,74,0,.9)}${sphere(40,40,30,1,.9)}`,
 bg_mid:`<div style="position:absolute;inset:0;background:linear-gradient(160deg,#F2F6FF 0%,#E4ECFF 60%,#D3E0FF 100%)"></div>
  ${ring(-300,520,900,60,.08)}${ring(1400,-300,800,50,.08)}${ribbon(1250,820,900,130,-18,.7)}${sphere(1660,760,180)}${sphere(1560,930,60,1)}${sphere(90,120,70,0,.9)}${sphere(40,60,30,1,.8)}`,
};
const SW=(c)=>`<svg class="sw" viewBox="0 0 200 24" preserveAspectRatio="none"><path d="M3 15 C 45 5, 120 3, 197 11" stroke="${c}" stroke-width="7" fill="none" stroke-linecap="round" opacity=".45"/></svg>`;
const HH=`<svg width="74" height="74" viewBox="0 0 48 48" fill="none" stroke="#13277A" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" style="flex:0 0 74px"><path d="M6 22 24 7l18 15"/><path d="M11 18v22h26V18"/><path d="M24 34s-7-4.3-7-9a3.8 3.8 0 0 1 7-2.2A3.8 3.8 0 0 1 31 25c0 4.7-7 9-7 9z" fill="#C9D7FF"/></svg>`;
const EM=(w,c='#2F5BEA')=>`<span class="em" style="color:${c}">${w}${SW(c)}</span>`;
const letters={
 // 표지: 성격(공공 계획의 단단함=굵은 고딕 / 돌봄의 따뜻함=명조), 대비되는 두 말=두 글꼴, 강조=‘건강하게’ 하나
 L_cover:`<div class="sans" style="font-weight:900;font-size:100px;letter-spacing:-3px;color:#13277A;line-height:1.12">살던 곳에서 ${EM('건강하게')},</div>
  <div style="display:flex;align-items:center;gap:22px;margin-top:14px">${HH}<span class="serif" style="font-size:94px;letter-spacing:-2px;color:#13277A;line-height:1.15">함께 살피는 의왕</span></div>`,
 // 2쪽(단원 시작): ‘법 시행 첫해’(명조) ↔ ‘먼저 찾고 연결한 2026년’(고딕), 강조=‘연결’
 L_s2:`<div style="text-align:center"><div class="serif" style="font-size:40px;color:#5B6785;letter-spacing:-.5px">법 시행 첫해,</div>
  <div class="sans" style="font-weight:900;font-size:76px;letter-spacing:-2.5px;color:#111B3D;margin-top:4px">먼저 찾고 ${EM('연결')}한 2026년</div></div>`,
 // 4쪽(단원 시작): 비전(명조) ↔ 목표(고딕), 강조=‘목표’
 L_s4:`<div style="text-align:center"><div class="serif" style="font-size:42px;color:#5B6785;letter-spacing:-.5px">살던 곳에서 건강하게, 함께 살피는 의왕</div>
  <div class="sans" style="font-weight:900;font-size:78px;letter-spacing:-2.5px;color:#111B3D;margin-top:6px">의왕을 위한 2027년 3대 ${EM('목표')}</div></div>`,
 // 닫는 슬라이드: 요청 대상(고딕) ↔ 요청(명조), 강조=‘심의’
 L_closing:`<div style="text-align:center"><div class="sans" style="font-weight:800;font-size:50px;letter-spacing:-1.5px;color:#13277A">2027년도 통합지원 지역계획(안)을</div>
  <div class="serif" style="font-size:88px;color:#111B3D;margin-top:8px;letter-spacing:-1.5px">${EM('심의')}해 주십시오</div></div>`,
};
(async()=>{const b=await chromium.launch();
 let p=await b.newPage({viewport:{width:1920,height:1080},deviceScaleFactor:1});
 for(const [k,h] of Object.entries(bgs)){await p.setContent(`<html><body style="margin:0;width:1920px;height:1080px;position:relative;overflow:hidden">${h}</body></html>`);await p.waitForTimeout(200);await p.screenshot({path:path.join(OUT,k+'.png')});console.log(k);}
 p=await b.newPage({viewport:{width:1800,height:600},deviceScaleFactor:3});
 for(const [k,h] of Object.entries(letters)){await p.setContent(`<!doctype html><html lang="ko"><head><meta charset="utf-8">${FONTS}<style>*{margin:0;padding:0}html,body{background:transparent}body{display:inline-block;padding:10px 16px 14px 4px;white-space:nowrap}.sans{font-family:'Gothic A1',sans-serif}.serif{font-family:'Gowun Batang',serif;font-weight:700}.em{position:relative;display:inline-block}.em .sw{position:absolute;left:-2%;bottom:-.08em;width:104%;height:.3em}</style></head><body>${h}</body></html>`,{waitUntil:'networkidle'});
  await p.evaluate(()=>document.fonts.ready);await p.waitForTimeout(500);const el=await p.$('body');await el.screenshot({path:path.join(OUT,k+'.png'),omitBackground:true});console.log(k);}
 await b.close();})();
