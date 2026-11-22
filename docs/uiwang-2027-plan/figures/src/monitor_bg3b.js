// 모니터 안내 화면 배경: v3 bg_mid 구성 + 마젠타 구슬·리본(띠 없음/남색 띠는 PPT에서 도형으로)
const {chromium}=require('playwright');const path=require('path');const OUT=process.argv[2];
const S=(x,y,d,blur,op,c1,c2,c3)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;opacity:${op};filter:blur(${blur}px);background:radial-gradient(circle at 32% 28%,#fff 0%,${c1} 18%,${c2} 48%,${c3} 78%,${c3} 100%);box-shadow:0 ${d/8}px ${d/4}px rgba(80,60,180,.25)"></div>`;
const blue=(x,y,d,b=0,o=1)=>S(x,y,d,b,o,'#dfe9ff','#9fb9ff','#3a55d6'), mag=(x,y,d,b=0,o=1)=>S(x,y,d,b,o,'#ffe0f1','#f590c8','#c8157f');
const ring=(x,y,d,w,c,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;border:${w}px solid rgba(${c},${op})"></div>`;
const rib=(x,y,w,h,rot,c1,c2,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;border-radius:${h}px;transform:rotate(${rot}deg);opacity:${op};background:linear-gradient(90deg,rgba(255,255,255,0) 0%,${c1} 35%,${c2} 60%,rgba(255,255,255,0) 100%);filter:blur(2px)"></div>`;
const html=`<div style="position:absolute;inset:0;background:linear-gradient(160deg,#F7F4FF 0%,#ECEBFF 55%,#E6E2FF 100%)"></div>
 ${ring(-300,520,900,60,'120,140,255',.08)}${ring(1400,-300,800,50,'200,21,127',.07)}
 ${rib(1250,820,900,130,-18,'rgba(245,170,215,.9)','rgba(200,21,127,.6)',.75)}
 ${mag(1660,760,180)}${blue(1560,930,60,1)}${mag(1180,40,46,0,.9)}${blue(1250,110,24,1,.8)}`;
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1920,height:1080}});
 await p.setContent(`<html><body style="margin:0;width:1920px;height:1080px;position:relative;overflow:hidden">${html}</body></html>`);await p.waitForTimeout(200);
 await p.screenshot({path:OUT});await b.close();console.log('ok');})();
