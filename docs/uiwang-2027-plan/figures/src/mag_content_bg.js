// 협의체 모니터(마젠타) 화면과 같은 계열의 본문 배경: node mag_content_bg.js <out.png>
const {chromium}=require('playwright');const OUT=process.argv[2];
const S=(x,y,d,blur,op,c1,c2,c3)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;opacity:${op};filter:blur(${blur}px);background:radial-gradient(circle at 32% 28%,#fff 0%,${c1} 18%,${c2} 48%,${c3} 78%,${c3} 100%);box-shadow:0 ${d/8}px ${d/4}px rgba(80,60,180,.22)"></div>`;
const blue=(x,y,d,b=0,o=1)=>S(x,y,d,b,o,'#dfe9ff','#9fb9ff','#3a55d6'), mag=(x,y,d,b=0,o=1)=>S(x,y,d,b,o,'#ffe0f1','#f590c8','#c8157f');
const ring=(x,y,d,w,c,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;border:${w}px solid rgba(${c},${op})"></div>`;
const html=`<div style="position:absolute;inset:0;background:linear-gradient(170deg,#F7F4FF 0%,#F3F1FF 40%,#FAF9FF 100%)"></div>
 ${ring(1500,-380,900,50,'200,21,127',.06)}${mag(-24,-20,74,0,.9)}${blue(40,40,30,1,.9)}`;
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1920,height:1080}});
 await p.setContent(`<html><body style="margin:0;width:1920px;height:1080px;position:relative;overflow:hidden">${html}</body></html>`);await p.waitForTimeout(200);
 await p.screenshot({path:OUT});await b.close();console.log('ok');})();
