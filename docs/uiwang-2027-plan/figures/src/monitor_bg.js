// 회의장 모니터 화면 배경(v3 유리구슬 + 마젠타 강조): node monitor_bg.js <out.png>
const {chromium}=require('playwright');
const sp=(x,y,d,c1,c2,c3,op=1,blur=0)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;opacity:${op};filter:blur(${blur}px);background:radial-gradient(circle at 32% 28%,#fff 0%,${c1} 20%,${c2} 55%,${c3} 100%);box-shadow:0 ${d/8}px ${d/4}px rgba(80,60,160,.22)"></div>`;
const B=(x,y,d,o=1,b=0)=>sp(x,y,d,'#dfe9ff','#9fb9ff','#3a55d6',o,b), M=(x,y,d,o=1,b=0)=>sp(x,y,d,'#ffd9ef','#f37cc0','#c8157f',o,b);
const ring=(x,y,d,w,c,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;border:${w}px solid rgba(${c},${op})"></div>`;
const rib=(x,y,w,h,rot,c1,c2,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;border-radius:${h}px;transform:rotate(${rot}deg);opacity:${op};background:linear-gradient(90deg,rgba(255,255,255,0) 0%,${c1} 35%,${c2} 65%,rgba(255,255,255,0) 100%);filter:blur(2px)"></div>`;
const html=`<div style="position:absolute;inset:0;background:linear-gradient(160deg,#FBF6FF 0%,#EEF2FF 45%,#E3EBFF 100%)"></div>
 ${ring(-320,-360,820,70,'200,21,127',.07)}${ring(1380,-300,860,70,'90,120,255',.08)}${ring(-260,620,700,50,'90,120,255',.07)}
 ${rib(-200,90,760,120,-24,'rgba(243,124,192,.75)','rgba(200,21,127,.55)',.8)}${rib(1250,40,900,110,-22,'rgba(190,210,255,.9)','rgba(120,150,255,.7)',.8)}
 ${rib(-150,900,1100,150,-8,'rgba(243,124,192,.7)','rgba(200,21,127,.5)',.85)}${rib(900,880,1200,160,-10,'rgba(190,210,255,.9)','rgba(110,140,250,.75)',.9)}
 ${M(120,150,120)}${M(60,300,46,.9,1)}${B(260,70,38,.9,1)}
 ${B(1650,120,150)}${M(1700,330,56,.95)}${B(1780,360,34,.9,1)}
 ${B(140,780,90,.95)}${M(1720,740,170)}${B(1580,890,60,.9,1)}${M(330,930,40,.85,1)}`;
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1920,height:1080}});
 await p.setContent(`<html><body style="margin:0;width:1920px;height:1080px;position:relative;overflow:hidden">${html}</body></html>`);await p.waitForTimeout(200);
 await p.screenshot({path:process.argv[2]});await b.close();console.log('bg');})();
