// 모니터 화면 배경 v2(심플 파스텔 소용돌이, 첨부 썸네일 분위기 재현): node monitor_bg2.js <out.png>
const {chromium}=require('playwright');
const arc=(cx,cy,r,w,col,op,rot,blur)=>`<div style="position:absolute;left:${cx-r}px;top:${cy-r}px;width:${2*r}px;height:${2*r}px;border-radius:50%;
 border:${w}px solid transparent;border-bottom-color:${col};border-right-color:${col};opacity:${op};transform:rotate(${rot}deg);filter:blur(${blur}px)"></div>`;
const cols=['#FFB0CF','#FFC9A3','#FFE6A6','#E77BC9','#B9C8FF'];
let a='';
cols.forEach((c,i)=>{ a+=arc(-200,-220,640+i*85,60,c,.55,-6+i*4,22); });
let b='';
['#FFC9A3','#E77BC9','#B9C8FF'].forEach((c,i)=>{ b+=arc(2250,1400,640+i*85,50,c,.35,180+i*4,22); });
const html=`<div style="position:absolute;inset:0;background:#FFFFFF"></div>
 <div style="position:absolute;inset:0;background:radial-gradient(ellipse at 0% 0%,rgba(255,214,235,.85) 0%,rgba(255,235,220,.45) 28%,rgba(255,255,255,0) 55%)"></div>
 ${a}${b}
 <div style="position:absolute;inset:0;background:radial-gradient(ellipse at 55% 55%,rgba(255,255,255,.95) 0%,rgba(255,255,255,.7) 38%,rgba(255,255,255,0) 75%)"></div>`;
(async()=>{const br=await chromium.launch();const p=await br.newPage({viewport:{width:1920,height:1080}});
 await p.setContent(`<html><body style="margin:0;width:1920px;height:1080px;position:relative;overflow:hidden">${html}</body></html>`);await p.waitForTimeout(200);
 await p.screenshot({path:process.argv[2]});await br.close();console.log('bg2');})();
