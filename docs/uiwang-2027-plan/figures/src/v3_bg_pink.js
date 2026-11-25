// v3 배경 연분홍판: node v3_bg_pink.js <출력폴더>
const {chromium}=require('playwright');const path=require('path');const OUT=process.argv[2];
const sphere=(x,y,d,blur=0,op=1)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;opacity:${op};filter:blur(${blur}px);
 background:radial-gradient(circle at 32% 28%,#ffffff 0%,#fff0f7 18%,#f9c6df 48%,#ee9cc6 78%,#e283b5 100%);box-shadow:0 ${d/8}px ${d/4}px rgba(220,110,170,.22)"></div>`;
const msph=(x,y,d,blur=0,op=1)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;opacity:${op};filter:blur(${blur}px);
 background:radial-gradient(circle at 32% 28%,#ffffff 0%,#ffe0f1 18%,#f590c8 48%,#d9369a 78%,#c8157f 100%);box-shadow:0 ${d/8}px ${d/4}px rgba(200,21,127,.22)"></div>`;
const ring=(x,y,d,w,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${d}px;height:${d}px;border-radius:50%;border:${w}px solid rgba(235,140,190,${op})"></div>`;
const ribbon=(x,y,w,h,rot,op)=>`<div style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;border-radius:${h}px;transform:rotate(${rot}deg);opacity:${op};
 background:linear-gradient(90deg,rgba(255,255,255,0) 0%,rgba(255,214,234,.9) 35%,rgba(242,160,200,.65) 60%,rgba(255,255,255,0) 100%);filter:blur(2px)"></div>`;
const bgs={
 bg_cover:`<div style="position:absolute;inset:0;background:linear-gradient(120deg,#FFF6FA 0%,#FDEAF3 55%,#FADCEB 100%)"></div>
  ${ring(1020,-260,1100,90,.12)}${ring(1180,120,760,40,.16)}${ribbon(880,560,1300,170,-24,.85)}${ribbon(1000,180,1100,110,-36,.6)}
  ${msph(1460,180,170)}${sphere(1290,470,90)}${msph(1650,640,240,0,.95)}${sphere(1180,860,60,1)}${sphere(1730,90,50,1)}${sphere(980,740,36,2,.8)}`,
 bg_content:`<div style="position:absolute;inset:0;background:linear-gradient(180deg,#FDEEF5 0%,#FFF6FA 45%,#FFFBFD 100%)"></div>
  ${ring(1500,-380,900,50,.08)}${ribbon(-160,10,520,70,-20,.55)}${msph(-24,-20,74,0,.9)}${sphere(40,40,30,1,.9)}`,
 bg_mid:`<div style="position:absolute;inset:0;background:linear-gradient(160deg,#FFF4F9 0%,#FDE7F1 60%,#FAD9E9 100%)"></div>
  ${ring(-300,520,900,60,.10)}${ring(1400,-300,800,50,.10)}${ribbon(1250,820,900,130,-18,.7)}${msph(1660,760,180)}${sphere(1560,930,60,1)}${msph(90,120,70,0,.9)}${sphere(40,60,30,1,.8)}`,
};
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1920,height:1080},deviceScaleFactor:1});
 for(const [k,h] of Object.entries(bgs)){await p.setContent(`<html><body style="margin:0;width:1920px;height:1080px;position:relative;overflow:hidden">${h}</body></html>`);await p.waitForTimeout(200);await p.screenshot({path:path.join(OUT,k+'.png')});console.log(k);}
 await b.close();})();
