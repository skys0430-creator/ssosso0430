const {chromium}=require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1500,height:900},deviceScaleFactor:2});
await p.goto('file://'+process.argv[2],{waitUntil:'networkidle'}).catch(()=>{});
await p.waitForTimeout(1500);
const f=await p.evaluate(()=>document.fonts.check('16px "Noto Sans KR"'));console.log('noto',f);
const el=await p.$('body');await el.screenshot({path:process.argv[3]});await b.close();})();
