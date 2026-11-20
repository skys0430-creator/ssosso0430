// 투명 배경 캡처: node shot_tr.js <html> <png> [폭]
const {chromium}=require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:+(process.argv[4]||1300),height:400},deviceScaleFactor:3});
await p.goto('file://'+process.argv[2],{waitUntil:'networkidle'}).catch(()=>{});
await p.evaluate(()=>document.fonts.ready);await p.waitForTimeout(1200);
console.log('fonts',await p.evaluate(()=>[document.fonts.check('80px "Gothic A1"'),document.fonts.check('40px "Gowun Batang"')]));
const el=await p.$('body');await el.screenshot({path:process.argv[3],omitBackground:true});await b.close();})();
