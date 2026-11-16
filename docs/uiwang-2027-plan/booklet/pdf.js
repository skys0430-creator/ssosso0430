const {chromium}=require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:794,height:1123}});
p.on('console',m=>console.log('LOG',m.text()));p.on('pageerror',e=>console.log('ERR',e.message));
await p.goto('file://'+process.argv[2],{waitUntil:'networkidle'}).catch(e=>console.log(e.message));
await p.waitForFunction(()=>window.PAGES,null,{timeout:30000});
console.log('noto',await p.evaluate(()=>document.fonts.check('16px "Noto Sans KR"')),'pages',await p.evaluate(()=>window.PAGES));
await p.pdf({path:process.argv[3],preferCSSPageSize:true,printBackground:true});
await b.close();})();
