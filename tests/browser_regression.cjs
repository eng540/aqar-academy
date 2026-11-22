/* Run with NODE_PATH pointing to a local playwright installation. No account/site writes. */
const {chromium}=require('playwright');
const {spawn}=require('child_process');
const fs=require('fs'),path=require('path');
const ROOT=path.resolve(__dirname,'..'), OUT=process.env.AQAR_EVIDENCE||'/tmp/aqar-evidence';
fs.mkdirSync(OUT,{recursive:true});
const catalog=JSON.parse(fs.readFileSync(path.join(ROOT,'assets/catalog.json'),'utf8'));
const routeDocuments=catalog.files.filter(p=>p.endsWith('.md')&&!p.startsWith('assets/vendor/')).map(p=>({path:p}));
const server=spawn('python3',['-m','http.server','8765','--bind','127.0.0.1','--directory',ROOT],{stdio:['ignore','pipe','pipe']});
const base='http://127.0.0.1:8765/';
const checks=[], screenshots=[];let browser;
function check(name,condition,details=''){checks.push({name,pass:!!condition,details});if(!condition)console.error('FAIL',name,details);}
async function ready(page){await page.waitForFunction(()=>document.querySelector('#content')?.getAttribute('aria-busy')==='false'&&document.querySelector('#content')?.dataset.route===location.hash);}
async function route(page,p){await page.goto(base+'#/'+p);await ready(page);}
(async()=>{await new Promise(r=>setTimeout(r,700));browser=await chromium.launch({executablePath:'/usr/bin/google-chrome',headless:true,args:['--no-sandbox']});
 const errors=[];const context=await browser.newContext();const page=await context.newPage();page.on('pageerror',e=>errors.push(e.message));
 for(const [label,width,height] of [['desktop',1366,960],['mobile',390,844]]){
  await page.setViewportSize({width,height});await page.waitForTimeout(300);
  for(let i=0;i<routeDocuments.length;i++){
   const d=routeDocuments[i];await route(page,d.path);
   const data=await page.evaluate(()=>({error:!!document.querySelector('#content .err'),h1:!!document.querySelector('#content h1'),width:document.documentElement.scrollWidth,viewport:innerWidth}));
   check(label+': route '+d.path,!data.error&&data.h1&&data.width<=data.viewport+1,JSON.stringify(data));
   const file=path.join(OUT,`${label}-${String(i+1).padStart(2,'0')}.png`);await page.screenshot({path:file});const fullFile=path.join(OUT,`full-${label}-${String(i+1).padStart(2,'0')}.png`);await page.screenshot({path:fullFile,fullPage:true});screenshots.push({label,path:d.path,file,fullFile});
  }
 }
 await page.setViewportSize({width:1366,height:960});
 await route(page,'06-دليل-الدراسة-الكاملة.md');
 await page.locator('#search').fill('XIRR');await page.waitForFunction(()=>document.querySelectorAll('.search-hit').length>0);
 check('Full-text search finds body term',await page.locator('.search-hit').count()>2);
 await page.screenshot({path:path.join(OUT,'search-desktop.png')});
 await page.locator('#search').fill('');
 const box=page.locator('.item input[type=checkbox]').first();await box.check();await page.reload();await ready(page);check('Reading progress survives reload',await box.isChecked());
 const [download]=await Promise.all([page.waitForEvent('download'),page.locator('#exportProgress').click()]);const downloaded=path.join(OUT,'reading-progress.json');await download.saveAs(downloaded);check('Progress export has reviewed schema',JSON.parse(fs.readFileSync(downloaded)).version===1);
 page.on('dialog',d=>d.accept());await page.locator('#resetBtn').click();check('Reset clears progress',!(await box.isChecked()));
 await page.locator('#importProgress').setInputFiles(downloaded);await page.waitForTimeout(150);check('Progress import restores reading mark',await box.isChecked());
 await page.locator('#importProgress').setInputFiles({name:'bad.json',mimeType:'application/json',buffer:Buffer.from('{not json}')});check('Invalid import preserves existing progress',await box.isChecked());
 await route(page,'التطبيقات-المحلولة/مختبر-التطبيق-العملي.md');
 const csv=page.locator('#content a').filter({hasText:'عينة العقارات'});const csvHref=await csv.getAttribute('href');check('Nested CSV points to its actual directory',csvHref.includes(encodeURIComponent('التطبيقات-المحلولة'))&&csvHref.endsWith('.csv'));
 const nested=page.locator('#content a').filter({hasText:'المحاكاة'}).first();await nested.click();await ready(page);check('Nested Markdown link routes to intended document',decodeURIComponent(page.url()).includes('الحالات-التدريبية/محاكاة-التطوير-أرض-1500م.md'));
 await route(page,'التطبيقات-المحلولة/مختبر-التطبيق-العملي.md');
 const financeAnchor=await page.locator('#content h2').nth(3).getAttribute('id');
 await page.goto(base+'#/التطبيقات-المحلولة/مختبر-التطبيق-العملي.md#'+encodeURIComponent(financeAnchor));await ready(page);
 check('Heading anchor scrolls inside document',await page.evaluate(()=>scrollY)>100);await page.screenshot({path:path.join(OUT,'worked-finance-desktop.png')});
 for(const [name,selector] of [['answer-key','التقييم/دليل-الإجابات-والتصحيح.md'],['financial-stage','الدورة-الكاملة/المرحلة-05-التطوير-العقاري.md'],['cashflow-guide','نماذج/دليل-النموذج-المالي.md']]){
  await route(page,selector);const tables=page.locator('#content table');if(await tables.count())await tables.first().scrollIntoViewIfNeeded();await page.screenshot({path:path.join(OUT,name+'-detail.png')});
 }
 await route(page,'الدورة-الكاملة/المرحلة-06-التقييم-والاستثمار.md');await page.locator('#content tr').filter({hasText:'654,700'}).scrollIntoViewIfNeeded();await page.screenshot({path:path.join(OUT,'sensitivity-detail.png')});
 await page.goto(base+'media-platform/');await page.waitForSelector('.card');await page.screenshot({path:path.join(OUT,'media-dashboard-desktop.png')});check('Every media document card routes through reader, CSV stays downloadable',(await page.locator('a.card').evaluateAll(as=>as.every(a=>a.getAttribute('href').includes('../index.html#/media-platform/')||a.getAttribute('href').endsWith('.csv')))));
 await page.setViewportSize({width:390,height:844});await page.screenshot({path:path.join(OUT,'media-dashboard-mobile.png')});
 await route(page,'03-الجدول-الزمني.md');
 const tableMotion=await page.locator('.tw').evaluateAll(elements=>{const el=elements.find(x=>x.scrollWidth>x.clientWidth);if(!el)return {overflow:false,moved:false};el.scrollLeft=-180;return {overflow:true,moved:el.scrollLeft<0};});
 check('Mobile wide table actually scrolls',tableMotion.overflow&&tableMotion.moved,JSON.stringify(tableMotion));
 await page.screenshot({path:path.join(OUT,'mobile-table-scrolled.png')});
 await route(page,'06-دليل-الدراسة-الكاملة.md');await page.locator('#menuBtn').click();check('Mobile menu opens with accessible state',await page.locator('#menuBtn').getAttribute('aria-expanded')==='true');await page.screenshot({path:path.join(OUT,'mobile-menu.png')});await page.keyboard.press('Escape');check('Escape closes menu',await page.locator('#menuBtn').getAttribute('aria-expanded')==='false');
 await page.goto(base+'#/../../etc/passwd');await ready(page);check('Traversal route rejected',await page.locator('#content .err').count()===1);
 await page.goto(base+'#/%E0%A4%A');await ready(page);check('Malformed route rejected',await page.locator('#content .err').count()===1);
 await page.evaluate(()=>localStorage.setItem('aqarProgress','null'));await page.reload();await ready(page);check('Null storage handled',errors.length===0);
 await context.close();
 const blocked=await browser.newContext();await blocked.addInitScript(()=>Object.defineProperty(window,'localStorage',{get(){throw new Error('disabled');}}));const bp=await blocked.newPage();await route(bp,'06-دليل-الدراسة-الكاملة.md');await bp.locator('.item input[type=checkbox]').first().check();check('Blocked storage still allows reading and marking',(await bp.locator('#storageStatus').innerText()).includes('غير متاح'));await blocked.close();
 const safe=await browser.newContext();await safe.route('**/06-*',r=>r.fulfill({body:'# Safety fixture\n<img src=x onerror="window.aqarInjected=true">\n<script>window.aqarInjected=true</script>\n[unsafe](javascript:alert(1))',contentType:'text/markdown'}));const sp=await safe.newPage();await route(sp,'06-دليل-الدراسة-الكاملة.md');check('Markdown HTML sanitized',await sp.evaluate(()=>!window.aqarInjected&&!document.querySelector('#content script')&&!document.querySelector('#content a[href^="javascript:"]')));await safe.close();
 const fallback=await browser.newContext();await fallback.route('**/vendor/*',r=>r.abort());const fp=await fallback.newPage();await route(fp,'06-دليل-الدراسة-الكاملة.md');check('Renderer failure falls back to safe readable text',await fp.locator('#content pre').count()===1);await fallback.close();
 check('No uncaught errors in primary browser session',errors.length===0,errors.join(';'));
 if(checks.some(c=>!c.pass))process.exitCode=1;
 fs.writeFileSync(path.join(OUT,'browser-results.json'),JSON.stringify({checks,screenshots,errors},null,2));console.log(JSON.stringify({checks:checks.length,passed:checks.filter(x=>x.pass).length,failed:checks.filter(x=>!x.pass),screenshotCount:screenshots.length},null,2));
})().catch(e=>{console.error(e);process.exitCode=1;}).finally(async()=>{if(browser)await browser.close();server.kill();});
