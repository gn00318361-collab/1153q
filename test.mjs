import fs from 'node:fs';import vm from 'node:vm';import assert from 'node:assert/strict';import http from 'node:http';import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
const {chromium}=require('C:/Users/csps/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const html=fs.readFileSync('Index.html','utf8'),code=fs.readFileSync('Code.gs','utf8'),cfg=JSON.parse(fs.readFileSync('data.json','utf8'));
let passed=0;function test(name,fn){fn();passed++;console.log('PASS',name);}
function backend(){
 const sheets=new Map(),props=new Map();let serial=0,locked=false,failWrite=false;
 class Sheet{
  constructor(name){this.name=name;this.rows=[];}
  getLastRow(){return this.rows.length;}
  getDataRange(){return {getValues:()=>structuredClone(this.rows)};}
  setFrozenRows(){}
  getRange(row,col,n,m){return {setNumberFormat(){},setValues:values=>{
    if(failWrite&&this.name==='提交狀態'&&row>1){failWrite=false;throw Error('mock write failure');}
    assert.equal(values.length,n);values.forEach((v,i)=>{assert.equal(v.length,m);this.rows[row-1+i]??=[];v.forEach((x,j)=>this.rows[row-1+i][col-1+j]=x);});
   }};}
 }
 const ss={getId:()=> 'test-sheet',getSheetByName:n=>sheets.get(n),insertSheet:n=>{const s=new Sheet(n);sheets.set(n,s);return s;}};
 const c=vm.createContext({console,Map,Set,Date,JSON,Number,Object,Array,Error,PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)??null,setProperty:(k,v)=>props.set(k,v)})},SpreadsheetApp:{openById:()=>ss,getActiveSpreadsheet:()=>ss,flush(){}},Utilities:{getUuid:()=>`version-${++serial}`,formatDate:()=>new Date().toISOString()},LockService:{getScriptLock:()=>({tryLock:()=>!locked,releaseLock(){}})},HtmlService:{createHtmlOutputFromFile:()=>({getContent:()=>html})}});
 vm.runInContext(code,c);c.setup_();c.openCollection_();
 return {c,sheets,props,lock:()=>locked=true,unlock:()=>locked=false,fail:()=>failWrite=true};
}
const b=backend(),j=cfg.judgesList[0];
let rid=0;const payload=(judge=j,cat='AQ',score=80,base='')=>({judge,category:cat,eventId:cfg.eventId,baseVersion:base,requestId:'request-'+(++rid),records:cfg.candidatesData[cat].map(s=>({id:s.id,score}))});
const sav=p=>JSON.parse(JSON.stringify(b.c.saveCategoryScores(p)));
test('32 candidates, distinct IDs and no excluded judge',()=>{assert.equal(Object.values(cfg.candidatesData).flat().length,32);assert.equal(new Set(Object.values(cfg.candidatesData).flat().map(s=>s.id)).size,32);assert(!cfg.judgesList.some(j=>j.includes('張芯')));});
test('empty cloud has stable schema',()=>{const s=b.c.getJudgeStatus(j);assert.equal(s.savedStatus.AQ,false);assert.equal(s.versions.AQ,'');});
for(const value of [-1,101,NaN,Infinity,'',null,'80'])test('reject invalid score '+value,()=>{const p=payload();p.records[0].score=value;assert.equal(sav(p).success,false);});
test('reject unknown judge',()=>assert.equal(sav(payload('not-authorized')).success,false));
test('reject incomplete category',()=>{const p=payload();p.records.pop();assert.equal(sav(p).success,false);});
test('reject duplicate IDs',()=>{const p=payload();p.records[1]=p.records[0];assert.equal(sav(p).success,false);});
test('reject wrong category student',()=>{const p=payload();p.records[0].id='MQ-601';assert.equal(sav(p).success,false);});
test('reject stale event',()=>{const p=payload();p.eventId='wrong';assert.equal(sav(p).success,false);});
let first=payload();first.records[0].score=0;first.records[1].score=100;first.records[2].score=88.25;let r;
test('save zero, 100 and decimal',()=>{r=sav(first);assert(r.success);assert.equal(b.c.getJudgeStatus(j).scores.AQ[first.records[0].id],0);});
test('retry same request is idempotent',()=>{assert.equal(sav(first).version,r.version);assert.equal(b.sheets.get('提交狀態').rows.length,2);assert.equal(b.sheets.get('評分明細').rows.length,13);});
test('request ID cannot carry different content',()=>{const p=structuredClone(first);p.records[0].score=2;assert.equal(sav(p).success,false);});
test('optimistic concurrency detects another device',()=>{const p=payload();assert.equal(sav(p).code,'CONFLICT');});
test('overwrite has one latest submission',()=>{const p=payload(j,'AQ',90,r.version);r=sav(p);assert(r.success);assert.equal(b.sheets.get('評分明細').rows.length,13);assert.equal(b.c.getJudgeStatus(j).scores.AQ['AQ-601'],90);});
test('write failure keeps previous submission',()=>{b.fail();const p=payload(j,'AQ',2,r.version);assert.equal(sav(p).success,false);assert.equal(b.c.getJudgeStatus(j).scores.AQ['AQ-601'],90);});
test('other judge has equal weight',()=>{assert(sav(payload(cfg.judgesList[1],'AQ',70)).success);const rows=b.sheets.get('總表').rows.filter(r=>r[0]==='AQ');assert(rows.every(r=>r[4]===2&&r[5]===80));});
test('boundary tie requires decision',()=>assert(b.sheets.get('總表').rows.filter(r=>r[0]==='AQ').every(r=>String(r[6]).includes('待決議'))));
test('category save does not touch other categories',()=>{assert(sav(payload(j,'MQ',55)).success);assert.equal(b.c.getJudgeStatus(j).scores.AQ['AQ-601'],90);assert.equal(b.c.getJudgeStatus(j).savedStatus.EQ,false);});
test('closed collection rejects mutation',()=>{b.c.closeCollection_();assert.equal(sav(payload(j,'EQ')).success,false);b.c.openCollection_();});
test('deadline is enforced on server',()=>{const original=b.c.config_;b.c.config_=()=>({...cfg,deadline:'2020-01-01T00:00:00+08:00'});assert.equal(sav(payload(j,'EQ')).success,false);b.c.config_=original;});
test('lock contention cannot mutate',()=>{b.lock();assert.equal(sav(payload(j,'EQ')).success,false);b.unlock();});
test('metadata uses trusted roster',()=>{const p=payload(j,'EQ');p.records.forEach(s=>{s.class='FAKE';s.name='FAKE';});assert(sav(p).success);assert(!JSON.stringify(b.sheets.get('評分明細').rows).includes('FAKE'));});
test('admin functions are private to RPC',()=>{for(const name of ['setup','refresh','openCollection','closeCollection','rebuildViews'])assert(code.includes('function '+name+'_('));});
// Actual browser UI wired to the same tested Apps Script logic via a fake google.script.run transport.
const server=http.createServer((req,res)=>{res.setHeader('Content-Type','text/html;charset=utf-8');res.end(html);});await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
const browser=await chromium.launch({headless:true,executablePath:'C:/Program Files/Google/Chrome/Application/chrome.exe'});
const ui=backend();let failNetwork=false;
const context=await browser.newContext({viewport:{width:390,height:844}});
await context.exposeBinding('backendCall',async(_,{name,arg})=>{if(failNetwork)throw Error('offline');return JSON.parse(JSON.stringify(ui.c[name](arg)));});
await context.addInitScript(()=>{window.google={script:{get run(){let success,fail;const api={withSuccessHandler(f){success=f;return api;},withFailureHandler(f){fail=f;return api;}};for(const name of ['getJudgeStatus','saveCategoryScores'])api[name]=arg=>window.backendCall({name,arg}).then(success,fail);return api;}}};});
const page=await context.newPage();const errors=[];page.on('pageerror',e=>errors.push(e.message));
async function utest(name,fn){await fn();passed++;console.log('PASS UI',name);}
try{
 await page.goto(`http://127.0.0.1:${server.address().port}`);await page.selectOption('#judge',j);await page.waitForFunction(()=>document.getElementById('message').textContent.includes('同步完成'));
 await utest('mobile layout and no rankings',async()=>{assert.equal(await page.locator('.student').count(),12);assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));assert.equal(await page.getByText('排行榜').count(),0);});
 await page.locator('#score-AQ-601').fill('0');await page.locator('#score-AQ-602').fill('101');
 await utest('incomplete/invalid cannot save',async()=>assert(await page.locator('#save').isDisabled()));
 await page.reload();await page.selectOption('#judge',j);await page.waitForFunction(()=>document.getElementById('message').textContent.includes('同步完成'));
 await utest('reload preserves drafts including zero',async()=>{assert.equal(await page.locator('#score-AQ-601').inputValue(),'0');assert.equal(await page.locator('#score-AQ-602').inputValue(),'101');});
 for(const s of cfg.candidatesData.AQ)await page.locator('#score-'+s.id).fill('86.5');
 await page.screenshot({path:'preview-mobile.png',fullPage:false});
 await page.click('#save');await page.waitForFunction(()=>document.getElementById('message').textContent.includes('成功儲存'));
 await utest('complete category saved',async()=>{assert.equal(ui.c.getJudgeStatus(j).scores.AQ['AQ-601'],86.5);assert((await page.locator('[data-cat="AQ"] small').textContent()).includes('已儲存'));});
 await page.locator('#score-AQ-601').fill('77');
 await utest('saved marker becomes dirty after edit',async()=>assert((await page.locator('[data-cat="AQ"] small').textContent()).includes('未送出')));
 failNetwork=true;await page.click('#save');await page.waitForFunction(()=>document.getElementById('message').textContent.includes('未確認'));
 await utest('network failure preserves draft',async()=>assert.equal(await page.locator('#score-AQ-601').inputValue(),'77'));failNetwork=false;
 const latest=ui.c.getJudgeStatus(j);ui.c.saveCategoryScores(payload(j,'AQ',91,latest.versions.AQ));
 await page.click('#sync');await page.waitForFunction(()=>!document.getElementById('conflict').classList.contains('hidden'));
 await utest('cross-device conflict retains local',async()=>assert.equal(await page.locator('#score-AQ-601').inputValue(),'77'));
 await page.click('#useCloud');await utest('choose cloud',async()=>assert.equal(await page.locator('#score-AQ-601').inputValue(),'91'));
 await page.locator('#score-AQ-601').fill('75');const l=ui.c.getJudgeStatus(j);ui.c.saveCategoryScores(payload(j,'AQ',92,l.versions.AQ));await page.click('#sync');await page.waitForFunction(()=>!document.getElementById('conflict').classList.contains('hidden'));await page.click('#keepLocal');await page.click('#save');await page.waitForFunction(()=>document.getElementById('message').textContent.includes('成功儲存'));
 await utest('explicit local resolution saves correctly',async()=>assert.equal(ui.c.getJudgeStatus(j).scores.AQ['AQ-601'],75));
 await page.selectOption('#judge',cfg.judgesList[1]);await page.waitForFunction(()=>document.getElementById('message').textContent.includes('同步完成'));
 await utest('judge drafts isolated',async()=>assert.equal(await page.locator('#score-AQ-601').inputValue(),''));
 await page.setViewportSize({width:1280,height:900});await page.screenshot({path:'preview-desktop.png',fullPage:false});
 await utest('no browser exceptions',async()=>assert.deepEqual(errors,[]));
}finally{await browser.close();server.close();}
fs.writeFileSync('TEST-RESULTS.txt',`${passed} checks passed. Google Sheets and Apps Script services mocked; live deployment validation still required.\n`);console.log('TOTAL',passed);
