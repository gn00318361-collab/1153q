/* Bound to the administrator's Google Sheet. Index.html app-data is the single configuration source. */
const CATEGORIES_ = ['AQ','MQ','EQ'];
const SUB_HEADERS_ = ['評審','類別','版本','時間','完整分數JSON','請求ID'];
const DETAIL_HEADERS_ = ['時間戳記','評審委員','類別','班級','學生姓名','給分','學生ID'];
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('青溪國小 3Q達人評選系統')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function config_() {
  const raw=HtmlService.createHtmlOutputFromFile('Index').getContent();
  const c=JSON.parse(raw.match(/<script id="app-data" type="application\/json">([\s\S]*?)<\/script>/)[1]);
  if(c.deadline && !Number.isFinite(Date.parse(c.deadline))) throw Error('截止時間設定格式不正確');
  if(new Set(c.judgesList).size!==c.judgesList.length) throw Error('評審名單重複');
  return c;
}
function ss_() {
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id) throw Error('請管理員先執行 setup_ 初始化');
  return SpreadsheetApp.openById(id);
}
function sheet_(name,headers) {
  let s=ss_().getSheetByName(name);
  if(!s){s=ss_().insertSheet(name);s.getRange(1,1,1,headers.length).setValues([headers]);s.setFrozenRows(1);}
  return s;
}
function withLock_(fn) {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(20000))throw Error('目前有人正在儲存，請稍後重試；本機草稿仍保留。');
  try{return fn();}finally{lock.releaseLock();}
}
function validateJudge_(c,judge){if(typeof judge!=='string'||!c.judgesList.includes(judge))throw Error('請選擇有效評審姓名');}
function validateRecords_(c,cat,records){
  if(!CATEGORIES_.includes(cat)||!Array.isArray(records))throw Error('類別或資料格式不正確');
  const expected=c.candidatesData[cat];
  if(records.length!==expected.length)throw Error('請完成本類別所有候選人的評分');
  const map=new Map();
  records.forEach(r=>{
    if(!r||typeof r.id!=='string'||map.has(r.id)||typeof r.score!=='number'||!Number.isFinite(r.score)||r.score<0||r.score>100)throw Error('分數須為0～100，且學生不可重複');
    map.set(r.id,r.score);
  });
  return expected.map(s=>{if(!map.has(s.id))throw Error('候選名單不一致，請重新整理');return {id:s.id,class:s.class,name:s.name,score:map.get(s.id)};});
}
function stateRows_(){return sheet_('提交狀態',SUB_HEADERS_).getDataRange().getValues().slice(1);}
function isOpen_(c){return PropertiesService.getScriptProperties().getProperty('COLLECTION_OPEN')==='true'&&(!c.deadline||Date.now()<=Date.parse(c.deadline));}
function getJudgeStatus(judgeName){
 return withLock_(()=>{
  const c=config_();validateJudge_(c,judgeName);
  const out={scores:{AQ:{},MQ:{},EQ:{}},savedStatus:{AQ:false,MQ:false,EQ:false},versions:{AQ:'',MQ:'',EQ:''},times:{},open:isOpen_(c),eventId:c.eventId,deadline:c.deadline};
  stateRows_().filter(r=>r[0]===judgeName&&CATEGORIES_.includes(r[1])).forEach(r=>{
    const valid=validateRecords_(c,r[1],JSON.parse(r[4]));
    out.scores[r[1]]=Object.fromEntries(valid.map(v=>[v.id,v.score]));out.savedStatus[r[1]]=true;out.versions[r[1]]=r[2];out.times[r[1]]=String(r[3]);
  });return out;
 });
}
function saveCategoryScores(payload){
 try{return withLock_(()=>{
  const c=config_();
  if(!payload||payload.eventId!==c.eventId)throw Error('活動版本已變更，請重新整理');
  validateJudge_(c,payload.judge);
  const records=validateRecords_(c,payload.category,payload.records);
  if(typeof payload.requestId!=='string'||!/^[\w-]{8,100}$/.test(payload.requestId))throw Error('請求識別碼無效');
  const s=sheet_('提交狀態',SUB_HEADERS_), rows=stateRows_();
  const index=rows.findIndex(r=>r[0]===payload.judge&&r[1]===payload.category),old=index<0?null:rows[index];
  if(old&&old[5]===payload.requestId){
    if(old[4]!==JSON.stringify(records))throw Error('同一請求不可更改內容');
    return {success:true,version:old[2],time:String(old[3]),replayed:true};
  }
  if(!isOpen_(c))throw Error('目前已停止收件，尚未送出的修改仍保留於本機');
  if((old?old[2]:'')!==payload.baseVersion)return {success:false,code:'CONFLICT',message:'另一裝置已更新此類別。請重新同步後選擇保留哪份分數。'};
  const version=Utilities.getUuid(),time=Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss');
  // One row is the authoritative commit; do not delete old scores before writing.
  s.getRange(index<0?s.getLastRow()+1:index+2,1,1,6).setValues([[payload.judge,payload.category,version,time,JSON.stringify(records),payload.requestId]]);
  SpreadsheetApp.flush();
  let warning='';
  try{rebuildViews_();}catch(e){warning='分數已儲存，管理總表待管理員重新整理。';}
  return {success:true,version,time,warning};
 });}catch(e){return {success:false,code:'ERROR',message:e.message||String(e)};}
}
function writeView_(name,headers,rows){
 const s=sheet_(name,headers),n=Math.max(s.getLastRow(),rows.length+1);
 const values=[headers,...rows];while(values.length<n)values.push(headers.map(()=>''));
 s.getRange(1,1,n,headers.length).setValues(values);s.setFrozenRows(1);
}
function aggregate_(c,submissions){
 const details=[],progress=[],rankings=[],latest=new Map();
 submissions.forEach(r=>{if(c.judgesList.includes(r[0])&&CATEGORIES_.includes(r[1]))latest.set(JSON.stringify([r[0],r[1]]),r);});
 c.judgesList.forEach(judge=>{
  const done=[];
  CATEGORIES_.forEach(cat=>{
   const row=latest.get(JSON.stringify([judge,cat]));
   try{const rec=validateRecords_(c,cat,JSON.parse(row[4]));rec.forEach(r=>details.push([String(row[3]),judge,cat,r.class,r.name,r.score,r.id]));done.push('已儲存');}
   catch{done.push(row?'資料異常，未計入':'未儲存');}
  });progress.push([judge,...done]);
 });
 CATEGORIES_.forEach(cat=>{
  const items=c.candidatesData[cat].map(s=>{const ds=details.filter(r=>r[2]===cat&&r[6]===s.id);return {...s,n:ds.length,total:ds.reduce((a,r)=>a+r[5],0)};});
  items.sort((a,b)=>b.total-a.total||a.id.localeCompare(b.id));
  const boundary=items[2],cutTie=boundary&&items[3]&&boundary.n&&boundary.total===items[3].total;
  items.forEach((s,i)=>{
   const rank=s.n?items.filter(x=>x.total>s.total).length+1:'';
   const status=!s.n?'尚無有效評分':cutTie&&s.total===boundary.total?'第三名邊界同分，待決議':rank<=3?'前三名（暫定）':'未列前三';
   rankings.push([cat,rank,s.class,s.name,s.n,s.n?s.total/s.n:'',status,s.id]);
  });
 });return {details,progress,rankings};
}
function rebuildViews_(){
 const a=aggregate_(config_(),stateRows_());
 writeView_('評分明細',DETAIL_HEADERS_,a.details);
 writeView_('評審進度',['評審委員','AQ','MQ','EQ'],a.progress);
 writeView_('總表',['類別','名次','班級','學生姓名','有效評審數','平均得分','錄取狀態','學生ID'],a.rankings);
 ss_().getSheetByName('總表').getRange(2,6,Math.max(1,a.rankings.length),1).setNumberFormat('0.00');
}
function onOpen(){SpreadsheetApp.getUi().createMenu('3Q評選管理').addItem('初始化（不清除分數）','setup_').addItem('更新明細與總表','refresh_').addItem('開放收件','openCollection_').addItem('停止收件','closeCollection_').addToUi();}
function setup_(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw Error('請由儲存評分的Google試算表開啟Apps Script');
 const props=PropertiesService.getScriptProperties();props.setProperty('SPREADSHEET_ID',ss.getId());
 if(props.getProperty('COLLECTION_OPEN')===null)props.setProperty('COLLECTION_OPEN','false');
 withLock_(()=>rebuildViews_());
}
function refresh_(){return withLock_(()=>rebuildViews_());}
function openCollection_(){return withLock_(()=>{config_();PropertiesService.getScriptProperties().setProperty('COLLECTION_OPEN','true');});}
function closeCollection_(){return withLock_(()=>PropertiesService.getScriptProperties().setProperty('COLLECTION_OPEN','false'));}
