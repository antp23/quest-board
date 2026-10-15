// Quest Board <-> "Family to-do list" sheet. Deploy: Web app, Execute as Me, Access Anyone.
const SHEET_ID='1RNJgln7R7rDdPywD2xHMP5WnccqTXxgMexRQjRr3fx0', TAB='Untitled';
const KEYS={'Task':'task','Owner':'owner','Due date':'due_date','Status':'status','Category':'category','Priority':'priority','Next step':'next_step'};
function sh_(){return SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);}
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function fmt_(v){return v instanceof Date?Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd'):String(v==null?'':v);}
function list_(){const v=sh_().getDataRange().getValues(),h=v[0].map(String),tasks=[];
  for(let i=1;i<v.length;i++){if(!String(v[i][0]).trim())continue;const t={row:i+1};
    h.forEach((n,j)=>{if(KEYS[n.trim()])t[KEYS[n.trim()]]=fmt_(v[i][j]);});tasks.push(t);}
  return {updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm"),tasks};}
function set_(p){const s=sh_(),h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const sc=h.indexOf('Status')+1,tc=h.indexOf('Task')+1,row=parseInt(p.row,10);
  if(!(row>1&&row<=s.getLastRow()))return {ok:false,error:'bad row'};
  if(p.task&&String(s.getRange(row,tc).getValue())!==p.task)return {ok:false,error:'row moved; reload'};
  s.getRange(row,sc).setValue(p.action==='done'?'Done':'Open');return {ok:true,row};}
function handle_(p){const l=LockService.getScriptLock();l.waitLock(10000);
  try{return out_(p.action==='done'||p.action==='undo'?set_(p):list_());}finally{l.releaseLock();}}
function doGet(e){return handle_(e.parameter||{});}
function doPost(e){let p={};try{p=JSON.parse(e.postData.contents);}catch(x){p=e.parameter||{};}return handle_(p);}
