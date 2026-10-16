// HISTORYVERSE 360 - PEMARKAHAN RINGKAS UASA SEJARAH T2 2026
// Isi: Objektif (20), B1-B4 (masing-masing 10), Esei pilihan 5/6 (20).
// Untuk mengelakkan konflik dengan borang lama, data baharu disimpan pada tab khusus.
const BOOK_ID = '1ZOr9ASygtWhnIDlrunv2K-Y0iduvlT7toKIrN_Kygxo';
const DATA_TAB = 'DATA MARKAH RINGKAS';
const HEAD = ['Kelas','Nama Murid','Objektif /20','Soalan 1 /10','Soalan 2 /10','Soalan 3 /10','Soalan 4 /10','Esei Pilihan (5/6)','Esei /20','Struktur /40','Jumlah /80','Peratus (%)','Gred','Tahap Pencapaian'];
const MAX = [20,10,10,10,10,20];

function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HistoryVerse 360 | Pemarkahan UASA Ringkas')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function _sheet(){
  const ss=SpreadsheetApp.openById(BOOK_ID);
  let s=ss.getSheetByName(DATA_TAB);
  if(!s){
    s=ss.insertSheet(DATA_TAB);
    s.getRange(1,1,1,HEAD.length).setValues([HEAD]);
    s.setFrozenRows(1);
    s.getRange(1,1,1,HEAD.length).setBackground('#54258c').setFontColor('#fff').setFontWeight('bold');
    s.setColumnWidth(2,300);
    s.getRange('C:I').setHorizontalAlignment('center');
  }
  return s;
}
function setupOnce(){
  const s=_sheet();
  const existing=s.getLastRow()>1 ? s.getRange(2,1,s.getLastRow()-1,2).getValues().map(x=>x.join('\u0000')) : [];
  const have=new Set(existing), append=[];
  Object.keys(ROSTER).forEach(cls=>ROSTER[cls].forEach(name=>{
    if(!have.has(cls+'\u0000'+name)) append.push([cls,name,...Array(HEAD.length-2).fill('')]);
  }));
  if(append.length) s.getRange(s.getLastRow()+1,1,append.length,HEAD.length).setValues(append);
  return {added:append.length, total:existing.length+append.length};
}
function _grade(percent){
  const p=Math.round(percent);
  if(p>=82)return ['A','Cemerlang'];
  if(p>=66)return ['B','Kepujian'];
  if(p>=50)return ['C','Baik'];
  if(p>=35)return ['D','Memuaskan'];
  if(p>=20)return ['E','Mencapai Tahap Minimum'];
  return ['F','Tidak Mencapai Tahap Minimum'];
}
function getAppData(){
  setupOnce();
  const s=_sheet(), rows=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,HEAD.length).getValues():[];
  const grouped={};Object.keys(ROSTER).forEach(c=>grouped[c]=[]);
  rows.forEach((r,i)=>{
    const marks=[r[2],r[3],r[4],r[5],r[6],r[8]].map(v=>v===''?null:Number(v));
    const started=marks.some(v=>v!==null);
    grouped[r[0]] ??= [];
    grouped[r[0]].push({row:i+2,name:r[1],marks,essayQuestion:r[7]||'',score:started?Number(r[10]):null,pct:started?Number(r[11]):null,grade:started?r[12]:''});
  });
  return {classes:Object.keys(ROSTER),grouped,max:MAX};
}
function saveMarks(payload){
  if(!payload||typeof payload!=='object')throw new Error('Data tidak sah');
  const cls=String(payload.cls||''),name=String(payload.name||'');
  if(!Object.prototype.hasOwnProperty.call(ROSTER,cls)||!ROSTER[cls].includes(name))throw new Error('Nama atau kelas tiada dalam senarai');
  const marks=payload.marks;
  if(!Array.isArray(marks)||marks.length!==MAX.length)throw new Error('Markah tidak lengkap');
  const clean=marks.map((v,i)=>{
    if(v===''||v===null||v===undefined)return '';
    const n=Number(v);
    if(!Number.isInteger(n)||n<0||n>MAX[i])throw new Error('Markah '+(i+1)+' mesti 0 hingga '+MAX[i]);
    return n;
  });
  const question=String(payload.essayQuestion||'');
  if(question!==''&&question!=='5'&&question!=='6')throw new Error('Pilih Esei Soalan 5 atau 6');
  if(clean[5]!==''&&question==='')throw new Error('Pilih Soalan Esei 5 atau 6');
  const started=clean.some(v=>v!=='');
  const structure=clean.slice(1,5).reduce((a,v)=>a+(v===''?0:v),0);
  const total=(clean[0]===''?0:clean[0])+structure+(clean[5]===''?0:clean[5]);
  const pct=Math.round(total/80*1000)/10;
  const [grade,level]=_grade(pct);
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const s=_sheet();
    const vals=s.getRange(2,1,s.getLastRow()-1,2).getValues();
    const index=vals.findIndex(r=>r[0]===cls&&r[1]===name);
    if(index<0)throw new Error('Rekod tidak dijumpai. Jalankan setupOnce.');
    const row=index+2;
    const output=[clean[0],...clean.slice(1,5),question,clean[5],started?structure:'',started?total:'',started?pct:'',started?grade:'',started?level:''];
    s.getRange(row,3,1,output.length).setValues([output]);
    return {ok:true,row,score:started?total:null,pct:started?pct:null,grade:started?grade:''};
  }finally{lock.releaseLock()}
}
