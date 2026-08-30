const CONFIG={
  SPREADSHEET_ID:'1053aBXdsO0gTutrJuMMGMFFpSMSVlRZ6f-QBf1wAuJw',
  SHEETS:{POSTERS:'SENARAI POSTER',JUDGES:'NAMA HAKIM',SCORES:'DATA PENJURIAN',AUDIT:'AUDIT PERUBAHAN',SETTINGS:'TETAPAN SISTEM'},
  SESSION_HOURS:12,
  GALLERY_PAGE_SIZE:12
};
const HEADERS={
  POSTERS:['ID Poster','No. Giliran','Nama Peserta','Nama Sekolah','Nama Poster','Link Poster','Kategori','Status Aktif','Tarikh Kemas Kini'],
  JUDGES:['ID Hakim','Nama Hakim','Status Aktif','Susunan Hakim'],
  SCORES:['Timestamp','ID Hakim','Nama Hakim','ID Poster','Nama Poster','Nama Sekolah','Kandungan & Kaitan Tema','Reka Bentuk & Kreativiti','Keberkesanan & Kejelasan Maklumat','Kualiti Teknikal','Potensi & Impak','Jumlah','Status','Tarikh Kemas Kini','Kunci Rekod','Versi Rekod','Request ID'],
  AUDIT:['Timestamp','ID Audit','ID Hakim','ID Poster','Tindakan','Markah Lama','Markah Baharu','Jumlah Lama','Jumlah Baharu','Request ID'],
  SETTINGS:['Kunci','Nilai','Keterangan']
};
const SCORE_KEYS={content:'Kandungan & Kaitan Tema',design:'Reka Bentuk & Kreativiti',clarity:'Keberkesanan & Kejelasan Maklumat',technical:'Kualiti Teknikal',impact:'Potensi & Impak'};

function doGet(e){return route_((e&&e.parameter)||{},false);}
function doPost(e){
  let body={};try{body=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');}catch(err){return json_({ok:false,message:'Format permintaan tidak sah.'});}
  return route_(body,true);
}
function route_(req,isPost){
  try{
    const action=String(req.action||'health'); let data;
    if(action==='health')data={service:'PUICE 2026 Poster Digital API',status:'ready',time:new Date()};
    else if(action==='bootstrap')data=bootstrap_();
    else if(action==='login')data=login_(req);
    else if(action==='dashboard')data=dashboard_(req);
    else if(action==='saveScore')data=saveScore_(req);
    else if(action==='gallery')data=gallery_(req);
    else if(action==='leaderboard')data=leaderboard_(req);
    else throw new Error('Tindakan API tidak dikenali.');
    return json_({ok:true,data:data});
  }catch(err){console.error(err&&err.stack?err.stack:err);return json_({ok:false,message:cleanError_(err)});}
}
function json_(value){return ContentService.createTextOutput(JSON.stringify(value,function(k,v){return v instanceof Date?v.toISOString():v;})).setMimeType(ContentService.MimeType.JSON);}
function cleanError_(err){const msg=String(err&&err.message?err.message:err||'Ralat tidak diketahui.');return msg.indexOf('Exception:')===0?msg.slice(10).trim():msg;}

function setupSheets(){
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); const blank=ss.getSheetByName('Sheet1');
  if(blank&&!ss.getSheetByName(CONFIG.SHEETS.POSTERS)&&blank.getLastRow()===0)blank.setName(CONFIG.SHEETS.POSTERS);
  Object.keys(HEADERS).forEach(function(key){ensureSheet_(ss,CONFIG.SHEETS[key],HEADERS[key]);});
  upsertSeed_(CONFIG.SHEETS.JUDGES,'ID Hakim',[
    {'ID Hakim':'H01','Nama Hakim':'SITI HIDAYU BINTI LUDIN','Status Aktif':true,'Susunan Hakim':1},
    {'ID Hakim':'H02','Nama Hakim':'Khamizah binti Kamarudin','Status Aktif':true,'Susunan Hakim':2},
    {'ID Hakim':'H03','Nama Hakim':'Mohd Muhaimin bin Che Mansor','Status Aktif':true,'Susunan Hakim':3}
  ]);
  upsertSeed_(CONFIG.SHEETS.POSTERS,'ID Poster',[
    {'ID Poster':'P001','No. Giliran':1,'Nama Peserta':'Peserta Contoh 1','Nama Sekolah':'Sekolah Contoh 1','Nama Poster':'Poster Digital 1','Link Poster':'https://drive.google.com/file/d/1CPX1GEgSKbqvozIWDhctNiXC3cabaLbN/view','Kategori':'Terbuka','Status Aktif':true,'Tarikh Kemas Kini':new Date()},
    {'ID Poster':'P002','No. Giliran':2,'Nama Peserta':'Peserta Contoh 2','Nama Sekolah':'Sekolah Contoh 2','Nama Poster':'Poster Digital 2','Link Poster':'https://drive.google.com/file/d/1q9n6D7zkn2BJHn6R4JxdvuEzpbHheT0W/view','Kategori':'Terbuka','Status Aktif':true,'Tarikh Kemas Kini':new Date()},
    {'ID Poster':'P003','No. Giliran':3,'Nama Peserta':'Peserta Contoh 3','Nama Sekolah':'Sekolah Contoh 3','Nama Poster':'Poster Digital 3','Link Poster':'https://drive.google.com/file/d/1UPowuwWolILsdzELk4XAV-yMhEZF3_cD/view','Kategori':'Terbuka','Status Aktif':true,'Tarikh Kemas Kini':new Date()}
  ]);
  upsertSeed_(CONFIG.SHEETS.SETTINGS,'Kunci',[
    {'Kunci':'PENJURIAN_DIBUKA','Nilai':'TRUE','Keterangan':'Benarkan hakim menghantar markah'},
    {'Kunci':'EDIT_MARKAH_DIBENARKAN','Nilai':'TRUE','Keterangan':'Benarkan kemas kini markah sedia ada'},
    {'Kunci':'KUOTA_PLATINUM','Nilai':'5','Keterangan':'Kedudukan 1 hingga 5'},
    {'Kunci':'KUOTA_EMAS','Nilai':'36','Keterangan':'36 penerima selepas Platinum'},
    {'Kunci':'KUOTA_PERAK','Nilai':'63','Keterangan':'63 penerima selepas Emas'},
    {'Kunci':'SEMUA_TERIMA_ANUGERAH','Nilai':'TRUE','Keterangan':'Baki peserta menerima Gangsa'},
    {'Kunci':'TIE_BREAK','Nilai':'IMPAK>KANDUNGAN>REKA_BENTUK>MANUAL','Keterangan':'Dasar tie-break sempadan anugerah'}
  ]);
  formatSheets_(ss); SpreadsheetApp.flush(); return 'Setup selesai.';
}
function ensureSheet_(ss,name,headers){
  let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);const existing=sh.getRange(1,1,1,headers.length).getDisplayValues()[0];
  if(existing.every(function(v){return !v;}))sh.getRange(1,1,1,headers.length).setValues([headers]);
  else if(headers.some(function(h,i){return existing[i]!==h;}))throw new Error('Header tab '+name+' tidak sepadan. Tiada data diubah.');return sh;
}
function upsertSeed_(sheetName,keyName,records){
  const sh=getSheet_(sheetName),headers=HEADERS[Object.keys(CONFIG.SHEETS).find(function(k){return CONFIG.SHEETS[k]===sheetName;})],keyIndex=headers.indexOf(keyName),values=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],known={};
  values.forEach(function(row){known[String(row[keyIndex])]=true;});const add=records.filter(function(r){return !known[String(r[keyName])];}).map(function(r){return headers.map(function(h){return Object.prototype.hasOwnProperty.call(r,h)?r[h]:'';});});if(add.length)sh.getRange(sh.getLastRow()+1,1,add.length,headers.length).setValues(add);
}
function formatSheets_(ss){Object.keys(CONFIG.SHEETS).forEach(function(k){const sh=ss.getSheetByName(CONFIG.SHEETS[k]);if(!sh)return;const cols=HEADERS[k].length;sh.setFrozenRows(1);sh.getRange(1,1,1,cols).setBackground('#10233f').setFontColor('#ffffff').setFontWeight('bold');sh.autoResizeColumns(1,cols);});}

function bootstrap_(){const judges=getActiveJudges_(),posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES);return {judges:judges.map(function(j){const p=progressFor_(j.id,posters,scores);return {id:j.id,name:j.name,completed:p.completed,total:p.total};}),posterCount:posters.length,ready:true};}
function login_(req){const judge=getActiveJudges_().find(function(j){return j.id===String(req.judgeId||'');});if(!judge)throw new Error('Hakim tidak sah atau tidak aktif.');if(!verifyPin_(judge.id,String(req.pin||'')))throw new Error('PIN tidak tepat. Sila cuba semula.');return {token:createToken_(judge.id),judge:{id:judge.id,name:judge.name}};}
function dashboard_(req){const judge=authenticatedJudge_(req.token),posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES),mine=scores.filter(function(x){return String(x['ID Hakim'])===judge.id;}),judged={};mine.forEach(function(x){judged[String(x['ID Poster'])]=x;});const remaining=posters.filter(function(p){return !judged[p.id];}),progress=progressFor_(judge.id,posters,scores),last=mine.reduce(function(max,x){const d=new Date(x['Tarikh Kemas Kini']);return d>max?d:max;},new Date(0));return {judge:{id:judge.id,name:judge.name},poster:remaining[0]?posterPayload_(remaining[0],null):null,nextPoster:remaining[1]?posterPayload_(remaining[1],null):null,progress:progress,lastUpdated:last.getTime()?last:null};}
function saveScore_(req){
  const judge=authenticatedJudge_(req.token);if(String(getSetting_('PENJURIAN_DIBUKA','TRUE')).toUpperCase()!=='TRUE')throw new Error('Penjurian sedang ditutup oleh urusetia.');
  const posters=getActivePosters_(),poster=posters.find(function(p){return p.id===String(req.posterId||'');});if(!poster)throw new Error('ID poster tidak sah atau tidak aktif.');const clean=validateScores_(req.scores||{}),total=Object.keys(clean).reduce(function(a,k){return a+clean[k];},0),requestId=String(req.requestId||Utilities.getUuid()),lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const sh=getSheet_(CONFIG.SHEETS.SCORES),headers=HEADERS.SCORES,all=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],key=judge.id+'::'+poster.id,keyCol=headers.indexOf('Kunci Rekod'),requestCol=headers.indexOf('Request ID'),existingIndex=all.findIndex(function(r){return String(r[keyCol])===key;}),duplicate=all.find(function(r){return String(r[requestCol])===requestId;});if(duplicate)return dashboard_(req);
    const now=new Date(),old=existingIndex>=0?rowObject_(headers,all[existingIndex]):null;if(old&&String(getSetting_('EDIT_MARKAH_DIBENARKAN','TRUE')).toUpperCase()!=='TRUE')throw new Error('Markah sedia ada tidak boleh diedit.');
    const record={'Timestamp':old?old.Timestamp:now,'ID Hakim':judge.id,'Nama Hakim':judge.name,'ID Poster':poster.id,'Nama Poster':poster.title,'Nama Sekolah':poster.school,'Kandungan & Kaitan Tema':clean.content,'Reka Bentuk & Kreativiti':clean.design,'Keberkesanan & Kejelasan Maklumat':clean.clarity,'Kualiti Teknikal':clean.technical,'Potensi & Impak':clean.impact,'Jumlah':total,'Status':'DISIMPAN','Tarikh Kemas Kini':now,'Kunci Rekod':key,'Versi Rekod':old?(Number(old['Versi Rekod'])||1)+1:1,'Request ID':requestId},row=headers.map(function(h){return record[h];});
    if(existingIndex>=0)sh.getRange(existingIndex+2,1,1,headers.length).setValues([row]);else sh.appendRow(row);appendAudit_(judge.id,poster.id,old?'KEMAS_KINI':'CIPTA',old,record,requestId);SpreadsheetApp.flush();return dashboard_(req);
  }finally{lock.releaseLock();}
}
function gallery_(req){const judge=authenticatedJudge_(req.token),posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES),mine={};scores.filter(function(x){return String(x['ID Hakim'])===judge.id;}).forEach(function(x){mine[String(x['ID Poster'])]=x;});let items=posters.map(function(p){return posterPayload_(p,mine[p.id]||null);}),status=String(req.status||'all'),query=String(req.query||'').toLowerCase(),category=String(req.category||'all');if(status==='judged')items=items.filter(function(x){return x.judged;});if(status==='unjudged')items=items.filter(function(x){return !x.judged;});if(query)items=items.filter(function(x){return (x.title+' '+x.school+' '+x.id).toLowerCase().indexOf(query)>=0;});if(category!=='all')items=items.filter(function(x){return x.category===category;});const size=Math.min(12,Math.max(1,Number(req.pageSize)||CONFIG.GALLERY_PAGE_SIZE)),pages=Math.max(1,Math.ceil(items.length/size)),page=Math.min(pages,Math.max(1,Number(req.page)||1));return {items:items.slice((page-1)*size,page*size),page:page,pages:pages,total:items.length,categories:Array.from(new Set(posters.map(function(p){return p.category;}).filter(Boolean))).sort()};}
function leaderboard_(req){authenticatedJudge_(req.token);const judges=getActiveJudges_(),posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES),progress=judges.map(function(j){const p=progressFor_(j.id,posters,scores);return {completed:p.completed,total:p.total};});if(progress.some(function(p){return p.completed<p.total;}))return {complete:false,judges:progress};const map={};scores.forEach(function(s){const pid=String(s['ID Poster']),jid=String(s['ID Hakim']);if(!map[pid])map[pid]={};map[pid][jid]=s;});let items=posters.map(function(p){const by=map[p.id]||{},judgeTotals=judges.map(function(j){return Number((by[j.id]||{}).Jumlah)||0;}),aspect=function(header){return judges.reduce(function(a,j){return a+(Number((by[j.id]||{})[header])||0);},0);};return {id:p.id,participant:p.participant,title:p.title,school:p.school,judgeTotals:judgeTotals,total:judgeTotals.reduce(function(a,b){return a+b;},0),impact:aspect(SCORE_KEYS.impact),content:aspect(SCORE_KEYS.content),design:aspect(SCORE_KEYS.design)};});items.sort(function(a,b){return b.total-a.total||b.impact-a.impact||b.content-a.content||b.design-a.design||a.id.localeCompare(b.id);});const qP=Number(getSetting_('KUOTA_PLATINUM',5)),qE=Number(getSetting_('KUOTA_EMAS',36)),qS=Number(getSetting_('KUOTA_PERAK',63)),allAward=String(getSetting_('SEMUA_TERIMA_ANUGERAH','TRUE')).toUpperCase()==='TRUE';items.forEach(function(x,i){x.rank=i+1;x.award=i<qP?'Platinum':i<qP+qE?'Emas':i<qP+qE+qS?'Perak':allAward?'Gangsa':'Tiada Anugerah';delete x.impact;delete x.content;delete x.design;});return {complete:true,items:items};}

function authenticatedJudge_(token){const payload=verifyToken_(String(token||'')),judge=getActiveJudges_().find(function(j){return j.id===payload.judgeId;});if(!judge)throw new Error('Sesi hakim tidak sah. Sila log masuk semula.');return judge;}
function getActiveJudges_(){return getRows_(CONFIG.SHEETS.JUDGES).filter(function(r){return isActive_(r['Status Aktif']);}).map(function(r){return {id:String(r['ID Hakim']),name:String(r['Nama Hakim']),order:Number(r['Susunan Hakim'])||999};}).sort(function(a,b){return a.order-b.order;});}
function getActivePosters_(){return getRows_(CONFIG.SHEETS.POSTERS).filter(function(r){return isActive_(r['Status Aktif']);}).map(function(r){const link=String(r['Link Poster']||'');return {id:String(r['ID Poster']),number:r['No. Giliran'],participant:String(r['Nama Peserta']||''),school:String(r['Nama Sekolah']),title:String(r['Nama Poster']),link:link,imageUrl:driveImage_(link,1600),thumbnailUrl:driveImage_(link,600),category:String(r.Kategori||'')};}).sort(function(a,b){return Number(a.number)-Number(b.number)||a.id.localeCompare(b.id);});}
function posterPayload_(p,score){const out={id:p.id,number:p.number,school:p.school,title:p.title,link:p.link,imageUrl:p.imageUrl,thumbnailUrl:p.thumbnailUrl,category:p.category,judged:!!score};if(score){out.scores={content:Number(score[SCORE_KEYS.content]),design:Number(score[SCORE_KEYS.design]),clarity:Number(score[SCORE_KEYS.clarity]),technical:Number(score[SCORE_KEYS.technical]),impact:Number(score[SCORE_KEYS.impact])};out.total=Number(score.Jumlah);out.updatedAt=score['Tarikh Kemas Kini'];}return out;}
function progressFor_(judgeId,posters,scores){const active={};posters.forEach(function(p){active[p.id]=true;});const done={};scores.forEach(function(s){if(String(s['ID Hakim'])===judgeId&&active[String(s['ID Poster'])])done[String(s['ID Poster'])]=true;});return {completed:Object.keys(done).length,total:posters.length};}
function validateScores_(scores){const out={};Object.keys(SCORE_KEYS).forEach(function(k){const n=Number(scores[k]);if(!Number.isInteger(n)||n<1||n>20)throw new Error('Semua lima aspek mestilah markah bulat antara 1 hingga 20.');out[k]=n;});return out;}
function appendAudit_(judgeId,posterId,action,oldRow,newRow,requestId){const oldScores=oldRow?scoreSnapshot_(oldRow):'',newScores=scoreSnapshot_(newRow),row=[new Date(),Utilities.getUuid(),judgeId,posterId,action,oldScores?JSON.stringify(oldScores):'',JSON.stringify(newScores),oldRow?Number(oldRow.Jumlah)||0:'',Number(newRow.Jumlah)||0,requestId];getSheet_(CONFIG.SHEETS.AUDIT).appendRow(row);}
function scoreSnapshot_(r){return {content:Number(r[SCORE_KEYS.content]),design:Number(r[SCORE_KEYS.design]),clarity:Number(r[SCORE_KEYS.clarity]),technical:Number(r[SCORE_KEYS.technical]),impact:Number(r[SCORE_KEYS.impact])};}
function getRows_(name){const sh=getSheet_(name),last=sh.getLastRow(),cols=sh.getLastColumn();if(last<2)return [];const values=sh.getRange(1,1,last,cols).getValues(),headers=values.shift();return values.filter(function(row){return row.some(function(v){return v!=='';});}).map(function(row){return rowObject_(headers,row);});}
function rowObject_(headers,row){const o={};headers.forEach(function(h,i){o[h]=row[i];});return o;}
function getSheet_(name){const sh=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(name);if(!sh)throw new Error('Tab '+name+' belum disediakan.');return sh;}
function getSetting_(key,fallback){const row=getRows_(CONFIG.SHEETS.SETTINGS).find(function(r){return String(r.Kunci)===key;});return row?row.Nilai:fallback;}
function isActive_(v){return v===true||['TRUE','YA','AKTIF','1'].indexOf(String(v).toUpperCase())>=0;}
function driveImage_(link,size){const m=String(link).match(/\/d\/([a-zA-Z0-9_-]+)/)||String(link).match(/[?&]id=([a-zA-Z0-9_-]+)/);return m?'https://drive.google.com/thumbnail?id='+m[1]+'&sz=w'+size:link;}

function verifyPin_(judgeId,pin){const props=PropertiesService.getScriptProperties(),salt=props.getProperty('PIN_SALT'),expected=props.getProperty('PIN_HASH_'+judgeId);if(!salt||!expected)throw new Error('PIN hakim belum dikonfigurasi oleh urusetia.');return hash_(salt+':'+pin)===expected;}
function createToken_(judgeId){const props=PropertiesService.getScriptProperties(),secret=props.getProperty('SESSION_SECRET');if(!secret)throw new Error('Rahsia sesi belum dikonfigurasi.');const payload=Utilities.base64EncodeWebSafe(JSON.stringify({judgeId:judgeId,exp:Date.now()+CONFIG.SESSION_HOURS*3600000}),Utilities.Charset.UTF_8).replace(/=+$/,'');return payload+'.'+sign_(payload,secret);}
function verifyToken_(token){const parts=token.split('.');if(parts.length!==2)throw new Error('Sesi telah tamat. Sila log masuk semula.');const secret=PropertiesService.getScriptProperties().getProperty('SESSION_SECRET');if(!secret||sign_(parts[0],secret)!==parts[1])throw new Error('Sesi tidak sah. Sila log masuk semula.');let data;try{data=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());}catch(e){throw new Error('Sesi tidak sah.');}if(!data.exp||Date.now()>data.exp)throw new Error('Sesi telah tamat. Sila log masuk semula.');return data;}
function sign_(value,secret){return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value,secret)).replace(/=+$/,'');}
function hash_(value){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value,Utilities.Charset.UTF_8).map(function(b){return ('0'+(b<0?b+256:b).toString(16)).slice(-2);}).join('');}
