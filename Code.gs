const CONFIG={
  SPREADSHEET_ID:'1053aBXdsO0gTutrJuMMGMFFpSMSVlRZ6f-QBf1wAuJw',
  SHEETS:{POSTERS:'SENARAI POSTER',JUDGES:'NAMA HAKIM',SCORES:'DATA PENJURIAN',AUDIT:'AUDIT PERUBAHAN',SETTINGS:'TETAPAN SISTEM'},
  SESSION_HOURS:12,
  GALLERY_PAGE_SIZE:12
};
const HEADERS={
  POSTERS:['ID Poster','No. Giliran','Nama Peserta','Nama Sekolah','Nama Poster','Link Poster','Kategori','Status Aktif','Tarikh Kemas Kini'],
  JUDGES:['ID Hakim','Nama Hakim','Status Aktif','Susunan Hakim'],
  SCORES:['Timestamp','ID Hakim','Nama Hakim','ID Poster','Nama Poster','Nama Sekolah','Kandungan & Kaitan Tema','Reka Bentuk & Kreativiti','Keberkesanan & Kejelasan Maklumat','Kualiti Teknikal','Potensi & Impak','Jumlah','Status','Tarikh Kemas Kini','Kunci Rekod','Versi Rekod','Request ID','Catatan DQ'],
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
    else if(action==='selectPoster')data=selectPoster_(req);
    else if(action==='saveScore')data=saveScore_(req);
    else if(action==='saveDisqualification')data=saveDisqualification_(req);
    else if(action==='gallery')data=gallery_(req);
    else if(action==='summary')data=summary_(req);
    else if(action==='leaderboard')data=leaderboard_(req);
    else if(action==='publicLeaderboard')data=publicLeaderboard_();
    else if(action==='adminLogin')data=adminLogin_(req);
    else if(action==='adminLeaderboard')data=adminLeaderboard_(req);
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
  let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);const used=Math.min(headers.length,Math.max(1,sh.getLastColumn())),existing=sh.getRange(1,1,1,used).getDisplayValues()[0];
  if(existing.every(function(v){return !v;}))sh.getRange(1,1,1,headers.length).setValues([headers]);
  else{if(existing.some(function(h,i){return h!==headers[i];}))throw new Error('Header tab '+name+' tidak sepadan. Tiada data diubah.');if(used<headers.length)sh.getRange(1,used+1,1,headers.length-used).setValues([headers.slice(used)]);}return sh;
}
function upsertSeed_(sheetName,keyName,records){
  const sh=getSheet_(sheetName),headers=HEADERS[Object.keys(CONFIG.SHEETS).find(function(k){return CONFIG.SHEETS[k]===sheetName;})],keyIndex=headers.indexOf(keyName),values=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],known={};
  values.forEach(function(row){known[String(row[keyIndex])]=true;});const add=records.filter(function(r){return !known[String(r[keyName])];}).map(function(r){return headers.map(function(h){return Object.prototype.hasOwnProperty.call(r,h)?r[h]:'';});});if(add.length)sh.getRange(sh.getLastRow()+1,1,add.length,headers.length).setValues(add);
}
function formatSheets_(ss){Object.keys(CONFIG.SHEETS).forEach(function(k){const sh=ss.getSheetByName(CONFIG.SHEETS[k]);if(!sh)return;const cols=HEADERS[k].length;sh.setFrozenRows(1);sh.getRange(1,1,1,cols).setBackground('#10233f').setFontColor('#ffffff').setFontWeight('bold');sh.autoResizeColumns(1,cols);});}

function bootstrap_(){const judges=getActiveJudges_(),posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES);return {judges:judges.map(function(j){const p=progressFor_(j.id,posters,scores);return {id:j.id,name:j.name,completed:p.completed,total:p.total};}),posterCount:posters.length,ready:true};}
function login_(req){const judge=getActiveJudges_().find(function(j){return j.id===String(req.judgeId||'');});if(!judge)throw new Error('Hakim tidak sah atau tidak aktif.');if(!verifyPin_(judge.id,String(req.pin||'')))throw new Error('PIN tidak tepat. Sila cuba semula.');const token=createToken_(judge.id);return {token:token,judge:{id:judge.id,name:judge.name},dashboard:dashboardForJudge_(judge,String(req.preferredPosterId||''),false,'')};}
function dashboard_(req){return dashboardForJudge_(authenticatedJudge_(req.token),String(req.preferredPosterId||''),false,'');}
function selectPoster_(req){const judge=authenticatedJudge_(req.token),posterId=String(req.posterId||'');if(!posterId)throw new Error('ID poster diperlukan.');return dashboardForJudge_(judge,posterId,true,posterId);}
function dashboardForJudge_(judge,preferredPosterId,includeJudgedPreferred,afterPosterId){const posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES),mine=scores.filter(function(x){return String(x['ID Hakim'])===judge.id;}),judged={};mine.forEach(function(x){judged[String(x['ID Poster'])]=x;});let remaining=posters.filter(function(p){return !judged[p.id];});const orderAfter=function(items,id){const anchor=posters.findIndex(function(p){return p.id===id;});if(anchor<0)return items;const after=[],before=[];items.forEach(function(p){(posters.findIndex(function(x){return x.id===p.id;})>anchor?after:before).push(p);});return after.concat(before);};if(afterPosterId)remaining=orderAfter(remaining,afterPosterId);const preferred=posters.find(function(p){return p.id===preferredPosterId;}),selected=preferred&&(includeJudgedPreferred||!judged[preferred.id])?preferred:(remaining[0]||null),nextPool=selected?orderAfter(remaining.filter(function(p){return p.id!==selected.id;}),selected.id):remaining,progress=progressFor_(judge.id,posters,scores),last=mine.reduce(function(max,x){const d=new Date(x['Tarikh Kemas Kini']);return d>max?d:max;},new Date(0));if(preferredPosterId&&!preferred)throw new Error('Poster yang dipilih tidak sah atau tidak aktif.');return {judge:{id:judge.id,name:judge.name},poster:selected?posterPayload_(selected,judged[selected.id]||null):null,nextPoster:nextPool[0]?posterPayload_(nextPool[0],judged[nextPool[0].id]||null):null,progress:progress,lastUpdated:last.getTime()?last:null};}
function saveScore_(req){
  const judge=authenticatedJudge_(req.token);if(String(getSetting_('PENJURIAN_DIBUKA','TRUE')).toUpperCase()!=='TRUE')throw new Error('Penjurian sedang ditutup oleh urusetia.');
  const posters=getActivePosters_(),poster=posters.find(function(p){return p.id===String(req.posterId||'');});if(!poster)throw new Error('ID poster tidak sah atau tidak aktif.');const clean=validateScores_(req.scores||{}),total=Object.keys(clean).reduce(function(a,k){return a+clean[k];},0),requestId=String(req.requestId||Utilities.getUuid()),lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const sh=getSheet_(CONFIG.SHEETS.SCORES),headers=HEADERS.SCORES,all=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],key=judge.id+'::'+poster.id,keyCol=headers.indexOf('Kunci Rekod'),requestCol=headers.indexOf('Request ID'),existingIndex=all.findIndex(function(r){return String(r[keyCol])===key;}),duplicate=all.find(function(r){return String(r[requestCol])===requestId;});if(duplicate)return dashboardForJudge_(judge,'',false,poster.id);
    const now=new Date(),old=existingIndex>=0?rowObject_(headers,all[existingIndex]):null;if(old&&String(getSetting_('EDIT_MARKAH_DIBENARKAN','TRUE')).toUpperCase()!=='TRUE')throw new Error('Markah sedia ada tidak boleh diedit.');
    const record={'Timestamp':old?old.Timestamp:now,'ID Hakim':judge.id,'Nama Hakim':judge.name,'ID Poster':poster.id,'Nama Poster':poster.title,'Nama Sekolah':poster.school,'Kandungan & Kaitan Tema':clean.content,'Reka Bentuk & Kreativiti':clean.design,'Keberkesanan & Kejelasan Maklumat':clean.clarity,'Kualiti Teknikal':clean.technical,'Potensi & Impak':clean.impact,'Jumlah':total,'Status':'DISIMPAN','Tarikh Kemas Kini':now,'Kunci Rekod':key,'Versi Rekod':old?(Number(old['Versi Rekod'])||1)+1:1,'Request ID':requestId,'Catatan DQ':''},row=headers.map(function(h){return record[h];});
    if(existingIndex>=0)sh.getRange(existingIndex+2,1,1,headers.length).setValues([row]);else sh.appendRow(row);appendAudit_(judge.id,poster.id,old?'KEMAS_KINI':'CIPTA',old,record,requestId);SpreadsheetApp.flush();return dashboardForJudge_(judge,'',false,poster.id);
  }finally{lock.releaseLock();}
}
function saveDisqualification_(req){
  const judge=authenticatedJudge_(req.token);if(String(getSetting_('PENJURIAN_DIBUKA','TRUE')).toUpperCase()!=='TRUE')throw new Error('Penjurian sedang ditutup oleh urusetia.');
  const poster=getActivePosters_().find(function(p){return p.id===String(req.posterId||'');}),note=String(req.note||'').trim(),requestId=String(req.requestId||Utilities.getUuid());if(!poster)throw new Error('ID poster tidak sah atau tidak aktif.');if(note.length<5)throw new Error('Sila nyatakan catatan DQ yang jelas.');if(note.length>500)throw new Error('Catatan DQ tidak boleh melebihi 500 aksara.');const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{const sh=getSheet_(CONFIG.SHEETS.SCORES),headers=HEADERS.SCORES,all=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],key=judge.id+'::'+poster.id,keyCol=headers.indexOf('Kunci Rekod'),requestCol=headers.indexOf('Request ID'),existingIndex=all.findIndex(function(r){return String(r[keyCol])===key;});if(all.some(function(r){return String(r[requestCol])===requestId;}))return dashboardForJudge_(judge,'',false,poster.id);const now=new Date(),old=existingIndex>=0?rowObject_(headers,all[existingIndex]):null,record={'Timestamp':old?old.Timestamp:now,'ID Hakim':judge.id,'Nama Hakim':judge.name,'ID Poster':poster.id,'Nama Poster':poster.title,'Nama Sekolah':poster.school,'Kandungan & Kaitan Tema':'','Reka Bentuk & Kreativiti':'','Keberkesanan & Kejelasan Maklumat':'','Kualiti Teknikal':'','Potensi & Impak':'','Jumlah':'','Status':'DQ','Tarikh Kemas Kini':now,'Kunci Rekod':key,'Versi Rekod':old?(Number(old['Versi Rekod'])||1)+1:1,'Request ID':requestId,'Catatan DQ':note},row=headers.map(function(h){return record[h];});if(existingIndex>=0)sh.getRange(existingIndex+2,1,1,headers.length).setValues([row]);else sh.appendRow(row);appendAudit_(judge.id,poster.id,old?'TANDA_DQ_KEMAS_KINI':'TANDA_DQ',old,record,requestId);SpreadsheetApp.flush();return dashboardForJudge_(judge,'',false,poster.id);}finally{lock.releaseLock();}
}
function gallery_(req){const judge=authenticatedJudge_(req.token),posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES),mine={};scores.filter(function(x){return String(x['ID Hakim'])===judge.id;}).forEach(function(x){mine[String(x['ID Poster'])]=x;});let items=posters.map(function(p){return posterPayload_(p,mine[p.id]||null);}),status=String(req.status||'all'),query=String(req.query||'').toLowerCase(),category=String(req.category||'all');if(status==='judged')items=items.filter(function(x){return x.judged&&!x.disqualifiedByMe;});if(status==='unjudged')items=items.filter(function(x){return !x.judged;});if(status==='dq')items=items.filter(function(x){return x.disqualifiedByMe;});if(query)items=items.filter(function(x){return (x.title+' '+x.school+' '+x.id).toLowerCase().indexOf(query)>=0;});if(category!=='all')items=items.filter(function(x){return x.category===category;});const size=Math.min(12,Math.max(1,Number(req.pageSize)||CONFIG.GALLERY_PAGE_SIZE)),pages=Math.max(1,Math.ceil(items.length/size)),page=Math.min(pages,Math.max(1,Number(req.page)||1));return {items:items.slice((page-1)*size,page*size),page:page,pages:pages,total:items.length,categories:Array.from(new Set(posters.map(function(p){return p.category;}).filter(Boolean))).sort()};}
function leaderboard_(req){return summaryForJudge_(authenticatedJudge_(req.token));}
function summary_(req){return summaryForJudge_(authenticatedJudge_(req.token));}
function adminLeaderboard_(req){authenticatedAdmin_(req.token);return leaderboardData_();}
function publicLeaderboard_(){const data=leaderboardData_();if(!data.complete)return data;return {complete:true,items:data.items.map(function(x){const copy={id:x.id,participant:x.participant,title:x.title,school:x.school,judgeTotals:x.disqualified?[null,null,null]:x.judgeTotals,total:x.total,rank:x.rank,award:x.award,disqualified:x.disqualified};if(x.disqualified)copy.publicNote='Tidak memenuhi syarat pertandingan.';return copy;})};}
function summaryForJudge_(judge){const posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES),mine={};scores.filter(function(s){return String(s['ID Hakim'])===judge.id;}).forEach(function(s){mine[String(s['ID Poster'])]=s;});const dq=disqualificationMap_(scores);let items=posters.filter(function(p){return !!mine[p.id];}).map(function(p){const row=mine[p.id],global=dq[p.id]||null,ownDQ=isDQ_(row);return {id:p.id,number:p.number,title:p.title,school:p.school,status:global?'DQ':'DINILAI',ownDQ:ownDQ,total:ownDQ?null:Number(row.Jumlah)||0,updatedAt:row['Tarikh Kemas Kini'],dqNotes:global?global.notes:[]};});const ranked=items.filter(function(x){return x.status!=='DQ';}).sort(function(a,b){return b.total-a.total||Number(a.number)-Number(b.number)||a.id.localeCompare(b.id);});ranked.forEach(function(x,i){x.rank=i+1;});const disqualified=items.filter(function(x){return x.status==='DQ';}).sort(function(a,b){return Number(a.number)-Number(b.number)||a.id.localeCompare(b.id);});return {complete:true,personal:true,completed:items.length,total:posters.length,items:ranked.concat(disqualified)};}
function leaderboardData_(){const judges=getActiveJudges_(),posters=getActivePosters_(),scores=getRows_(CONFIG.SHEETS.SCORES),progress=judges.map(function(j){const p=progressFor_(j.id,posters,scores);return {completed:p.completed,total:p.total};});if(progress.some(function(p){return p.completed<p.total;}))return {complete:false,judges:progress};const map={},dq=disqualificationMap_(scores);scores.forEach(function(s){const pid=String(s['ID Poster']),jid=String(s['ID Hakim']);if(!map[pid])map[pid]={};map[pid][jid]=s;});let items=posters.map(function(p){const by=map[p.id]||{},judgeTotals=judges.map(function(j){const row=by[j.id];return row&&!isDQ_(row)?Number(row.Jumlah)||0:null;}),aspect=function(header){return judges.reduce(function(a,j){const row=by[j.id];return a+(row&&!isDQ_(row)?Number(row[header])||0:0);},0);},global=dq[p.id]||null;return {id:p.id,participant:p.participant,title:p.title,school:p.school,judgeTotals:judgeTotals,total:judgeTotals.reduce(function(a,b){return a+(b||0);},0),impact:aspect(SCORE_KEYS.impact),content:aspect(SCORE_KEYS.content),design:aspect(SCORE_KEYS.design),disqualified:!!global,dqNotes:global?global.notes:[]};}),eligible=items.filter(function(x){return !x.disqualified;}),disqualified=items.filter(function(x){return x.disqualified;});eligible.sort(function(a,b){return b.total-a.total||b.impact-a.impact||b.content-a.content||b.design-a.design||a.id.localeCompare(b.id);});const qP=Number(getSetting_('KUOTA_PLATINUM',5)),qE=Number(getSetting_('KUOTA_EMAS',36)),qS=Number(getSetting_('KUOTA_PERAK',63)),allAward=String(getSetting_('SEMUA_TERIMA_ANUGERAH','TRUE')).toUpperCase()==='TRUE';eligible.forEach(function(x,i){x.rank=i+1;x.award=i<qP?'Platinum':i<qP+qE?'Emas':i<qP+qE+qS?'Perak':allAward?'Gangsa':'Tiada Anugerah';delete x.impact;delete x.content;delete x.design;});disqualified.forEach(function(x){x.rank=null;x.award='Anugerah Penyertaan';delete x.impact;delete x.content;delete x.design;});return {complete:true,items:eligible.concat(disqualified)};}
function isDQ_(row){return String((row||{}).Status||'').toUpperCase()==='DQ';}
function disqualificationMap_(scores){const out={};scores.forEach(function(row){if(!isDQ_(row))return;const id=String(row['ID Poster']);if(!out[id])out[id]={notes:[]};out[id].notes.push({judgeId:String(row['ID Hakim']),judgeName:String(row['Nama Hakim']),note:String(row['Catatan DQ']||''),updatedAt:row['Tarikh Kemas Kini']});});return out;}

function authenticatedJudge_(token){const payload=verifyToken_(String(token||'')),judge=getActiveJudges_().find(function(j){return j.id===payload.judgeId;});if(!judge)throw new Error('Sesi hakim tidak sah. Sila log masuk semula.');return judge;}
function adminLogin_(req){const props=PropertiesService.getScriptProperties(),salt=props.getProperty('ADMIN_PIN_SALT'),expected=props.getProperty('ADMIN_PIN_HASH');if(!salt||!expected)throw new Error('PIN urusetia belum dikonfigurasi.');if(hash_(salt+':'+String(req.pin||''))!==expected)throw new Error('PIN urusetia tidak tepat.');return {token:createAdminToken_()};}
function authenticatedAdmin_(token){const payload=verifyToken_(String(token||''));if(payload.role!=='admin')throw new Error('Sesi urusetia tidak sah. Sila log masuk semula.');return true;}
function getActiveJudges_(){return getRows_(CONFIG.SHEETS.JUDGES).filter(function(r){return isActive_(r['Status Aktif']);}).map(function(r){return {id:String(r['ID Hakim']),name:String(r['Nama Hakim']),order:Number(r['Susunan Hakim'])||999};}).sort(function(a,b){return a.order-b.order;});}
function getActivePosters_(){return getRows_(CONFIG.SHEETS.POSTERS).filter(function(r){return isActive_(r['Status Aktif']);}).map(function(r){const link=String(r['Link Poster']||''),imageUrls=driveImages_(link,1600),thumbnailUrls=driveImages_(link,600);return {id:String(r['ID Poster']),number:r['No. Giliran'],participant:String(r['Nama Peserta']||''),school:String(r['Nama Sekolah']),title:String(r['Nama Poster']),link:link,imageUrl:imageUrls[0]||link,thumbnailUrl:thumbnailUrls[0]||link,imageUrls:imageUrls,thumbnailUrls:thumbnailUrls,category:String(r.Kategori||'')};}).sort(function(a,b){return Number(a.number)-Number(b.number)||a.id.localeCompare(b.id);});}
function posterPayload_(p,score){const out={id:p.id,number:p.number,school:p.school,title:p.title,link:p.link,imageUrl:p.imageUrl,thumbnailUrl:p.thumbnailUrl,imageUrls:p.imageUrls,thumbnailUrls:p.thumbnailUrls,category:p.category,judged:!!score,disqualifiedByMe:isDQ_(score)};if(score){if(!out.disqualifiedByMe){out.scores={content:Number(score[SCORE_KEYS.content]),design:Number(score[SCORE_KEYS.design]),clarity:Number(score[SCORE_KEYS.clarity]),technical:Number(score[SCORE_KEYS.technical]),impact:Number(score[SCORE_KEYS.impact])};out.total=Number(score.Jumlah);}else out.dqNote=String(score['Catatan DQ']||'');out.updatedAt=score['Tarikh Kemas Kini'];}return out;}
function progressFor_(judgeId,posters,scores){const active={};posters.forEach(function(p){active[p.id]=true;});const done={};scores.forEach(function(s){if(String(s['ID Hakim'])===judgeId&&active[String(s['ID Poster'])])done[String(s['ID Poster'])]=true;});return {completed:Object.keys(done).length,total:posters.length};}
function validateScores_(scores){const out={};Object.keys(SCORE_KEYS).forEach(function(k){const n=Number(scores[k]);if(!Number.isInteger(n)||n<1||n>20)throw new Error('Semua lima aspek mestilah markah bulat antara 1 hingga 20.');out[k]=n;});return out;}
function appendAudit_(judgeId,posterId,action,oldRow,newRow,requestId){const oldScores=oldRow?scoreSnapshot_(oldRow):'',newScores=scoreSnapshot_(newRow),row=[new Date(),Utilities.getUuid(),judgeId,posterId,action,oldScores?JSON.stringify(oldScores):'',JSON.stringify(newScores),oldRow?Number(oldRow.Jumlah)||0:'',Number(newRow.Jumlah)||0,requestId];getSheet_(CONFIG.SHEETS.AUDIT).appendRow(row);}
function scoreSnapshot_(r){return {content:Number(r[SCORE_KEYS.content])||null,design:Number(r[SCORE_KEYS.design])||null,clarity:Number(r[SCORE_KEYS.clarity])||null,technical:Number(r[SCORE_KEYS.technical])||null,impact:Number(r[SCORE_KEYS.impact])||null,status:String(r.Status||''),dqNote:String(r['Catatan DQ']||'')};}
function getRows_(name){const sh=getSheet_(name),last=sh.getLastRow(),cols=sh.getLastColumn();if(last<2)return [];const values=sh.getRange(1,1,last,cols).getValues(),headers=values.shift();return values.filter(function(row){return row.some(function(v){return v!=='';});}).map(function(row){return rowObject_(headers,row);});}
function rowObject_(headers,row){const o={};headers.forEach(function(h,i){o[h]=row[i];});return o;}
function getSheet_(name){const sh=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(name);if(!sh)throw new Error('Tab '+name+' belum disediakan.');return sh;}
function getSetting_(key,fallback){const row=getRows_(CONFIG.SHEETS.SETTINGS).find(function(r){return String(r.Kunci)===key;});return row?row.Nilai:fallback;}
function isActive_(v){return v===true||['TRUE','YA','AKTIF','1'].indexOf(String(v).toUpperCase())>=0;}
function driveImages_(link,size){const m=String(link).match(/\/d\/([a-zA-Z0-9_-]+)/)||String(link).match(/[?&]id=([a-zA-Z0-9_-]+)/);if(!m)return link?[link]:[];const id=m[1];return ['https://drive.google.com/thumbnail?id='+id+'&sz=w'+size,'https://lh3.googleusercontent.com/d/'+id+'=w'+size,'https://drive.usercontent.google.com/download?id='+id+'&export=view&authuser=0'];}

function verifyPin_(judgeId,pin){const props=PropertiesService.getScriptProperties(),salt=props.getProperty('PIN_SALT'),expected=props.getProperty('PIN_HASH_'+judgeId);if(!salt||!expected)throw new Error('PIN hakim belum dikonfigurasi oleh urusetia.');return hash_(salt+':'+pin)===expected;}
function createToken_(judgeId){const props=PropertiesService.getScriptProperties(),secret=props.getProperty('SESSION_SECRET');if(!secret)throw new Error('Rahsia sesi belum dikonfigurasi.');const payload=Utilities.base64EncodeWebSafe(JSON.stringify({judgeId:judgeId,exp:Date.now()+CONFIG.SESSION_HOURS*3600000}),Utilities.Charset.UTF_8).replace(/=+$/,'');return payload+'.'+sign_(payload,secret);}
function createAdminToken_(){const props=PropertiesService.getScriptProperties(),secret=props.getProperty('SESSION_SECRET');if(!secret)throw new Error('Rahsia sesi belum dikonfigurasi.');const payload=Utilities.base64EncodeWebSafe(JSON.stringify({role:'admin',exp:Date.now()+CONFIG.SESSION_HOURS*3600000}),Utilities.Charset.UTF_8).replace(/=+$/,'');return payload+'.'+sign_(payload,secret);}
function verifyToken_(token){const parts=token.split('.');if(parts.length!==2)throw new Error('Sesi telah tamat. Sila log masuk semula.');const secret=PropertiesService.getScriptProperties().getProperty('SESSION_SECRET');if(!secret||sign_(parts[0],secret)!==parts[1])throw new Error('Sesi tidak sah. Sila log masuk semula.');let data;try{data=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());}catch(e){throw new Error('Sesi tidak sah.');}if(!data.exp||Date.now()>data.exp)throw new Error('Sesi telah tamat. Sila log masuk semula.');return data;}
function sign_(value,secret){return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value,secret)).replace(/=+$/,'');}
function hash_(value){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value,Utilities.Charset.UTF_8).map(function(b){return ('0'+(b<0?b+256:b).toString(16)).slice(-2);}).join('');}
