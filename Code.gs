const EXP='Financieel_Kosten', INC='Financieel_Inkomsten', BAL='Financieel_Saldi';
const M=['Jan','Feb','Mrt','Apr','Mei','Jun','Jul','Aug','Sep','Okt','Nov','Dec'];
function doGet(e){
  try {
    const action=e.parameter.action||'getState';
    const payload=e.parameter.data?JSON.parse(e.parameter.data):{};
    const result=route_(action,payload);
    return output_({ok:true,data:result},e.parameter.callback);
  } catch(err){ return output_({ok:false,error:String(err.message||err)},e.parameter.callback); }
}
function route_(a,d){
  guard_();
  if(a==='getState') return state_(Number(d.year||new Date().getFullYear()));
  if(a==='saveExpense') return save_(EXP,d.row,'expense');
  if(a==='saveIncome') return save_(INC,d.row,'income');
  if(a==='deleteExpense') return remove_(EXP,d.year,d.id);
  if(a==='deleteIncome') return remove_(INC,d.year,d.id);
  if(a==='saveBalance') return balance_(d.year,d.month,d.amount);
  if(a==='createYear') return newYear_(Number(d.fromYear),Number(d.toYear));
  throw new Error('Onbekende actie: '+a);
}
function setupFinancieel(){
  const ss=SpreadsheetApp.getActive();
  ensure_(ss,EXP,['year','id','category','description','type',...M,'active']);
  ensure_(ss,INC,['year','id','description',...M,'active']);
  ensure_(ss,BAL,['year','month','closingBalance']);
  PropertiesService.getScriptProperties().setProperty('OWNER',Session.getEffectiveUser().getEmail());
  SpreadsheetApp.getUi().alert('Tabellen aangemaakt. Importeer nu je oorspronkelijke Excel in deze spreadsheet en voer daarna importeerExcel2026() uit.');
}
function importeerExcel2026(){
  const ss=SpreadsheetApp.getActive();
  const costs=ss.getSheetByName('Kosten 2026'), summary=ss.getSheetByName('Samenvatting 2026'), oldBal=ss.getSheetByName('SaldoOverzicht');
  if(!costs||!summary) throw new Error('De tabbladen Kosten 2026 en Samenvatting 2026 zijn niet gevonden. Importeer eerst je Excelbestand in Google Sheets.');
  [EXP,INC,BAL].forEach(n=>{const s=ss.getSheetByName(n);if(s&&s.getLastRow()>1)s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).clearContent();});
  const expenseData=costs.getDataRange().getValues(), ex=[];
  for(let i=3;i<expenseData.length;i++){
    const r=expenseData[i], cat=r[1], desc=r[2];
    if(!cat||!desc||String(desc).toLowerCase()==='totaal')continue;
    const vals=r.slice(3,15).map(v=>v===''||v==null?'':Number(v));
    const filled=vals.filter(v=>v!=='');
    const same=filled.length>=10&&filled.every(v=>Math.abs(v-filled[0])<0.005);
    ex.push([2026,Utilities.getUuid(),String(cat).trim(),String(desc).trim(),same?'Vast maandelijks':'Variabel / gepland',...vals,true]);
  }
  if(ex.length)ss.getSheetByName(EXP).getRange(2,1,ex.length,ex[0].length).setValues(ex);
  const incData=summary.getDataRange().getValues(), ins=[];
  // Inkomsten staan in rijen 8 t/m 11; naam in kolom B, maanden C:N.
  for(let i=7;i<11;i++){const r=incData[i],name=r[1];if(!name)continue;ins.push([2026,Utilities.getUuid(),String(name).trim(),...r.slice(2,14).map(v=>v===''||v==null?'':Number(v)),true]);}
  if(ins.length)ss.getSheetByName(INC).getRange(2,1,ins.length,ins[0].length).setValues(ins);
  if(oldBal){const b=[];const rows=oldBal.getDataRange().getValues();for(let i=1;i<rows.length;i++){const dt=rows[i][0],amount=rows[i][1];if(!dt||amount===''||amount==null)continue;let y,m;if(dt instanceof Date){y=dt.getFullYear();m=dt.getMonth()+1;}else{const z=String(dt).match(/(januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december)\s+(\d{4})/i);if(z){y=+z[2];m=['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'].indexOf(z[1].toLowerCase())+1;}}if(y&&m)b.push([y,m,Number(amount)]);}if(b.length)ss.getSheetByName(BAL).getRange(2,1,b.length,3).setValues(b);}
  SpreadsheetApp.getUi().alert(ex.length+' onkostenregels en '+ins.length+' inkomstenregels geïmporteerd.');
}
function state_(year){
 const ss=SpreadsheetApp.getActive();
 return {year:year,months:M,expenses:read_(ss.getSheetByName(EXP)).filter(x=>+x.year===year&&x.active!==false),incomes:read_(ss.getSheetByName(INC)).filter(x=>+x.year===year&&x.active!==false),balances:read_(ss.getSheetByName(BAL)).filter(x=>+x.year===year)};
}
function read_(s){if(!s||s.getLastRow()<2)return[];const a=s.getDataRange().getValues(),h=a[0];return a.slice(1).filter(r=>r.some(v=>v!==''&&v!=null)).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);if(h.includes('Jan'))o.months=M.map(m=>o[m]===''||o[m]==null?null:Number(o[m]));return o;});}
function save_(sheetName,row,kind){
 if(!row||!row.id||!row.year)throw new Error('Ontbrekende gegevens.');
 const s=SpreadsheetApp.getActive().getSheetByName(sheetName);const values=kind==='expense'?[+row.year,row.id,row.category||'',row.description||'',row.type||'Variabel / gepland',...M.map((m,i)=>num_(row.months&&row.months[i])),true]:[+row.year,row.id,row.description||'',...M.map((m,i)=>num_(row.months&&row.months[i])),true];
 const a=s.getDataRange().getValues();for(let i=1;i<a.length;i++)if(+a[i][0]===+row.year&&String(a[i][1])===String(row.id)){s.getRange(i+1,1,1,values.length).setValues([values]);return{saved:true};}
 s.appendRow(values);return{saved:true};
}
function num_(v){return v===''||v==null?'':Number(v);}
function remove_(sheetName,year,id){const s=SpreadsheetApp.getActive().getSheetByName(sheetName),a=s.getDataRange().getValues();for(let i=1;i<a.length;i++)if(+a[i][0]===+year&&String(a[i][1])===String(id)){s.deleteRow(i+1);break;}return{deleted:true};}
function balance_(year,month,amount){const s=SpreadsheetApp.getActive().getSheetByName(BAL),a=s.getDataRange().getValues();for(let i=1;i<a.length;i++)if(+a[i][0]===+year&&+a[i][1]===+month){s.getRange(i+1,3).setValue(Number(amount));return{saved:true};}s.appendRow([+year,+month,Number(amount)]);return{saved:true};}
function newYear_(from,to){
 if(!from||!to||to<=from)throw new Error('Ongeldig boekjaar.');
 const ss=SpreadsheetApp.getActive();
 const es=ss.getSheetByName(EXP),er=es.getDataRange().getValues(),ec=[];
 er.slice(1).filter(r=>+r[0]===from&&r[r.length-1]!==false).forEach(r=>{const fixed=r[4]==='Vast maandelijks';ec.push([to,Utilities.getUuid(),r[2],r[3],r[4],...r.slice(5,17).map(v=>fixed?v:''),true]);});
 if(ec.length)es.getRange(es.getLastRow()+1,1,ec.length,ec[0].length).setValues(ec);
 const is=ss.getSheetByName(INC),ir=is.getDataRange().getValues(),ic=[];
 ir.slice(1).filter(r=>+r[0]===from&&r[r.length-1]!==false).forEach(r=>ic.push([to,Utilities.getUuid(),r[2],...Array(12).fill(''),true]));
 if(ic.length)is.getRange(is.getLastRow()+1,1,ic.length,ic[0].length).setValues(ic);
 return{created:true,year:to,expenses:ec.length,incomes:ic.length};
}
function ensure_(ss,n,h){let s=ss.getSheetByName(n);if(!s)s=ss.insertSheet(n);if(s.getLastRow()===0)s.appendRow(h);}
function guard_(){
 const owner=PropertiesService.getScriptProperties().getProperty('OWNER');
 const active=Session.getActiveUser().getEmail();
 if(!owner||!active||owner.toLowerCase()!==active.toLowerCase())throw new Error('Geen toegang. Open de toepassing met het Google-account van de eigenaar.');
}
function output_(obj,cb){const body=JSON.stringify(obj).replace(/</g,'\\u003c');if(cb)return ContentService.createTextOutput(cb+'('+body+');').setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);}
