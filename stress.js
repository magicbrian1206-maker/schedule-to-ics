// 壓力測試：用與網頁完全相同的 ExcelJS 邏輯，為每個人產生「隨機班表」Excel 到桌面
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const SRC = '/Users/Brian/Desktop/FY26Q3 W10-W13 ASC Scheduling.xlsx';
const OUT_DIR = '/Users/Brian/Desktop/壓力測試_班表產出';
const FILE_NAME = path.basename(SRC);

// ===== 複製網頁常數/helper =====
const pad = n => String(n).padStart(2,'0');
const ymdDash = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const ALLOWED_FIRST = new Set(['Brian','Adam','Kay','Vic','Haidee','Daisy','Elay','Emily','Jessie','Vincent','Kristin','Paty']);
const NAME_ALIAS = {'Chou Meng Hua':'Haidee'};
const EARLY_FIRST = new Set(['Vic','Brian','Vincent','Jessie','Daisy','Kristin']);
const NEEDS_TIME = new Set(['All day in store','AM - Meeting/Training, PM - in store','AM - in store, PM - Meeting/Training','All day Meeting/Training']);
function defaultTime(name){const first=String(name||'').split(/\s+/)[0];return EARLY_FIRST.has(first)?{start:'12:00',end:'21:00'}:{start:'12:30',end:'21:30'};}
function parseWeekNum(v){if(v==null)return null;const m=String(v).match(/(\d+)/);return m?parseInt(m[1],10):null;}
function parseFilenameWeeks(name){if(!name)return null;let m=name.match(/W(\d+)\s*[-~–—]\s*W?(\d+)/i);if(m){const a=+m[1],b=+m[2],lo=Math.min(a,b),hi=Math.max(a,b),arr=[];for(let w=lo;w<=hi;w++)arr.push(w);return arr;}m=name.match(/W(\d+)/i);return m?[+m[1]]:null;}
function excelDate(v){if(v instanceof Date)return v;if(typeof v==='number')return new Date(Math.round((v-25569)*86400*1000));if(v&&v.result instanceof Date)return v.result;return null;}
function detectColumns(aoa){
  for(let r=0;r<Math.min(5,aoa.length);r++){
    const row=aoa[r]||[];const idx={name:-1,date:-1,status:-1,week:-1};
    row.forEach((v,i)=>{const s=String(v||'').toLowerCase().trim();
      if(idx.name<0&&(s==='eng name'||s==='name'||s==='姓名'||s==='員工'||s==='人員'))idx.name=i;
      if(idx.date<0&&(s==='date'||s==='日期'))idx.date=i;
      if(idx.status<0&&(s==='status'||s==='假別'||s==='狀態'||s==='請假'))idx.status=i;
      if(idx.week<0&&(s==='week'||s==='週次'||s==='周次'))idx.week=i;});
    if(idx.name>=0&&idx.date>=0&&idx.status>=0){if(idx.week<0)idx.week=8;return{...idx,headerRow:r};}
  }
  return {name:4,date:9,status:13,week:8,headerRow:1};
}

// 隨機狀態池（含工作/會議/各種假別）
const RAND_STATUSES = ['All day in store','All day in store','All day in store',
  'AM - Meeting/Training, PM - in store','AM - in store, PM - Meeting/Training','All day Meeting/Training',
  'OFF','AL','SL','M','PH11','RPH'];
const rnd = arr => arr[Math.floor(Math.random()*arr.length)];

(async ()=>{
  // 1) 用 XLSX 解析範本（同網頁 loadTemplate）
  const buf = fs.readFileSync(SRC);
  const wbX = XLSX.read(buf,{cellDates:true});
  let pickAoa=null,pickCol=null,pickSheet=null,best=0;
  for(const sn of wbX.SheetNames){
    const aoa=XLSX.utils.sheet_to_json(wbX.Sheets[sn],{header:1,defval:null,raw:true});
    const col=detectColumns(aoa);let cnt=0;
    for(let i=col.headerRow+1;i<aoa.length;i++){const r=aoa[i]||[];if(r[col.name]&&r[col.date])cnt++;}
    if(cnt>best){best=cnt;pickAoa=aoa;pickCol=col;pickSheet=sn;}
  }
  const rows=[];
  for(let i=pickCol.headerRow+1;i<pickAoa.length;i++){
    const r=pickAoa[i]||[];const name=r[pickCol.name],date=excelDate(r[pickCol.date]);
    if(!name||!date)continue;
    rows.push({name:String(name).trim(),date,week:parseWeekNum(r[pickCol.week]),
      status:r[pickCol.status]||'',o:r[14],q:r[16],rh:r[17],t:r[19],l:r[11]});
  }
  const people=[...new Set(rows.map(r=>r.name))].sort();
  const weeks=[...new Set(rows.map(r=>r.week).filter(w=>w!=null))].sort((a,b)=>a-b);
  const fnWeeks=parseFilenameWeeks(FILE_NAME);
  const selectedWeeks=new Set((fnWeeks&&fnWeeks.filter(w=>weeks.includes(w)).length?fnWeeks.filter(w=>weeks.includes(w)):weeks));
  const sel=rows.filter(r=>selectedWeeks.has(r.week));
  const selectedDateSet=new Set(sel.map(r=>ymdDash(r.date)));

  // 抓店清單供隨機指派（壓測店號寫入路徑）
  const storeList=[];
  for(const sn of wbX.SheetNames){
    if(sn===pickSheet)continue;
    const aoa=XLSX.utils.sheet_to_json(wbX.Sheets[sn],{header:1,defval:null,raw:true});
    aoa.forEach(r=>{if(r&&typeof r[0]==='number'&&r[1])storeList.push(String(r[0]));});
  }
  console.log('店清單數:',storeList.length);

  const allowedPeople = people.filter(n=>{const first=NAME_ALIAS[n]||n.split(/\s+/)[0];return ALLOWED_FIRST.has(first);});

  console.log('範本工作表:',pickSheet,'| 欄位:',JSON.stringify(pickCol));
  console.log('檔案週次:',weeks.join(','),'| 預設勾選:',[...selectedWeeks].join(','));
  console.log('選取日期數:',selectedDateSet.size,'| 名單:',allowedPeople.length,'人');
  console.log('人員:',allowedPeople.map(n=>NAME_ALIAS[n]||n.split(/\s+/)[0]).join(', '));

  if(!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR,{recursive:true});

  const results=[];
  for(const p of allowedPeople){
    // 2) 建立隨機 editState（同網頁 loadExistingIntoState + 使用者隨機改動）
    const dt=defaultTime(p);
    const storeFreq={};
    rows.filter(r=>r.name===p && r.l).forEach(r=>{storeFreq[r.l]=(storeFreq[r.l]||0)+1});
    const defaultStore=Object.keys(storeFreq).sort((a,b)=>storeFreq[b]-storeFreq[a])[0]||'';
    const editState={};
    rows.filter(r=>r.name===p && selectedWeeks.has(r.week)).forEach(r=>{
      const k=`${p}|${ymdDash(r.date)}`;
      const baseStore=r.l?String(r.l):defaultStore;
      // 隨機指派一個狀態（壓力測試重點：每天都隨機）
      let status=rnd(RAND_STATUSES);
      let start=dt.start,end=dt.end;
      editState[k]={status,start,end,storeId:baseStore,remark: status==='AL'?'壓測備註':''};
    });

    // 3) 用 ExcelJS 產出（完全複製網頁 t2Export）
    const wb=new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws=wb.getWorksheet(pickSheet);
    const colName=pickCol.name+1, colDate=pickCol.date+1;
    const colL=12,colM=13,colN=14,colO=15,colQ=17,colR_=18,colT=20,colW=23;
    let modifiedCount=0, firstRow=null;
    ws.eachRow({includeEmpty:false},(row,rNum)=>{
      if(rNum<4)return;
      const nv=row.getCell(colName).value; if(!nv)return;
      if(String(nv).trim()!==p)return;
      const date=excelDate(row.getCell(colDate).value);
      if(!date)return;
      if(!selectedDateSet.has(ymdDash(date)))return;
      const st=editState[`${p}|${ymdDash(date)}`]; if(!st)return;
      if(firstRow===null)firstRow=rNum;
      row.getCell(colN).value=st.status;
      if(st.storeId){const idNum=Number(st.storeId);row.getCell(colL).value=isNaN(idNum)?st.storeId:idNum;}else{row.getCell(colL).value=null;}
      row.getCell(colL).alignment={horizontal:'center',vertical:'middle'};
      if(NEEDS_TIME.has(st.status)){
        const [oH,qM]=st.start.split(':');const [rH,tM]=st.end.split(':');
        row.getCell(colO).value=parseInt(oH);row.getCell(colQ).value=qM;
        row.getCell(colR_).value=parseInt(rH);row.getCell(colT).value=tM;
      }else{row.getCell(colO).value=null;row.getCell(colQ).value=null;row.getCell(colR_).value=null;row.getCell(colT).value=null;}
      row.getCell(colW).value=(st.status!=='All day in store'&&st.remark)?st.remark:null;
      modifiedCount++;
    });
    const lastR=ws.rowCount;
    const clearCols=[colL,colM,colN,colO,colQ,colR_,colT,colW];
    for(let r=4;r<=lastR;r++){
      const row=ws.getRow(r);const nv=row.getCell(colName).value;const nm=nv?String(nv).trim():'';
      let keep=false;
      if(nm===p){const date=excelDate(row.getCell(colDate).value);if(date&&selectedDateSet.has(ymdDash(date)))keep=true;}
      row.hidden=!keep;
      if(!keep)clearCols.forEach(c=>{row.getCell(c).value=null;});
    }
    const newLast=ws.rowCount;
    ws.views=[{state:'normal',topLeftCell:`A${Math.max(1,(firstRow||4)-1)}`,activeCell:`A${firstRow||4}`,showGridLines:true}];
    ws.autoFilter={from:{row:2,column:1},to:{row:newLast,column:23}};
    ws.dataValidations=new (ws.dataValidations.constructor)();
    const N_LIST='"All day in store,AM - Meeting/Training，PM - in store,AM - in store，PM - Meeting/Training,All day Meeting/Training,OFF,AL,SL,M,MT,PT,B,BT,CL,NP,W/On Leave,On Leave/W,RPH,PH1,PH2,PH3,PH4,PH5,PH6,PH7,PH8,PH9,PH10,PH11,PH12,PH13,PH14,PH15,PH16"';
    ws.dataValidations.add(`N4:N${newLast}`,{type:'list',allowBlank:true,formulae:[N_LIST]});
    ws.dataValidations.add(`O4:O${newLast}`,{type:'list',allowBlank:true,formulae:['"1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24"']});
    ws.dataValidations.add(`R4:R${newLast}`,{type:'list',allowBlank:true,formulae:['"1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24"']});
    ws.dataValidations.add(`Q4:Q${newLast}`,{type:'list',allowBlank:true,formulae:['"00,30"']});
    ws.dataValidations.add(`T4:T${newLast}`,{type:'list',allowBlank:true,formulae:['"00,30"']});
    ws.dataValidations.add(`V4:V${newLast}`,{type:'list',allowBlank:true,formulae:['"AM,PM"']});

    const first=NAME_ALIAS[p]||p.split(/\s+/)[0];
    const baseName=FILE_NAME.replace(/\.xlsx?$/i,'').trim();
    const outName=`${baseName} by ${first}.xlsx`;
    const outPath=path.join(OUT_DIR,outName);
    await wb.xlsx.writeFile(outPath);
    results.push({p,first,modifiedCount,outName,outPath});
  }

  console.log('\n=== 產出完成，開始驗證每個檔案 ===');
  let allOk=true;
  for(const r of results){
    // 驗證：能重新打開、可見列數=選取日期數、N欄都有值、時段欄位一致
    const wb2=new ExcelJS.Workbook();
    await wb2.xlsx.readFile(r.outPath);
    const ws2=wb2.getWorksheet(pickSheet);
    let visibleRows=0,nFilled=0,timeOK=true,storeFilled=0,badTime=0;
    ws2.eachRow({includeEmpty:false},(row,rNum)=>{
      if(rNum<4)return;
      if(row.hidden)return;
      const nm=row.getCell(12+ -11).value; // colName? careful
    });
    // 用固定欄位重新掃
    for(let rr=4;rr<=ws2.rowCount;rr++){
      const row=ws2.getRow(rr);
      if(row.hidden)continue;
      const nm=row.getCell(pickCol.name+1).value;
      if(!nm||String(nm).trim()!==r.p)continue;
      visibleRows++;
      const nVal=row.getCell(14).value;
      if(nVal!=null&&String(nVal).trim()!=='')nFilled++;
      if(row.getCell(12).value!=null)storeFilled++;
      const needs=NEEDS_TIME.has(String(nVal).trim());
      const oVal=row.getCell(15).value, rVal=row.getCell(18).value;
      if(needs){if(oVal==null||rVal==null)badTime++;}
      else{if(oVal!=null||rVal!=null)badTime++;}
    }
    const sizeKB=Math.round(fs.statSync(r.outPath).size/1024);
    const ok=(visibleRows===selectedDateSet.size)&&(nFilled===visibleRows)&&(badTime===0);
    if(!ok)allOk=false;
    console.log(`${ok?'✅':'❌'} ${r.first.padEnd(8)} 可見列:${visibleRows}/${selectedDateSet.size} N欄填:${nFilled} 店號填:${storeFilled} 時段錯:${badTime} ${sizeKB}KB`);
  }
  console.log('\n'+(allOk?'🎉 全部檔案驗證通過！':'⚠️ 有檔案驗證失敗，請看上面 ❌'));
  console.log('輸出資料夾:',OUT_DIR);
})().catch(e=>{console.error('壓測失敗:',e);process.exit(1);});
