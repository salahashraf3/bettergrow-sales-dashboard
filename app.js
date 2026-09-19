const sampleRows=[
['Aisha Rahman','Jan 2026',85000,91200],['Aisha Rahman','Feb 2026',88000,86400],['Aisha Rahman','Mar 2026',90000,96750],['Aisha Rahman','Apr 2026',92000,101300],
['Omar Khalid','Jan 2026',78000,73100],['Omar Khalid','Feb 2026',80000,82400],['Omar Khalid','Mar 2026',82000,79300],['Omar Khalid','Apr 2026',84000,88600],
['Meera Nair','Jan 2026',72000,76800],['Meera Nair','Feb 2026',74000,80100],['Meera Nair','Mar 2026',76000,84500],['Meera Nair','Apr 2026',78000,81900],
['Daniel Joseph','Jan 2026',68000,62400],['Daniel Joseph','Feb 2026',70000,66700],['Daniel Joseph','Mar 2026',72000,71400],['Daniel Joseph','Apr 2026',74000,76800],
['Fatima Noor','Jan 2026',65000,70100],['Fatima Noor','Feb 2026',67000,68900],['Fatima Noor','Mar 2026',69000,74800],['Fatima Noor','Apr 2026',71000,79000]
].map(([salesperson,month,budget,actual])=>({salesperson,month,budget,actual}));
const ids=['fileInput','browseButton','fileState','resetButton','summaryBody','chart','totalBudget','totalActual','achievement','aboveCount','belowCount','budgetRows','actualVariance','errorNotice','dropZone','periodLabel'];
const els=Object.fromEntries(ids.map(id=>[id,document.getElementById(id)]));
const money=new Intl.NumberFormat('en-AE',{style:'currency',currency:'AED',maximumFractionDigits:0});
const compact=new Intl.NumberFormat('en-AE',{style:'currency',currency:'AED',notation:'compact',maximumFractionDigits:1});
const cleanKey=k=>String(k).trim().toLowerCase().replace(/[^a-z0-9]/g,'');
const number=v=>typeof v==='number'?v:Number(String(v??'').replace(/[^0-9.-]/g,''));
function parseWorkbook(buffer){
 const book=XLSX.read(buffer,{type:'array',cellDates:true});
 if(!book.SheetNames.length)throw Error('The workbook has no worksheets.');
 const raw=XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{defval:null});
 if(!raw.length)throw Error('The first worksheet has no data rows.');
 const keys={};Object.keys(raw[0]).forEach(k=>keys[cleanKey(k)]=k);
 const aliases={salesperson:['salesperson','salespersonname'],month:['month','salesmonth'],budget:['budgetamount','budget'],actual:['actualsales','actual','sales']};
 const col={};for(const [name,list] of Object.entries(aliases)){col[name]=list.map(x=>keys[x]).find(Boolean);if(!col[name])throw Error(`Missing required column: ${{salesperson:'Salesperson',month:'Month',budget:'Budget Amount',actual:'Actual Sales'}[name]}.`)}
 const rows=raw.map((r,i)=>({salesperson:String(r[col.salesperson]??'').trim(),month:String(r[col.month]??'').trim(),budget:number(r[col.budget]),actual:number(r[col.actual]),row:i+2}));
 const bad=rows.find(r=>!r.salesperson||!r.month||!Number.isFinite(r.budget)||!Number.isFinite(r.actual)||r.budget<0||r.actual<0);
 if(bad)throw Error(`Invalid data in row ${bad.row}. Names and months are required; amounts must be non-negative numbers.`);
 return rows;
}
function summarize(rows){const grouped=new Map();for(const r of rows){const x=grouped.get(r.salesperson)||{name:r.salesperson,budget:0,actual:0,months:new Set};x.budget+=r.budget;x.actual+=r.actual;x.months.add(r.month);grouped.set(r.salesperson,x)}return [...grouped.values()].map(x=>({...x,pct:x.budget===0?(x.actual>0?Infinity:0):x.actual/x.budget*100})).sort((a,b)=>b.pct-a.pct)}
function html(v){const d=document.createElement('div');d.textContent=v;return d.innerHTML}
function shortName(v){const p=v.split(/\s+/);return p.length>1?`${p[0]} ${p.at(-1)[0]}.`:v}
function render(rows){
 const data=summarize(rows),budget=data.reduce((s,x)=>s+x.budget,0),actual=data.reduce((s,x)=>s+x.actual,0),pct=budget?actual/budget*100:0,above=data.filter(x=>x.actual>=x.budget).length,variance=actual-budget;
 els.totalBudget.textContent=compact.format(budget);els.totalActual.textContent=compact.format(actual);els.achievement.textContent=`${pct.toFixed(1)}%`;els.aboveCount.textContent=`${above} / ${data.length}`;els.belowCount.textContent=`${data.length-above} below target`;els.budgetRows.textContent=`Across ${data.length} salespeople`;els.actualVariance.textContent=`${variance>=0?'+':'−'}${compact.format(Math.abs(variance))} vs budget`;els.actualVariance.className=variance>=0?'positive':'negative';
 els.summaryBody.innerHTML=data.map(x=>`<tr><td><span class="person">${html(x.name)}</span><span class="subline">${x.months.size} month${x.months.size===1?'':'s'} reported</span></td><td>${money.format(x.budget)}</td><td>${money.format(x.actual)}</td><td><span class="progress"><span class="progress-track"><span class="progress-fill" style="width:${Math.min(x.pct,100)}%"></span></span><strong>${Number.isFinite(x.pct)?x.pct.toFixed(1):'∞'}%</strong></span></td><td><span class="status ${x.actual>=x.budget?'above':'below'}">${x.actual>=x.budget?'Above target':'Below target'}</span></td></tr>`).join('');
 const max=Math.max(...data.flatMap(x=>[x.budget,x.actual]),1);els.chart.innerHTML=data.map(x=>`<div class="chart-group"><div class="bars"><div class="bar budget" tabindex="0" aria-label="${html(x.name)} budget ${money.format(x.budget)}" data-value="${money.format(x.budget)}" style="height:${x.budget/max*100}%"></div><div class="bar actual" tabindex="0" aria-label="${html(x.name)} actual ${money.format(x.actual)}" data-value="${money.format(x.actual)}" style="height:${x.actual/max*100}%"></div></div><div class="chart-label">${html(shortName(x.name))}</div></div>`).join('');
 const periods=[...new Set(rows.map(x=>x.month))];els.periodLabel.textContent=periods.length<=3?periods.join(' · '):`${periods[0]} - ${periods.at(-1)}`;els.errorNotice.hidden=true;
}
function error(message){els.errorNotice.textContent=message;els.errorNotice.hidden=false;els.errorNotice.scrollIntoView({behavior:'smooth',block:'center'})}
async function handleFile(file){if(!file)return;if(!/\.xlsx?$/i.test(file.name))return error('Choose an Excel file ending in .xlsx or .xls.');if(file.size>10*1024*1024)return error('This file is over 10 MB. Please use a smaller workbook.');try{render(parseWorkbook(await file.arrayBuffer()));els.fileState.textContent=file.name;els.resetButton.hidden=false}catch(e){error(e.message);els.fileInput.value=''}}
els.browseButton.onclick=()=>els.fileInput.click();els.fileInput.onchange=e=>handleFile(e.target.files[0]);els.resetButton.onclick=()=>{render(sampleRows);els.fileState.textContent='Showing sample data';els.fileInput.value='';els.resetButton.hidden=true};
for(const event of ['dragenter','dragover'])els.dropZone.addEventListener(event,e=>{e.preventDefault();els.dropZone.classList.add('dragging')});for(const event of ['dragleave','drop'])els.dropZone.addEventListener(event,e=>{e.preventDefault();els.dropZone.classList.remove('dragging')});els.dropZone.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));render(sampleRows);
