const MAX_FILE_SIZE=10*1024*1024;
const ids=['fileInput','browseButton','fileState','summaryBody','chart','totalBudget','totalActual','achievement','aboveCount','belowCount','budgetRows','actualVariance','errorNotice','dropZone','periodLabel','dashboard','emptyState','resetButton','headerStatus','rowCount'];
const els=Object.fromEntries(ids.map(id=>[id,document.getElementById(id)]));
const money=new Intl.NumberFormat('en-AE',{style:'currency',currency:'AED',maximumFractionDigits:0});
const compact=new Intl.NumberFormat('en-AE',{style:'currency',currency:'AED',notation:'compact',maximumFractionDigits:1});
const cleanKey=value=>String(value??'').trim().toLowerCase().replace(/[^a-z0-9]/g,'');
const aliases={salesperson:['salesperson','salespersonname','salesrep','salesrepresentative'],month:['month','salesmonth','period'],budget:['budgetamount','budget','target','targetamount'],actual:['actualsales','actual','sales','salesamount']};
const labels={salesperson:'Salesperson',month:'Month',budget:'Budget Amount',actual:'Actual Sales'};

function parseAmount(value){
 if(typeof value==='number')return Number.isFinite(value)?value:NaN;
 const raw=String(value??'').trim();
 if(!raw)return NaN;
 const normalized=raw.replace(/\s/g,'').replace(/,/g,'').replace(/^(AED|USD|EUR|GBP|INR)/i,'').replace(/[₹$€£]/g,'');
 if(!/^-?(?:\d+\.?\d*|\.\d+)$/.test(normalized))return NaN;
 return Number(normalized);
}
function locateHeader(matrix){
 for(let index=0;index<Math.min(matrix.length,20);index++){
  const keys=matrix[index].map(cleanKey);
  const found=Object.fromEntries(Object.entries(aliases).map(([name,list])=>[name,keys.findIndex(key=>list.includes(key))]));
  if(Object.values(found).every(position=>position>=0))return {index,found};
 }
 return null;
}
function rowsFromMatrix(matrix){
 if(!matrix.length||matrix.every(row=>row.every(value=>value==null||String(value).trim()==='')))throw Error('The first worksheet is empty. Add a header row and at least one data row.');
 const header=locateHeader(matrix);
 if(!header){
  const flattened=matrix.slice(0,20).flat().map(cleanKey);
  const missing=Object.entries(aliases).filter(([,list])=>!list.some(key=>flattened.includes(key))).map(([name])=>labels[name]);
  throw Error(`Missing required column${missing.length===1?'':'s'}: ${missing.join(', ')}.`);
 }
 const rows=[];
 for(let index=header.index+1;index<matrix.length;index++){
  const source=matrix[index];
  if(!source||source.every(value=>value==null||String(value).trim()===''))continue;
  const salesperson=String(source[header.found.salesperson]??'').trim();
  const month=String(source[header.found.month]??'').trim();
  const budget=parseAmount(source[header.found.budget]);
  const actual=parseAmount(source[header.found.actual]);
  const row=index+1;
  if(!salesperson||!month)throw Error(`Invalid data in row ${row}. Salesperson and Month are required.`);
  if(!Number.isFinite(budget)||!Number.isFinite(actual))throw Error(`Invalid amount in row ${row}. Budget Amount and Actual Sales must be numbers.`);
  if(budget<0||actual<0)throw Error(`Invalid amount in row ${row}. Negative values are not supported.`);
  rows.push({salesperson,month,budget,actual,row});
 }
 if(!rows.length)throw Error('The worksheet has headers but no data rows.');
 return rows;
}
function parseWorkbook(buffer){
 if(typeof XLSX==='undefined')throw Error('The Excel parser could not load. Check your connection and try again.');
 let book;
 try{book=XLSX.read(buffer,{type:'array',cellDates:true});}catch{throw Error('This workbook could not be read. It may be damaged or password-protected.');}
 if(!book.SheetNames.length)throw Error('The workbook has no worksheets.');
 const sheet=book.Sheets[book.SheetNames[0]];
 if(!sheet)throw Error('The first worksheet could not be read.');
 return rowsFromMatrix(XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true}));
}
function summarize(rows){
 const grouped=new Map();
 for(const row of rows){
  const key=row.salesperson.trim().toLocaleLowerCase();
  const item=grouped.get(key)||{name:row.salesperson.trim(),budget:0,actual:0,months:new Set(),rows:0};
  item.budget+=row.budget;item.actual+=row.actual;item.months.add(row.month);item.rows++;grouped.set(key,item);
 }
 return [...grouped.values()].map(item=>({...item,pct:item.budget===0?(item.actual>0?null:0):(item.actual/item.budget*100),onTarget:item.actual>=item.budget})).sort((a,b)=>(b.pct??Infinity)-(a.pct??Infinity)||b.actual-a.actual||a.name.localeCompare(b.name));
}
const escapeHtml=value=>{const node=document.createElement('div');node.textContent=value;return node.innerHTML};
const shortName=value=>{const parts=value.split(/\s+/);return parts.length>1?`${parts[0]} ${parts.at(-1)[0]}.`:value};
const percentage=value=>value===null?'N/A':`${value.toFixed(1)}%`;
function getPeriod(rows){const values=[...new Set(rows.map(row=>row.month))];return values.length<=3?values.join(' · '):`${values[0]} - ${values.at(-1)}`;}
function activateReveals(root=document){
 const nodes=[...root.querySelectorAll('.reveal:not(.is-visible)')];
 if(matchMedia('(prefers-reduced-motion: reduce)').matches||!('IntersectionObserver'in window)){nodes.forEach(node=>node.classList.add('is-visible'));return;}
 const observer=new IntersectionObserver(entries=>entries.forEach(entry=>{if(entry.isIntersecting){entry.target.classList.add('is-visible');observer.unobserve(entry.target);}}),{threshold:.12});
 nodes.forEach((node,index)=>{node.style.setProperty('--delay',`${Math.min(index*55,220)}ms`);observer.observe(node);});
}
function render(rows){
 const data=summarize(rows);const budget=data.reduce((sum,item)=>sum+item.budget,0);const actual=data.reduce((sum,item)=>sum+item.actual,0);const pct=budget===0?(actual>0?null:0):(actual/budget*100);const above=data.filter(item=>item.onTarget).length;const variance=actual-budget;
 els.totalBudget.textContent=compact.format(budget);els.totalActual.textContent=compact.format(actual);els.achievement.textContent=percentage(pct);els.aboveCount.textContent=`${above} / ${data.length}`;els.belowCount.textContent=`${data.length-above} below target`;els.budgetRows.textContent=`Across ${data.length} salesperson${data.length===1?'':'s'}`;els.actualVariance.textContent=`${variance>=0?'+':'−'}${compact.format(Math.abs(variance))} vs budget`;els.actualVariance.className=variance>=0?'positive':'negative';
 els.summaryBody.innerHTML=data.map(item=>`<tr><td><span class="person">${escapeHtml(item.name)}</span><span class="subline">${item.months.size} month${item.months.size===1?'':'s'} · ${item.rows} row${item.rows===1?'':'s'}</span></td><td>${money.format(item.budget)}</td><td>${money.format(item.actual)}</td><td><span class="progress"><span class="progress-track"><span class="progress-fill" style="width:${item.pct===null?100:Math.min(item.pct,100)}%"></span></span><strong>${percentage(item.pct)}</strong></span></td><td><span class="status ${item.onTarget?'above':'below'}">${item.onTarget?'On target':'Below target'}</span></td></tr>`).join('');
 const max=Math.max(...data.flatMap(item=>[item.budget,item.actual]),1);els.chart.innerHTML=data.map(item=>`<div class="chart-group"><div class="bars"><div class="bar budget" tabindex="0" aria-label="${escapeHtml(item.name)} budget ${money.format(item.budget)}" data-value="${money.format(item.budget)}" style="height:${item.budget/max*100}%"></div><div class="bar actual" tabindex="0" aria-label="${escapeHtml(item.name)} actual ${money.format(item.actual)}" data-value="${money.format(item.actual)}" style="height:${item.actual/max*100}%"></div></div><div class="chart-label">${escapeHtml(shortName(item.name))}</div></div>`).join('');
 els.periodLabel.textContent=getPeriod(rows);els.rowCount.textContent=`${rows.length} source row${rows.length===1?'':'s'}`;els.errorNotice.hidden=true;els.emptyState.hidden=true;els.dashboard.hidden=false;els.headerStatus.textContent='Dashboard ready';document.body.classList.add('has-data');
 activateReveals(els.dashboard);requestAnimationFrame(()=>els.dashboard.scrollIntoView({behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'start'}));
}
function showError(message){els.errorNotice.textContent=message;els.errorNotice.hidden=false;els.headerStatus.textContent='Upload needs attention';els.fileState.textContent='No valid workbook loaded';els.errorNotice.scrollIntoView({behavior:'smooth',block:'center'});}
function reset(){els.dashboard.hidden=true;els.emptyState.hidden=false;els.errorNotice.hidden=true;els.fileInput.value='';els.fileState.textContent='No workbook loaded';els.headerStatus.textContent='Awaiting data';document.body.classList.remove('has-data');els.browseButton.focus();}
async function handleFile(file){
 if(!file)return;
 els.errorNotice.hidden=true;
 if(!/\.(xlsx|xls)$/i.test(file.name))return showError('Choose an Excel file ending in .xlsx or .xls.');
 if(file.size===0)return showError('This file is empty. Choose a workbook with sales data.');
 if(file.size>MAX_FILE_SIZE)return showError('This file is over 10 MB. Choose a smaller workbook.');
 els.fileState.textContent='Reading workbook…';els.headerStatus.textContent='Processing data';
 try{const rows=parseWorkbook(await file.arrayBuffer());render(rows);els.fileState.textContent=`${file.name} · ${rows.length} row${rows.length===1?'':'s'}`;}catch(error){showError(error.message);}finally{els.fileInput.value='';}
}
els.browseButton.addEventListener('click',()=>els.fileInput.click());els.fileInput.addEventListener('change',event=>handleFile(event.target.files[0]));els.resetButton.addEventListener('click',reset);
for(const event of ['dragenter','dragover'])els.dropZone.addEventListener(event,e=>{e.preventDefault();els.dropZone.classList.add('dragging')});
for(const event of ['dragleave','drop'])els.dropZone.addEventListener(event,e=>{e.preventDefault();els.dropZone.classList.remove('dragging')});
els.dropZone.addEventListener('drop',event=>handleFile(event.dataTransfer.files[0]));
window.__bettergrowTest={parseAmount,locateHeader,rowsFromMatrix,summarize,percentage};
activateReveals();
