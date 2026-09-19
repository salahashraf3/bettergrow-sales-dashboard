const MAX_FILE_SIZE=10*1024*1024;
const ids=['fileInput','browseButton','fileState','summaryBody','chart','totalBudget','totalActual','achievement','aboveCount','belowCount','budgetRows','actualVariance','errorNotice','dropZone','periodLabel','dashboard','emptyState','resetButton','headerStatus','rowCount'];
const els=Object.fromEntries(ids.map(id=>[id,document.getElementById(id)]));
const money=new Intl.NumberFormat('en-AE',{style:'currency',currency:'AED',maximumFractionDigits:2});
const compact=new Intl.NumberFormat('en-AE',{style:'currency',currency:'AED',notation:'compact',maximumFractionDigits:1});
const cleanHeader=value=>String(value??'').trim().toLowerCase().replace(/[^a-z0-9]/g,'');
const cleanText=value=>String(value??'').trim().replace(/\s+/g,' ');
const normalizeKey=value=>cleanText(value).toLocaleLowerCase('en');
const aliases={salesperson:['salesperson','salespersonname','salesrep','salesrepresentative'],month:['month','salesmonth','period'],budget:['budgetamount','budget','target','targetamount'],actual:['actualsales','actual','actualsalesamount','salesamount']};
const labels={salesperson:'Salesperson',month:'Month',budget:'Budget Amount',actual:'Actual Sales'};
let uploadSequence=0;
let currentData=null;

function parseAmount(value){
 if(typeof value==='number')return Number.isFinite(value)?value:NaN;
 if(typeof value==='boolean'||value==null)return NaN;
 const raw=String(value).trim();
 if(!raw)return NaN;
 // Text amounts may be plain numbers or explicitly AED. Symbols and other currency codes are rejected.
 const match=raw.match(/^(?:(AED)\s+)?([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?)$/i);
 if(!match)return NaN;
 return Number(match[2].replace(/,/g,''));
}
function locateHeader(matrix){
 for(let index=0;index<Math.min(matrix.length,20);index++){
  const mapped={};
  matrix[index].forEach((cell,column)=>{
   const key=cleanHeader(cell);
   for(const [field,list] of Object.entries(aliases))if(list.includes(key))(mapped[field]??=[]).push(column);
  });
  const hasEveryField=Object.keys(labels).every(field=>mapped[field]?.length>=1);
  if(hasEveryField){
   const duplicate=Object.entries(mapped).find(([,columns])=>columns.length>1);
   if(duplicate)throw Error(`Duplicate required column in Excel row ${index+1}: ${labels[duplicate[0]]}. Keep only one matching column.`);
   return {index,found:Object.fromEntries(Object.entries(mapped).map(([field,columns])=>[field,columns[0]]))};
  }
 }
 return null;
}
function rowsFromMatrix(matrix){
 if(!Array.isArray(matrix)||!matrix.length||matrix.every(row=>!row||row.every(value=>value==null||String(value).trim()==='')))throw Error('No data found. The first worksheet is empty.');
 const header=locateHeader(matrix);
 if(!header)throw Error('Required columns not found. Put Salesperson, Month, Budget Amount and Actual Sales together in one row within the first 20 Excel rows.');
 const rows=[];
 const pairs=new Map();
 for(let index=header.index+1;index<matrix.length;index++){
  const source=matrix[index]||[];
  if(source.every(value=>value==null||String(value).trim()===''))continue;
  const row=index+1;
  const salesperson=cleanText(source[header.found.salesperson]);
  const month=cleanText(source[header.found.month]);
  const budgetRaw=source[header.found.budget];
  const actualRaw=source[header.found.actual];
  const missing=[];
  if(!salesperson)missing.push('Salesperson');if(!month)missing.push('Month');
  if(budgetRaw==null||String(budgetRaw).trim()==='')missing.push('Budget Amount');
  if(actualRaw==null||String(actualRaw).trim()==='')missing.push('Actual Sales');
  if(missing.length)throw Error(`Incomplete record in Excel row ${row}. Missing: ${missing.join(', ')}.`);
  const budget=parseAmount(budgetRaw),actual=parseAmount(actualRaw);
  if(!Number.isFinite(budget))throw Error(`Invalid Budget Amount in Excel row ${row}. Use an Excel number, a plain number, or strict AED text such as AED 1,000.50.`);
  if(!Number.isFinite(actual))throw Error(`Invalid Actual Sales in Excel row ${row}. Use an Excel number, a plain number, or strict AED text such as AED 1,000.50.`);
  if(budget<0)throw Error(`Negative Budget Amount in Excel row ${row} is not supported.`);
  if(actual<0)throw Error(`Negative Actual Sales in Excel row ${row} is not supported. Returns and credit notes are outside this assessment's scope.`);
  const key=`${normalizeKey(salesperson)}\u0000${normalizeKey(month)}`;
  if(pairs.has(key))throw Error(`Duplicate salesperson and month in Excel rows ${pairs.get(key)} and ${row}: ${salesperson} · ${month}. Keep one record per salesperson per month.`);
  pairs.set(key,row);rows.push({salesperson,month,budget,actual,row});
 }
 if(!rows.length)throw Error('Headers found, but no sales records. Add at least one complete data row below the headers.');
 return rows;
}
function parseWorkbook(buffer){
 if(typeof XLSX==='undefined')throw Error('The Excel parser could not load. Check your connection and try again.');
 let book;
 try{book=XLSX.read(buffer,{type:'array',cellDates:true});}catch{throw Error('Unable to read this workbook. It may be password-protected or damaged. Upload an unprotected Excel file.');}
 if(!book.SheetNames.length)throw Error('No data found. The workbook has no worksheets.');
 const sheet=book.Sheets[book.SheetNames[0]];
 if(!sheet)throw Error('No data found. The first worksheet is empty.');
 return rowsFromMatrix(XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true}));
}
function summarize(rows){
 const grouped=new Map();
 for(const row of rows){
  const key=normalizeKey(row.salesperson);
  const item=grouped.get(key)||{name:cleanText(row.salesperson),budget:0,actual:0,months:new Set(),rows:0};
  item.budget+=row.budget;item.actual+=row.actual;item.months.add(cleanText(row.month));item.rows++;grouped.set(key,item);
 }
 return [...grouped.values()].map(item=>({...item,pct:item.budget===0?null:(item.actual/item.budget*100),status:item.budget===0?'no-budget':item.actual>=item.budget?'on-target':'below'})).sort((a,b)=>(b.pct??-Infinity)-(a.pct??-Infinity)||b.actual-a.actual||a.name.localeCompare(b.name));
}
const escapeHtml=value=>{const node=document.createElement('div');node.textContent=value;return node.innerHTML};
const shortName=value=>{const parts=value.split(/\s+/);return parts.length>1?`${parts[0]} ${parts.at(-1)[0]}.`:value};
const percentage=value=>value===null?'N/A':`${value.toFixed(1)}%`;
function getPeriod(rows){const values=[...new Set(rows.map(row=>row.month))];return values.length<=3?values.join(' · '):`${values[0]} - ${values.at(-1)}`;}
function activateReveals(root=document){const nodes=[...root.querySelectorAll('.reveal:not(.is-visible)')];if(matchMedia('(prefers-reduced-motion: reduce)').matches||!('IntersectionObserver'in window)){nodes.forEach(node=>node.classList.add('is-visible'));return;}const observer=new IntersectionObserver(entries=>entries.forEach(entry=>{if(entry.isIntersecting){entry.target.classList.add('is-visible');observer.unobserve(entry.target);}}),{threshold:.12});nodes.forEach((node,index)=>{node.style.setProperty('--delay',`${Math.min(index*55,220)}ms`);observer.observe(node);});}
function render(rows,fileName){
 const data=summarize(rows),budget=data.reduce((sum,item)=>sum+item.budget,0),actual=data.reduce((sum,item)=>sum+item.actual,0),pct=budget===0?null:(actual/budget*100),onTarget=data.filter(item=>item.status==='on-target').length,below=data.filter(item=>item.status==='below').length,noBudget=data.filter(item=>item.status==='no-budget').length,variance=actual-budget;
 els.totalBudget.textContent=compact.format(budget);els.totalActual.textContent=compact.format(actual);els.achievement.textContent=percentage(pct);els.aboveCount.textContent=`${onTarget} / ${data.length}`;els.belowCount.textContent=`${below} below${noBudget?` · ${noBudget} no budget`:''}`;els.budgetRows.textContent=`Across ${data.length} salesperson${data.length===1?'':'s'}`;els.actualVariance.textContent=`${variance>=0?'+':'−'}${compact.format(Math.abs(variance))} vs budget`;els.actualVariance.className=variance>=0?'positive':'negative';
 const statusText={"on-target":'On target',below:'Below target','no-budget':'No budget'};
 els.summaryBody.innerHTML=data.map(item=>`<tr><td><span class="person">${escapeHtml(item.name)}</span><span class="subline">${item.months.size} month${item.months.size===1?'':'s'} · ${item.rows} row${item.rows===1?'':'s'}</span></td><td>${money.format(item.budget)}</td><td>${money.format(item.actual)}</td><td><span class="progress"><span class="progress-track"><span class="progress-fill" style="width:${item.pct===null?0:Math.min(item.pct,100)}%"></span></span><strong>${percentage(item.pct)}</strong></span></td><td><span class="status ${item.status}">${statusText[item.status]}</span></td></tr>`).join('');
 const max=Math.max(...data.flatMap(item=>[item.budget,item.actual]),1);els.chart.innerHTML=data.map(item=>`<div class="chart-group"><div class="bars"><div class="bar budget" tabindex="0" aria-label="${escapeHtml(item.name)} budget ${money.format(item.budget)}" data-value="${money.format(item.budget)}" style="height:${item.budget/max*100}%"></div><div class="bar actual" tabindex="0" aria-label="${escapeHtml(item.name)} actual ${money.format(item.actual)}" data-value="${money.format(item.actual)}" style="height:${item.actual/max*100}%"></div></div><div class="chart-label">${escapeHtml(shortName(item.name))}</div></div>`).join('');
 els.periodLabel.textContent=getPeriod(rows);els.rowCount.textContent=`${rows.length} source row${rows.length===1?'':'s'}`;els.errorNotice.hidden=true;els.emptyState.hidden=true;els.dashboard.hidden=false;els.headerStatus.textContent='Dashboard ready';els.fileState.textContent=`${fileName} · ${rows.length} row${rows.length===1?'':'s'}`;document.body.classList.add('has-data');currentData={rows,fileName};activateReveals(els.dashboard);requestAnimationFrame(()=>els.dashboard.scrollIntoView({behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'start'}));
}
function showError(message){els.errorNotice.textContent=message;els.errorNotice.hidden=false;els.headerStatus.textContent=currentData?'Dashboard kept · upload rejected':'Upload needs attention';els.fileState.textContent=currentData?`${currentData.fileName} remains loaded`:'No valid workbook loaded';els.errorNotice.scrollIntoView({behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'center'});}
function reset(){uploadSequence++;currentData=null;els.dashboard.hidden=true;els.emptyState.hidden=false;els.errorNotice.hidden=true;els.fileInput.value='';els.fileState.textContent='No workbook loaded';els.headerStatus.textContent='Awaiting data';document.body.classList.remove('has-data');els.browseButton.focus();}
async function handleFile(file){
 if(!file)return;const sequence=++uploadSequence;els.errorNotice.hidden=true;
 if(!/\.(xlsx|xls)$/i.test(file.name))return showError('Choose an Excel file ending in .xlsx or .xls.');
 if(file.size===0)return showError('This file is empty. Choose a workbook with sales data.');
 if(file.size>MAX_FILE_SIZE)return showError('This file is over 10 MB. Choose a smaller workbook.');
 els.fileState.textContent=`Reading ${file.name}…`;els.headerStatus.textContent='Processing data';
 try{const buffer=await file.arrayBuffer();if(sequence!==uploadSequence)return;const rows=parseWorkbook(buffer);if(sequence!==uploadSequence)return;render(rows,file.name);}catch(error){if(sequence===uploadSequence)showError(error.message);}
}
function acceptSelectedFile(file){els.fileInput.value='';handleFile(file);}
els.browseButton.addEventListener('click',()=>els.fileInput.click());els.fileInput.addEventListener('change',event=>acceptSelectedFile(event.target.files[0]));els.resetButton.addEventListener('click',reset);
for(const event of ['dragenter','dragover'])els.dropZone.addEventListener(event,e=>{e.preventDefault();els.dropZone.classList.add('dragging')});for(const event of ['dragleave','drop'])els.dropZone.addEventListener(event,e=>{e.preventDefault();els.dropZone.classList.remove('dragging')});els.dropZone.addEventListener('drop',event=>handleFile(event.dataTransfer.files[0]));
window.__bettergrowTest={parseAmount,locateHeader,rowsFromMatrix,summarize,percentage,normalizeKey};activateReveals();
