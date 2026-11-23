/* Read Aloud Studio — Analyze data module: profile, insights, charts, ask-the-data, report export */
(function(){
const R=window.RAS;if(!R)return;
const $=id=>document.getElementById(id),esc=R.esc;
const CDN={chart:'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js',xlsx:'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',docx:'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',pdf:'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',pptx:'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js'};
const got={};const lib=u=>got[u]||(got[u]=new Promise((ok,no)=>{const s=document.createElement('script');s.src=u;s.onload=ok;s.onerror=()=>{delete got[u];no(new Error('A library could not load. Check your connection and try again.'))};document.head.appendChild(s)}));
const nf=new Intl.NumberFormat(undefined,{maximumFractionDigits:2}),fmt=v=>typeof v==='number'&&isFinite(v)?nf.format(v):String(v??'');
const pct=v=>(v>0?'+':'')+nf.format(v*100)+'%';

/* ---------- markup ---------- */
const css=document.createElement('style');css.textContent=`
.dat{margin:0 0 40px}.dat .tabs{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 14px}
.dat .tabs button[aria-selected=true]{background:var(--sky);color:#06213a;border-color:var(--sky);font-weight:600}
.dat .src{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:8px 0}
.dat .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin:8px 0 14px}
.dat .kpi{background:var(--deep);border:1px solid var(--line);border-radius:14px;padding:10px 14px}.dat .kpi b{display:block;font-size:26px;font-weight:300;color:var(--sky)}.dat .kpi span{font-size:13px;color:var(--soft)}
.dat table{width:100%;border-collapse:collapse;font-size:14px}.dat th,.dat td{border-bottom:1px solid var(--line);padding:7px 8px;text-align:left;vertical-align:top}.dat td.n{text-align:right;white-space:nowrap}
.dat .scroll{overflow-x:auto}.dat ul.ins{padding-left:20px;margin:6px 0}.dat ul.ins li{margin:6px 0}
.dat .charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,420px),1fr));gap:14px}
.dat .ch{background:var(--deep);border:1px solid var(--line);border-radius:14px;padding:12px}.dat .ch h3{font-size:16px;margin:0 0 8px;font-weight:600}
.dat .cv{position:relative;height:280px}.dat .row{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,150px),1fr));gap:10px;margin:10px 0}
.dat .acts2{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;align-items:center}.dat .acts2 button{min-height:36px;font-size:14px;padding:0 12px}
.dat input[type=text]{width:100%;font:inherit;font-size:16px;color:var(--ink);background:var(--deep);border:1px solid var(--line);border-radius:12px;min-height:44px;padding:0 12px}
.dat .go{background:var(--sky);color:#06213a;border-color:var(--sky);font-weight:600}
.dat .ans{background:var(--deep);border:1px solid var(--sky);border-radius:14px;padding:12px 14px;margin:10px 0}
.dat .note{font-size:13px;color:var(--soft)}
@media(max-width:600px){.dat table{font-size:13px}}`;
document.head.appendChild(css);
const box=document.createElement('section');box.className='panel dat';box.id='dat';
box.innerHTML=`<h2>Analyze data</h2>
<p class="how">Statistics, insights and charts for spreadsheets and CSV files. Numbers are calculated exactly in your browser; AI only explains them and turns your questions into calculations.</p>
<div class="src"><label class="btn" for="dt-file" style="display:inline-flex;align-items:center">Open a data file</label><input type="file" id="dt-file" accept=".xlsx,.xls,.csv" hidden>
<button id="dt-sample">Try sample sales data</button><label style="min-width:180px">Sheet<select id="dt-sheet"></select></label></div>
<div class="note" id="dt-src">Open a spreadsheet (here or in the reader) to begin.</div>
<div class="tabs" role="tablist"><button role="tab" aria-selected="true" data-t="ov">Overview</button><button role="tab" aria-selected="false" data-t="ch">Charts</button><button role="tab" aria-selected="false" data-t="ask">Ask the data</button><button role="tab" aria-selected="false" data-t="rep">Report</button></div>
<div data-p="ov"><div id="dt-ov"></div></div>
<div data-p="ch" hidden>
 <div class="row"><label>Chart<select id="cb-type"><option value="bar">Bar</option><option value="line">Line (trend)</option><option value="pie">Pie</option><option value="hist">Histogram</option><option value="scatter">Scatter</option></select></label>
 <label>Group by / X<select id="cb-x"></select></label><label>Value / Y<select id="cb-y"></select></label>
 <label>Calculate<select id="cb-agg"><option value="sum">Sum</option><option value="mean">Average</option><option value="count">Count</option><option value="max">Maximum</option><option value="min">Minimum</option></select></label>
 <label>Show top<select id="cb-top"><option value="10">10</option><option value="5">5</option><option value="20">20</option><option value="50">50</option></select></label></div>
 <button class="go" id="cb-go">Build chart</button><div class="charts" id="dt-charts" style="margin-top:14px"></div></div>
<div data-p="ask" hidden>
 <p class="note">Examples: "Which region had the highest revenue?", "Show monthly revenue trend", "Average units by product in the North".</p>
 <input type="text" id="aq-q" placeholder="Ask a question about this data">
 <div class="acts2"><button class="go" id="aq-go">Answer</button><span class="note" id="aq-tip"></span></div><div id="aq-out"></div></div>
<div data-p="rep" hidden>
 <label>Audience or goal (optional)<input type="text" id="rp-aud" placeholder="e.g. monthly board update, class project, client report"></label>
 <div class="acts2"><button class="go" id="rp-go">Write findings with AI</button><button id="rp-listen">Listen to findings</button><span class="note" id="rp-tip"></span></div>
 <div id="rp-out"></div>
 <div class="row"><label>Export as<select id="rp-fmt"><option value="docx">Word report (.docx)</option><option value="pdf">PDF report</option><option value="pptx">PowerPoint deck (.pptx)</option><option value="xlsx">Excel workbook (.xlsx)</option><option value="html">Web page (.html)</option></select></label></div>
 <div class="acts2"><button class="go" id="rp-export">Create report file</button><span class="note" id="rp-xtip"></span></div>
 <p class="note">Charts marked "In report" on the Charts tab are included.</p></div>`;
(document.getElementById('pro')||document.querySelector('.grid')).after(box);

/* ---------- data model ---------- */
let DS=null,sheetIdx=0,T=null,charts=[],cid=0,findings=null,selfLoad=false,ctrl=null;
const isNum=v=>typeof v==='number'?isFinite(v):typeof v==='string'&&/^\s*[-+]?[$€£]?\s*[\d,]*\.?\d+\s*%?\s*$/.test(v)&&v.trim()!=='';
const toNum=v=>typeof v==='number'?v:parseFloat(String(v).replace(/[$€£,\s%]/g,''));
const serial=v=>new Date(Math.round(v-25569)*864e5);
const toDate=(v,f)=>{if(typeof v==='number')return serial(v);const s=String(f??v).trim();if(/^\d{4}-\d{2}(-\d{2})?$/.test(s))return new Date(s.length===7?s+'-01T00:00:00Z':s+'T00:00:00Z');const t=Date.parse(s);if(isNaN(t))return null;const d=new Date(t);return new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()))};
function build(sheet){const rows=sheet.rows,raw=sheet.raw||sheet.rows;if(!rows||rows.length<2)return null;
 const nc=Math.max(...rows.map(r=>r.length));const head=Array.from({length:nc},(_,c)=>String(rows[0][c]??'').trim()||'Column '+(c+1));
 const seen={};head.forEach((h,i)=>{if(seen[h])head[i]=h+' ('+(++seen[h])+')';else seen[h]=1});
 const body=raw.slice(1).filter(r=>r.some(v=>v!==''&&v!=null)),fbody=rows.slice(1).filter(r=>r.some(v=>v!==''&&v!=null));
 const cols=head.map((name,c)=>{const vals=body.map(r=>r[c]),fv=fbody.map(r=>r[c]),ne=vals.filter(v=>v!==''&&v!=null);let type='text';
  const nums=ne.filter(isNum);const dateHint=/date|month|year|time|day|period|week/i.test(name);
  const dateStr=fv.filter(v=>typeof v==='string'&&/\d{1,4}[\/\-.]\d{1,2}([\/\-.]\d{1,4})?|[a-z]{3,9}[ \-]\d{2,4}/i.test(v)&&!isNaN(Date.parse(v))).length;
  if(ne.length&&(dateStr/ne.length>=.8||(dateHint&&nums.length/ne.length>=.8&&nums.every(v=>{const n=toNum(v);return n>20000&&n<80000}))))type='date';
  else if(ne.length&&nums.length/ne.length>=.8&&!/^(id|code|zip|phone)/i.test(name))type='number';
  else{const u=new Set(ne.map(String)).size;type=u<=Math.max(25,ne.length*.2)||u<=50?'category':'text'}
  const val=i=>{const v=vals[i];if(v===''||v==null)return null;if(type==='number')return isNum(v)?toNum(v):null;if(type==='date'){const d=toDate(v,fv[i]);return d&&!isNaN(d)?d:null}return String(v).trim()};
  return{name,type,v:vals.map((_,i)=>val(i))}});
 return{name:sheet.name,n:body.length,cols}}
function stats(a){const x=a.filter(v=>v!=null).sort((p,q)=>p-q),n=x.length;if(!n)return{n:0};const s=x.reduce((p,q)=>p+q,0),m=s/n,qt=p=>{const i=(n-1)*p,lo=Math.floor(i);return x[lo]+(x[Math.min(lo+1,n-1)]-x[lo])*(i-lo)};
 const sd=n>1?Math.sqrt(x.reduce((p,q)=>p+(q-m)**2,0)/(n-1)):0;return{n,sum:s,mean:m,median:qt(.5),sd,min:x[0],max:x[n-1],q1:qt(.25),q3:qt(.75)}}
function counts(a){const m=new Map();a.forEach(v=>{if(v!=null)m.set(v,(m.get(v)||0)+1)});return[...m].sort((p,q)=>q[1]-p[1])}
function corr(a,b){let n=0,sa=0,sb=0,saa=0,sbb=0,sab=0;for(let i=0;i<a.length;i++){const x=a[i],y=b[i];if(x==null||y==null)continue;n++;sa+=x;sb+=y;saa+=x*x;sbb+=y*y;sab+=x*y}if(n<5)return null;const d=Math.sqrt((n*saa-sa*sa)*(n*sbb-sb*sb));return d?(n*sab-sa*sb)/d:null}
const dkey=(d,span)=>span>400*864e5||span>60*864e5?d.toISOString().slice(0,7):d.toISOString().slice(0,10);
const col=n=>T.cols.find(c=>c.name===n);
const nums=()=>{const n=T.cols.filter(c=>c.type==='number'),key=c=>/revenue|sales|amount|profit|income|total|value|price|spend|cost/i.test(c.name)?1:0,sum=c=>Math.abs(stats(c.v).sum||0);return n.slice().sort((a,b)=>key(b)-key(a)||sum(b)-sum(a))};

/* ---------- compute a chart/query spec ---------- */
function compute(s){const X=col(s.x),Y=s.y&&s.y!=='(count)'?col(s.y):null,agg=Y?s.agg||'sum':'count';
 let idx=[...Array(T.n).keys()];
 (s.filters||[]).forEach(f=>{const C=col(f.col);if(!C)return;const fv=String(f.value??'').toLowerCase(),fn=toNum(f.value);idx=idx.filter(i=>{const v=C.v[i];if(v==null)return false;const sv=v instanceof Date?v.toISOString().slice(0,10):String(v).toLowerCase();
  switch(f.op){case'!=':return sv!==fv;case'>':return C.type==='date'?sv>fv:v>fn;case'<':return C.type==='date'?sv<fv:v<fn;case'>=':return v>=fn;case'<=':return v<=fn;case'contains':return sv.includes(fv);default:return sv===fv}})});
 if(s.type==='scatter'){if(!X||!Y)throw new Error('Pick two number columns for a scatter chart.');const pts=idx.filter(i=>X.v[i]!=null&&Y.v[i]!=null).slice(0,3000).map(i=>({x:X.v[i],y:Y.v[i]}));return{kind:'scatter',pts,r:corr(idx.map(i=>X.v[i]),idx.map(i=>Y.v[i])),xl:X.name,yl:Y.name}}
 if(s.type==='hist'){const C=Y||X;if(!C||C.type!=='number')throw new Error('Pick a number column for a histogram.');const v=idx.map(i=>C.v[i]).filter(x=>x!=null);const st=stats(v);const k=Math.min(20,Math.max(5,Math.ceil(Math.log2(v.length)+1))),w=(st.max-st.min)/k||1;const b=Array(k).fill(0);v.forEach(x=>b[Math.min(k-1,Math.floor((x-st.min)/w))]++);return{kind:'bar',labels:b.map((_,i)=>fmt(st.min+i*w)+'–'+fmt(st.min+(i+1)*w)),values:b,yl:'Count',xl:C.name,keepOrder:true}}
 if(!X)throw new Error('Pick a column to group by.');
 let span=0;if(X.type==='date'){const t=idx.map(i=>X.v[i]).filter(Boolean).map(d=>+d);span=Math.max(...t)-Math.min(...t)}
 const g=new Map();idx.forEach(i=>{let k=X.v[i];if(k==null)return;if(k instanceof Date)k=dkey(k,span);else if(X.type==='number')k=fmt(k);const y=Y?Y.v[i]:1;if(Y&&y==null)return;const a=g.get(k)||[];a.push(y);g.set(k,a)});
 let rows=[...g].map(([k,a])=>[k,agg==='count'?a.length:agg==='mean'?a.reduce((p,q)=>p+q,0)/a.length:agg==='max'?Math.max(...a):agg==='min'?Math.min(...a):a.reduce((p,q)=>p+q,0)]);
 const trend=s.type==='line'||X.type==='date';if(trend)rows.sort((a,b)=>a[0]<b[0]?-1:1);else rows.sort((a,b)=>s.sort==='asc'?a[1]-b[1]:b[1]-a[1]);
 const top=+s.top||(trend?500:10);if(rows.length>top&&!trend){const rest=rows.slice(top);rows=rows.slice(0,top);if(s.type==='pie'&&agg!=='mean')rows.push(['Other',rest.reduce((p,q)=>p+q[1],0)])}
 return{kind:s.type==='pie'?'pie':trend?'line':'bar',labels:rows.map(r=>String(r[0])),values:rows.map(r=>r[1]),xl:X.name,yl:Y?({sum:'Total ',mean:'Average ',max:'Max ',min:'Min ',count:'Count of '}[agg]||'')+Y.name:'Count',keepOrder:trend}}

/* ---------- charts ---------- */
const PAL=['#8fd3ff','#ff5c8a','#ffd166','#7ee0a1','#b39dff','#ff9f68','#5ec8c8','#f7a8d8'];
async function draw(canvas,res,light){await lib(CDN.chart);const C=window.Chart,ink=light?'#1a1a1a':'#b9cde2',grid=light?'#dddddd':'#1d4268',many=res.kind==='pie'||res.kind==='bar'&&res.values.length<=8&&!res.keepOrder;
 const data=res.kind==='scatter'?{datasets:[{label:res.yl,data:res.pts,backgroundColor:light?'#1e88c8':'#8fd3ff',pointRadius:3}]}:{labels:res.labels,datasets:[{label:res.yl,data:res.values,backgroundColor:res.kind==='pie'?PAL:many?PAL.map(c=>c):light?'#1e88c8':'#8fd3ff',borderColor:res.kind==='line'?(light?'#1e88c8':'#8fd3ff'):res.kind==='pie'?(light?'#fff':'#071b30'):undefined,borderWidth:res.kind==='line'?2.5:1,tension:.25,pointRadius:res.values.length>60?0:3,fill:false}]};
 return new C(canvas,{type:res.kind,data,options:{responsive:!light,maintainAspectRatio:false,animation:light?false:{duration:400},devicePixelRatio:light?1:undefined,plugins:{legend:{display:res.kind==='pie',labels:{color:ink}},tooltip:{callbacks:{label:c=>' '+fmt(c.parsed.y??c.parsed)}}},
  scales:res.kind==='pie'?{}:{x:{type:res.kind==='scatter'?'linear':'category',title:{display:true,text:res.xl,color:ink},ticks:{color:ink,maxRotation:50,autoSkip:true},grid:{color:grid}},y:{title:{display:true,text:res.yl,color:ink},ticks:{color:ink,callback:v=>fmt(v)},grid:{color:grid},beginAtZero:res.kind==='bar'}}}})}
async function png(res,title,type){const host=document.createElement('div');host.style.cssText='position:fixed;left:-9999px;top:0;width:1200px;height:675px;background:#fff';const cv=document.createElement('canvas');cv.width=1200;cv.height=675;host.appendChild(cv);document.body.appendChild(host);
 const ch=await draw(cv,res,true);const out=document.createElement('canvas');out.width=1200;out.height=740;const g=out.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,1200,740);g.fillStyle='#0a2540';g.font='600 30px Arial';g.fillText(title.slice(0,70),24,44);g.drawImage(cv,0,62);ch.destroy();host.remove();return out.toDataURL(type||'image/png',0.85)}
async function addChart(spec,title,inReport){let res;try{res=compute(spec)}catch(e){return e.message}
 const id='c'+(++cid);title=title||(spec.type==='hist'?`Distribution of ${res.xl}`:res.kind==='scatter'?`${res.yl} vs ${res.xl}${res.r!=null?` (r = ${nf.format(res.r)})`:''}`:`${res.yl} by ${res.xl}`);
 const card=document.createElement('div');card.className='ch';card.id=id;card.innerHTML=`<h3>${esc(title)}</h3><div class="cv"><canvas></canvas></div><div class="acts2"><label class="tog" style="min-height:36px;font-size:14px"><input type="checkbox" ${inReport?'checked':''}> In report</label><button data-a="png">Download image</button><button data-a="del">Remove</button></div>`;
 $('dt-charts').prepend(card);const item={id,spec,title,res,inReport,chart:null};charts.unshift(item);
 card.querySelector('input').onchange=e=>item.inReport=e.target.checked;
 card.querySelector('[data-a=png]').onclick=async()=>save(await(await fetch(await png(res,title))).blob(),slug(title)+'.png');
 card.querySelector('[data-a=del]').onclick=()=>{item.chart&&item.chart.destroy();card.remove();charts=charts.filter(c=>c!==item)};
 try{item.chart=await draw(card.querySelector('canvas'),res,false)}catch(e){card.querySelector('.cv').innerHTML=`<p class="note">${esc(e.message)}</p>`}return item}

/* ---------- overview + automatic insights ---------- */
function insights(){const out=[],num=nums(),cat=T.cols.filter(c=>c.type==='category'),date=T.cols.find(c=>c.type==='date');
 T.cols.forEach(c=>{const miss=c.v.filter(v=>v==null).length;if(miss/T.n>.1)out.push(`${c.name} is missing in ${nf.format(miss/T.n*100)}% of rows (${miss} of ${T.n}), so results using it may be incomplete.`)});
 if(date&&num.length){const y=num[0];const r=compute({type:'line',x:date.name,y:y.name,agg:'sum'});if(r.values.length>=3){const a=r.values[0],b=r.values[r.values.length-1],best=r.values.indexOf(Math.max(...r.values));out.push(`Total ${y.name} went from ${fmt(a)} in ${r.labels[0]} to ${fmt(b)} in ${r.labels[r.labels.length-1]}${a?` (${pct((b-a)/Math.abs(a))})`:''}; the peak was ${fmt(r.values[best])} in ${r.labels[best]}.`)}}
 if(cat.length&&num.length){const c=cat.find(c=>{const u=new Set(c.v.filter(v=>v!=null)).size;return u>=2&&u<=30})||cat[0];const y=num[0];const r=compute({type:'bar',x:c.name,y:y.name,agg:'sum',top:100});const tot=r.values.reduce((p,q)=>p+q,0);
  if(r.values.length>=2&&tot)out.push(`By ${c.name}, ${r.labels[0]} leads ${y.name} with ${fmt(r.values[0])} (${nf.format(r.values[0]/tot*100)}% of the total); ${r.labels[r.labels.length-1]} is lowest at ${fmt(r.values[r.values.length-1])}.`)}
 const pairs=[];for(let i=0;i<num.length&&i<12;i++)for(let j=i+1;j<num.length&&j<12;j++){const r=corr(num[i].v,num[j].v);if(r!=null&&Math.abs(r)>=.5)pairs.push([Math.abs(r),r,num[i].name,num[j].name])}
 pairs.sort((a,b)=>b[0]-a[0]).slice(0,3).forEach(p=>out.push(`${p[2]} and ${p[3]} ${p[1]>0?'rise and fall together':'move in opposite directions'} (correlation ${nf.format(p[1])}, ${p[0]>=.8?'very strong':'strong'}). This shows association, not cause.`));
 num.slice(0,6).forEach(c=>{const s=stats(c.v);if(s.n<8)return;const iqr=s.q3-s.q1,lo=s.q1-1.5*iqr,hi=s.q3+1.5*iqr,o=c.v.filter(v=>v!=null&&(v<lo||v>hi)).length;if(o&&iqr>0)out.push(`${c.name} has ${o} unusual value${o>1?'s':''} outside the typical range (${fmt(lo)} to ${fmt(hi)}); worth checking for errors or special cases.`)});
 return out}
function overview(){if(!T){$('dt-ov').innerHTML='<p class="note">No data yet.</p>';return}
 const miss=T.cols.reduce((p,c)=>p+c.v.filter(v=>v==null).length,0)/(T.n*T.cols.length||1);
 const rows=T.cols.map(c=>{const m=c.v.filter(v=>v==null).length;let d='';if(c.type==='number'){const s=stats(c.v);d=`avg ${fmt(s.mean)} · median ${fmt(s.median)} · ${fmt(s.min)} to ${fmt(s.max)} · total ${fmt(s.sum)}`}
  else if(c.type==='date'){const t=c.v.filter(Boolean).map(Number);d=t.length?`${new Date(Math.min(...t)).toISOString().slice(0,10)} to ${new Date(Math.max(...t)).toISOString().slice(0,10)}`:''}
  else{const k=counts(c.v);d=`${k.length} distinct · top: ${k.slice(0,3).map(([v,n])=>`${v} (${n})`).join(', ')}`}
  return`<tr><td>${esc(c.name)}</td><td>${c.type}</td><td class="n">${m}</td><td>${esc(d)}</td></tr>`}).join('');
 const ins=insights();T.ins=ins;
 $('dt-ov').innerHTML=`<div class="kpis"><div class="kpi"><b>${fmt(T.n)}</b><span>rows</span></div><div class="kpi"><b>${T.cols.length}</b><span>columns</span></div><div class="kpi"><b>${T.cols.filter(c=>c.type==='number').length}</b><span>number columns</span></div><div class="kpi"><b>${nf.format(miss*100)}%</b><span>cells missing</span></div></div>
 <h3 style="margin:6px 0">Key insights</h3>${ins.length?`<ul class="ins">${ins.map(i=>`<li>${esc(i)}</li>`).join('')}</ul>`:'<p class="note">No strong patterns detected automatically. Try the Charts or Ask tabs.</p>'}
 <h3 style="margin:14px 0 6px">Columns</h3><div class="scroll"><table class="stack"><thead><tr><th>Column</th><th>Type</th><th>Missing</th><th>Summary</th></tr></thead><tbody>${rows}</tbody></table></div>`}
async function autoCharts(){charts.forEach(c=>c.chart&&c.chart.destroy());charts=[];$('dt-charts').innerHTML='';
 const num=nums(),cat=T.cols.filter(c=>c.type==='category'&&new Set(c.v.filter(v=>v!=null)).size>=2),date=T.cols.find(c=>c.type==='date');const sp=[];
 if(num[0]&&num[1])sp.push({type:'scatter',x:num[0].name,y:num[1].name});
 if(num[0])sp.push({type:'hist',x:num[0].name});
 if(cat[0]&&num[0])sp.push({type:'bar',x:cat[0].name,y:num[0].name,agg:'sum',top:10});
 if(date&&num[0])sp.push({type:'line',x:date.name,y:num[0].name,agg:'sum'});
 if(!num.length&&cat[0])sp.push({type:'bar',x:cat[0].name,y:'(count)',top:10});
 for(const s of sp)await addChart(s,null,true)}
function fillSelects(){const opt=(c,sel)=>`<option ${sel?'selected':''}>${esc(c.name)}</option>`;
 const x=T.cols.find(c=>c.type==='category')||T.cols[0],y=nums()[0];
 $('cb-x').innerHTML=T.cols.map(c=>opt(c,c===x)).join('');$('cb-y').innerHTML='<option>(count)</option>'+T.cols.filter(c=>c.type==='number').map(c=>opt(c,c===y)).join('')}
async function load(ds,label){DS=ds;$('dt-sheet').innerHTML=ds.map((s,i)=>`<option value="${i}">${esc(s.name)}</option>`).join('');sheetIdx=0;$('dt-src').textContent='Analyzing: '+label;await pick()}
async function pick(){T=build(DS[sheetIdx]);findings=null;$('rp-out').innerHTML='';$('aq-out').innerHTML='';if(!T){$('dt-ov').innerHTML='<p class="note">This sheet has no data rows.</p>';return}overview();fillSelects();await autoCharts()}
$('dt-sheet').onchange=e=>{sheetIdx=+e.target.value;pick()};

/* ---------- sources ---------- */
function fromReader(){const I=R.info();if(I.kind!=='xlsx'||selfLoad)return;const raw=R.raw();if(raw&&raw.length)load(raw,I.docName)}
document.addEventListener('ras-doc',()=>{if(selfLoad){selfLoad=false;return}fromReader()});
setTimeout(fromReader,0);
$('dt-file').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{await lib(CDN.xlsx);const wb=XLSX.read(await f.arrayBuffer());load(wb.SheetNames.map(name=>{const ws=wb.Sheets[name];return{name,rows:XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:'',blankrows:false}),raw:XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:'',blankrows:false})}}),f.name)}catch(err){$('dt-src').textContent='Could not read that file: '+err.message}};
$('dt-sample').onclick=()=>{let s=7;const rnd=()=>(s=(s*16807)%2147483647)/2147483647;const reg={North:1.2,South:.9,East:1.05,West:.8},prod={Alpha:40,Beta:65,Gamma:120};const rows=[['Month','Region','Product','Units','Revenue','Cost']];
 for(let m=0;m<22;m++){const d=new Date(Date.UTC(2025,m,1)).toISOString().slice(0,7);for(const r in reg)for(const p in prod){const u=Math.round((50+m*2.2)*reg[r]*(0.75+rnd()*.5)*(p==='Gamma'?.5:1));const rev=Math.round(u*prod[p]*(0.95+rnd()*.1));rows.push([d,r,p,u,rev,Math.round(rev*(0.55+rnd()*.15))])}}
 load([{name:'Sales',rows,raw:rows}],'Sample sales data (Jan 2025 – Oct 2026, made up for demonstration)')};

/* ---------- tabs + builder ---------- */
box.querySelector('.tabs').onclick=e=>{const b=e.target.closest('[data-t]');if(!b)return;box.querySelectorAll('[role=tab]').forEach(t=>t.setAttribute('aria-selected',t===b));box.querySelectorAll('[data-p]').forEach(p=>p.hidden=p.dataset.p!==b.dataset.t);if(b.dataset.t==='ch')requestAnimationFrame(()=>charts.forEach(c=>{try{c.chart&&c.chart.resize()}catch(e){}}))};
$('cb-go').onclick=async()=>{if(!T)return;const r=await addChart({type:$('cb-type').value,x:$('cb-x').value,y:$('cb-y').value,agg:$('cb-agg').value,top:$('cb-top').value},null,true);if(typeof r==='string')alertNote(r)};
function alertNote(t){const p=document.createElement('p');p.className='note';p.textContent=t;$('dt-charts').prepend(p);setTimeout(()=>p.remove(),6000)}

/* ---------- AI helpers ---------- */
function parse(t){t=String(t||'').replace(/[“”]/g,'"');const a=t.indexOf('{'),b=t.lastIndexOf('}');if(a<0||b<a)return null;try{return JSON.parse(t.slice(a,b+1))}catch(e){try{return JSON.parse(t.slice(a,b+1).replace(/,\s*([}\]])/g,'$1'))}catch(e2){return null}}}
const ai=async(p,deep)=>(await R.ask([{role:'user',content:p}],{modelTier:deep?'complex':'default',cache:false})).text;
const errMsg=e=>e&&e.code==='rate_limited'?'The free AI allowance is used up for today. Try again tomorrow.':e&&e.code==='unavailable'?'The AI helper is not switched on for this page.':'The AI did not answer. Try again.';
function schema(){return T.cols.map(c=>{if(c.type==='number'){const s=stats(c.v);return`- "${c.name}" (number, ${fmt(s.min)} to ${fmt(s.max)})`}if(c.type==='date'){const t=c.v.filter(Boolean).map(Number);return`- "${c.name}" (date, ${new Date(Math.min(...t)).toISOString().slice(0,10)} to ${new Date(Math.max(...t)).toISOString().slice(0,10)})`}const k=counts(c.v);return`- "${c.name}" (${c.type}, values like: ${k.slice(0,8).map(x=>x[0]).join(', ')})`}).join('\n')}

/* ---------- ask the data ---------- */
$('aq-go').onclick=async()=>{const q=$('aq-q').value.trim();if(!q||!T)return;$('aq-tip').textContent='Working it out…';
 try{const j=parse(await ai(`You turn questions about a table into a calculation spec. The table "${T.name}" has ${T.n} rows and these columns:
${schema()}
Question: "${q}"
Return ONLY valid JSON: {"type":"bar|line|pie|hist|scatter","x":"column to group by (exact name)","y":"number column (exact name) or (count)","agg":"sum|mean|count|max|min","sort":"desc|asc","top":10,"filters":[{"col":"exact column","op":"=|!=|>|<|>=|<=|contains","value":"..."}],"title":"short chart title"}
Use "line" with a date column for trends over time. Use exact column names and exact category values from the list. Use "filters":[] if none.`));
  if(!j||!j.x)throw{code:'spec'};const res=compute(j);
  const rows=res.kind==='scatter'?[]:res.labels.map((l,i)=>[l,res.values[i]]);
  const table=res.kind==='scatter'?`Correlation between ${res.xl} and ${res.yl}: ${res.r==null?'not enough data':nf.format(res.r)} (${res.pts.length} rows)`:rows.slice(0,25).map(r=>`${r[0]}: ${fmt(r[1])}`).join('\n');
  const ans=await ai(`Answer the question in 1-3 plain sentences using ONLY the computed results below. Quote numbers exactly as given. If the results cannot answer it, say what they show instead.
Question: "${q}"
Calculation: ${res.yl} by ${res.xl}${(j.filters||[]).length?' where '+j.filters.map(f=>`${f.col} ${f.op} ${f.value}`).join(' and '):''}
Results:
${table}`);
  $('aq-out').innerHTML=`<div class="ans"><b>${esc(q)}</b><p style="margin:6px 0 0">${esc(ans.trim())}</p><p class="note" style="margin:6px 0 0">Calculated from ${res.kind==='scatter'?res.pts.length:T.n} rows: ${esc(res.yl)} by ${esc(res.xl)}${(j.filters||[]).length?', filtered: '+esc(j.filters.map(f=>`${f.col} ${f.op} ${f.value}`).join(', ')):''}.</p></div>`+(rows.length?`<div class="scroll"><table><thead><tr><th>${esc(res.xl)}</th><th style="text-align:right">${esc(res.yl)}</th></tr></thead><tbody>${rows.slice(0,25).map(r=>`<tr><td>${esc(r[0])}</td><td class="n">${fmt(r[1])}</td></tr>`).join('')}</tbody></table></div>`:'');
  await addChart(j,j.title||q,true);$('aq-tip').textContent='Chart added to the Charts tab.'}
 catch(e){$('aq-tip').textContent=e&&e.code==='spec'?'I could not turn that into a calculation. Try naming a column.':e&&e.message&&!e.code?e.message:errMsg(e)}};
$('aq-q').addEventListener('keydown',e=>{if(e.key==='Enter')$('aq-go').click()});

/* ---------- report ---------- */
function digest(){const lines=[`Table "${T.name}": ${T.n} rows, ${T.cols.length} columns.`];T.cols.forEach(c=>{if(c.type==='number'){const s=stats(c.v);lines.push(`${c.name}: total ${fmt(s.sum)}, average ${fmt(s.mean)}, median ${fmt(s.median)}, range ${fmt(s.min)}–${fmt(s.max)}.`)}else if(c.type==='category'){lines.push(`${c.name}: ${counts(c.v).slice(0,6).map(([v,n])=>`${v} (${n} rows)`).join(', ')}.`)}});
 lines.push('Automatic insights:',...(T.ins||[]).map(i=>'- '+i));charts.filter(c=>c.inReport&&c.res.kind!=='scatter').slice(0,6).forEach(c=>lines.push(`Chart "${c.title}": ${c.res.labels.slice(0,12).map((l,i)=>`${l}=${fmt(c.res.values[i])}`).join(', ')}`));return lines.join('\n').slice(0,12000)}
$('rp-go').onclick=async()=>{if(!T)return;$('rp-tip').textContent='Writing findings…';
 try{const j=parse(await ai(`You are a senior data analyst writing for: ${$('rp-aud').value.trim()||'a general business audience'}. Using ONLY the facts below (quote numbers exactly; do not invent data), write the findings.
Return ONLY valid JSON: {"title":"report title","summary":"3-4 sentence executive summary","findings":["finding with numbers", "..."],"recommendations":["actionable recommendation","..."],"caveats":["limitation of the data or analysis","..."]}
Give 4-6 findings, 3-4 recommendations, 1-3 caveats.
FACTS:
${digest()}`,true));if(!j||!j.summary)throw{code:'x'};findings=j;renderFindings();$('rp-tip').textContent=''}
 catch(e){$('rp-tip').textContent=errMsg(e)}};
function renderFindings(){const f=findings;$('rp-out').innerHTML=`<div class="ans"><h3 style="margin:0 0 6px">${esc(f.title||'Findings')}</h3><p>${esc(f.summary)}</p>${[['Key findings',f.findings],['Recommendations',f.recommendations],['Caveats',f.caveats]].map(([h,a])=>a&&a.length?`<b>${h}</b><ul class="ins">${a.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>`:'').join('')}</div>`}
$('rp-listen').onclick=()=>{if(!findings){$('rp-tip').textContent='Write findings first.';return}selfLoad=true;const f=findings;R.load((f.title||'Findings')+'.txt',[{label:f.title||'Findings',paras:[f.summary]},{label:'Key findings',paras:f.findings||[]},{label:'Recommendations',paras:f.recommendations||[]},{label:'Caveats',paras:f.caveats||[]}].filter(s=>s.paras.length));$('reader').scrollIntoView({behavior:'smooth'})};
const slug=t=>String(t||'report').replace(/[^\w\- ]+/g,'').trim().slice(0,60)||'report';
function save(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},4000)}
const b64bytes=u=>Uint8Array.from(atob(u.split(',')[1]),c=>c.charCodeAt(0));
$('rp-export').onclick=async()=>{if(!T){$('rp-xtip').textContent='Load data first.';return}const fmtx=$('rp-fmt').value,f=findings||{title:`Analysis of ${T.name}`,summary:`${T.n} rows and ${T.cols.length} columns analyzed.`,findings:T.ins||[],recommendations:[],caveats:[]};
 $('rp-xtip').textContent='Creating the file…';const btn=$('rp-export');btn.disabled=true;
 try{const sel=charts.filter(c=>c.inReport).slice(0,10).reverse(),imgs=[];for(const c of sel)imgs.push({title:c.title,url:await png(c.res,c.title,fmtx==='pdf'?'image/jpeg':'image/png'),res:c.res});
  const statRows=T.cols.filter(c=>c.type==='number').map(c=>{const s=stats(c.v);return[c.name,s.n,s.sum,s.mean,s.median,s.sd,s.min,s.max]}),statHead=['Column','Count','Total','Average','Median','Std dev','Min','Max'];
  const sections=[['Executive summary',[f.summary]],['Key findings',f.findings||[]],['Recommendations',f.recommendations||[]],['Caveats',f.caveats||[]]].filter(s=>s[1].length);
  const title=f.title||`Analysis of ${T.name}`;let blob,ext=fmtx;
  if(fmtx==='docx'){await lib(CDN.docx);const D=window.docx,k=[new D.Paragraph({text:title,heading:D.HeadingLevel.TITLE})];
   sections.forEach(([h,a])=>{k.push(new D.Paragraph({text:h,heading:D.HeadingLevel.HEADING_1}));a.forEach(x=>k.push(new D.Paragraph({children:[new D.TextRun({text:x,size:24})],bullet:h==='Executive summary'?undefined:{level:0},spacing:{after:120}})))});
   if(imgs.length){k.push(new D.Paragraph({text:'Charts',heading:D.HeadingLevel.HEADING_1}));imgs.forEach(i=>k.push(new D.Paragraph({children:[new D.ImageRun({data:b64bytes(i.url),transformation:{width:600,height:370}})],spacing:{after:200}})))}
   if(statRows.length){k.push(new D.Paragraph({text:'Summary statistics',heading:D.HeadingLevel.HEADING_1}));k.push(new D.Table({width:{size:100,type:D.WidthType.PERCENTAGE},rows:[statHead,...statRows].map((r,ri)=>new D.TableRow({children:r.map(v=>new D.TableCell({children:[new D.Paragraph({children:[new D.TextRun({text:typeof v==='number'?fmt(v):String(v),bold:ri===0,size:18})]})]}))}))}))}
   blob=await D.Packer.toBlob(new D.Document({creator:'Read Aloud Studio',title,sections:[{children:k}]}))}
  else if(fmtx==='pdf'){await lib(CDN.pdf);const doc=new window.jspdf.jsPDF({unit:'pt',format:'letter'}),M=54,W=doc.internal.pageSize.getWidth()-2*M,H=doc.internal.pageSize.getHeight();let y=M;
   const put=(t,sz,b,gap)=>{doc.setFont('helvetica',b?'bold':'normal');doc.setFontSize(sz);for(const l of doc.splitTextToSize(String(t).replace(/[^\x00-ɏ‐-‧]/g,''),W)){if(y+sz>H-M){doc.addPage();y=M}doc.text(l,M,y);y+=sz*1.4}y+=gap};
   put(title,20,true,8);sections.forEach(([h,a])=>{put(h,14,true,4);a.forEach(x=>put((h==='Executive summary'?'':'• ')+x,11,false,6))});
   imgs.forEach(i=>{const h=W*740/1200;if(y+h>H-M){doc.addPage();y=M}doc.addImage(i.url,'JPEG',M,y,W,h);y+=h+16});
   if(statRows.length){put('Summary statistics',14,true,4);statRows.forEach(r=>put(`${r[0]}: total ${fmt(r[2])}, average ${fmt(r[3])}, median ${fmt(r[4])}, range ${fmt(r[6])}–${fmt(r[7])}`,10,false,3))}blob=doc.output('blob')}
  else if(fmtx==='pptx'){await lib(CDN.pptx);const P=new window.PptxGenJS();P.layout='LAYOUT_WIDE';const t=P.addSlide();t.background={color:'0A2540'};t.addText(title,{x:.7,y:2.3,w:12,h:1.4,fontSize:38,bold:true,color:'FFFFFF',fontFace:'Arial'});t.addText(`${T.n} rows · ${T.name}`,{x:.7,y:3.8,w:12,h:.5,fontSize:16,color:'8FD3FF',fontFace:'Arial'});
   const bul=(h,a)=>{if(!a.length)return;const s=P.addSlide();s.addText(h,{x:.6,y:.4,w:12,h:.9,fontSize:28,bold:true,color:'0A2540',fontFace:'Arial'});s.addText(a.slice(0,6).map(x=>({text:x,options:{bullet:h!=='Executive summary',breakLine:true}})),{x:.6,y:1.5,w:12.1,h:5.5,fontSize:18,color:'1A1A1A',fontFace:'Arial',valign:'top',paraSpaceAfter:10})};
   sections.forEach(([h,a])=>bul(h,a));imgs.forEach(i=>{const s=P.addSlide();s.addImage({data:i.url.replace(/^data:/,''),x:.9,y:.45,w:11.5,h:7.1*.95})});blob=await P.write({outputType:'blob'})}
  else if(fmtx==='xlsx'){await lib(CDN.xlsx);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([[title],[],...sections.flatMap(([h,a])=>[[h],...a.map(x=>[x]),[]])]),'Findings');
   if(statRows.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([statHead,...statRows]),'Statistics');
   imgs.forEach((i,n)=>{if(i.res.kind!=='scatter')XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([[i.res.xl,i.res.yl],...i.res.labels.map((l,k)=>[l,i.res.values[k]])]),('Chart '+(n+1)).slice(0,31))});
   blob=new Blob([XLSX.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
  else{const h=`<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>body{font:17px/1.6 system-ui,sans-serif;max-width:860px;margin:40px auto;padding:0 20px;color:#1a1a1a}img{max-width:100%;border:1px solid #ddd;border-radius:8px;margin:10px 0}table{border-collapse:collapse;font-size:14px;width:100%}td,th{border:1px solid #ccc;padding:6px}td:not(:first-child){text-align:right}</style></head><body><h1>${esc(title)}</h1>${sections.map(([hh,a])=>`<h2>${hh}</h2>`+(hh==='Executive summary'?`<p>${esc(a[0])}</p>`:`<ul>${a.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>`)).join('')}${imgs.length?'<h2>Charts</h2>'+imgs.map(i=>`<img src="${i.url}" alt="${esc(i.title)}">`).join(''):''}${statRows.length?`<h2>Summary statistics</h2><table><tr>${statHead.map(x=>`<th>${x}</th>`).join('')}</tr>${statRows.map(r=>`<tr>${r.map(v=>`<td>${typeof v==='number'?fmt(v):esc(v)}</td>`).join('')}</tr>`).join('')}</table>`:''}</body></html>`;blob=new Blob([h],{type:'text/html'})}
  const name=slug(title)+'.'+ext;save(blob,name);$('rp-xtip').textContent=`Saved ${name} (${Math.max(1,Math.round(blob.size/1024))} KB) with ${imgs.length} chart${imgs.length===1?'':'s'}.`}
 catch(e){$('rp-xtip').textContent=e.message||'Could not create the file.'}finally{btn.disabled=false}};
})();
