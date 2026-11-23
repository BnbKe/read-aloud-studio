/* Read Aloud Studio — Improve & export module: Review, Edit (with approval), Convert/Export */
(function(){
const R=window.RAS;if(!R)return;
const $=id=>document.getElementById(id),esc=R.esc;
const CDN={docx:'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',pdf:'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',pptx:'https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js',xlsx:'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'};
const got={};const lib=u=>got[u]||(got[u]=new Promise((ok,no)=>{const s=document.createElement('script');s.src=u;s.onload=ok;s.onerror=()=>{delete got[u];no(new Error('A file library could not load. Check your connection and try again.'))};document.head.appendChild(s)}));

/* ---------- styles + markup ---------- */
const css=document.createElement('style');css.textContent=`
.pro{margin:0 0 40px}.pro h2{margin:0 0 4px}
.pro .tabs{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 14px}
.pro .tabs button[aria-selected=true]{background:var(--sky);color:#06213a;border-color:var(--sky);font-weight:600}
.pro .row{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,220px),1fr));gap:12px;margin-bottom:12px}
.pro input[type=text]{width:100%;font:inherit;font-size:16px;color:var(--ink);background:var(--deep);border:1px solid var(--line);border-radius:12px;min-height:44px;padding:0 12px}
.pro .go{background:var(--sky);color:#06213a;border-color:var(--sky);font-weight:600}
.pro .bar2{height:4px;background:var(--line);border-radius:4px;margin:12px 0 6px;overflow:hidden}.pro .bar2 i{display:block;height:100%;width:0;background:var(--pink);transition:width .3s}
.pro .card{background:var(--deep);border:1px solid var(--line);border-radius:14px;padding:12px 14px;margin:10px 0}
.pro .card h3{font-size:17px;margin:0 0 6px;font-weight:600}
.pro .tag{display:inline-block;font-size:12px;padding:2px 8px;border-radius:99px;background:var(--line);color:var(--sky);margin-right:6px;text-transform:capitalize}
.pro .q{font-style:italic;color:var(--soft)}
.pro .fix{margin-top:4px}.pro .fix b{color:var(--sky);font-weight:600}
.pro .acts2{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}.pro .acts2 button{min-height:36px;font-size:14px;padding:0 12px}
.pro del{background:rgba(255,92,138,.22);color:#ffb3c7;text-decoration:line-through}
.pro ins{background:rgba(143,211,255,.2);color:#cdeeff;text-decoration:none}
.pro .ok{border-color:var(--sky)}.pro .no{opacity:.55}
.pro ul{margin:4px 0 0;padding-left:20px}
.pro .big{font-size:28px;font-weight:300;color:var(--sky)}
@media(prefers-reduced-motion:reduce){.pro .bar2 i{transition:none}}`;
document.head.appendChild(css);
const box=document.createElement('section');box.className='panel pro';box.id='pro';
box.innerHTML=`<h2>Improve &amp; export</h2>
<p class="how">Get expert recommendations on a written paper, let AI edit it with your approval, and save any version in the format you need.</p>
<div class="tabs" role="tablist">
 <button role="tab" aria-selected="true" data-t="review">Review</button>
 <button role="tab" aria-selected="false" data-t="edit">Edit</button>
 <button role="tab" aria-selected="false" data-t="export">Convert &amp; export</button></div>
<div data-p="review">
 <div class="row"><label>Focus<select id="rv-focus">
  <option value="overall quality: clarity, structure, argument, grammar and style">Everything (overall)</option>
  <option value="clarity and readability">Clarity &amp; readability</option>
  <option value="structure, logical flow and strength of argument">Structure &amp; argument</option>
  <option value="grammar, spelling and punctuation">Grammar &amp; spelling</option>
  <option value="academic tone, evidence and citations">Academic tone &amp; citations</option>
  <option value="business writing: concise, action-oriented, professional">Business writing</option></select></label>
 <label>Audience or purpose (optional)<input type="text" id="rv-aud" placeholder="e.g. master's essay, grant proposal, client report"></label></div>
 <button class="go" id="rv-go">Review this document</button>
 <div class="bar2"><i id="rv-bar"></i></div><div class="tip" id="rv-tip"></div><div id="rv-out"></div></div>
<div data-p="edit" hidden>
 <div class="chips" id="ed-style">
  <button data-s="proof" aria-pressed="true">Proofread only</button><button data-s="clear" aria-pressed="false">Clearer</button>
  <button data-s="concise" aria-pressed="false">More concise</button><button data-s="formal" aria-pressed="false">More formal / academic</button>
  <button data-s="plain" aria-pressed="false">Plain language</button><button data-s="review" aria-pressed="false">Apply review advice</button>
  <button data-s="custom" aria-pressed="false">Custom</button></div>
 <input type="text" id="ed-custom" placeholder="Your instruction, e.g. 'make it persuasive for investors'" hidden style="margin-bottom:12px">
 <button class="go" id="ed-go">Edit this document</button>
 <div class="bar2"><i id="ed-bar"></i></div><div class="tip" id="ed-tip"></div>
 <div class="acts2" id="ed-acts" hidden><button id="ed-all">Accept all</button><button id="ed-none">Reject all</button><button id="ed-listen">Listen to edited version</button><button id="ed-export">Export edited version</button></div>
 <div id="ed-out"></div></div>
<div data-p="export" hidden>
 <div class="row"><label>What to export<select id="ex-src"><option value="orig">Original document</option><option value="edit">Edited version</option><option value="review">Review report</option></select></label>
 <label>Format<select id="ex-fmt"><option value="docx">Word (.docx)</option><option value="pdf">PDF</option><option value="pptx">PowerPoint (.pptx)</option><option value="xlsx">Excel (.xlsx)</option><option value="csv">CSV</option><option value="md">Markdown (.md)</option><option value="html">Web page (.html)</option><option value="txt">Plain text (.txt)</option></select></label></div>
 <label class="tog"><input type="checkbox" id="ex-ai"> Let AI restructure it for the format (paper → slides, slides → report, notes → polished document)</label>
 <button class="go" id="ex-go">Create file</button><div class="tip" id="ex-tip"></div></div>`;
document.querySelector('.grid').after(box);

/* ---------- helpers ---------- */
let ctrl=null,review=null,edits=null,base=null,style='proof';
const info=()=>R.info(),stem=()=>(info().docName||'document').replace(/ \(edited\)$/,'').replace(/\.[a-z0-9]+$/i,'').replace(/[^\w\- ]+/g,'').trim()||'document';
function units(secs,max=2400){const out=[];secs.forEach((s,si)=>{let cur=[],len=0,start=0;s.paras.forEach((p,pi)=>{if(len+p.length>max&&cur.length){out.push({si,label:s.label||'Start',start,paras:cur});cur=[];len=0;start=pi}cur.push(p);len+=p.length});if(cur.length)out.push({si,label:s.label||'Start',start,paras:cur})});return out}
function parse(t){t=String(t||'').replace(/[“”]/g,'"');const a=t.indexOf('{'),b=t.lastIndexOf('}');if(a<0||b<a)return null;try{return JSON.parse(t.slice(a,b+1))}catch(e){try{return JSON.parse(t.slice(a,b+1).replace(/,\s*([}\]])/g,'$1'))}catch(e2){return null}}}
async function ai(prompt,deep){const r=await R.ask([{role:'user',content:prompt}],{modelTier:deep?'complex':'default',cache:false,signal:ctrl&&ctrl.signal});return r.text}
const errMsg=e=>e&&e.code==='rate_limited'?'The free AI allowance is used up for today. Try again tomorrow.':e&&e.code==='unavailable'?'The AI helper is not switched on for this page.':e&&e.name==='AbortError'?'Stopped.':'The AI did not answer. Try again.';
function busy(btn,on,label){btn.textContent=on?'Stop':label;btn.dataset.busy=on?'1':''}
function words(a){return a.split(/\s+/).filter(Boolean)}
function diff(a,b){const x=words(a),y=words(b),n=x.length,m=y.length;if(n*m>250000)return`<del>${esc(a)}</del> <ins>${esc(b)}</ins>`;
 const d=Array.from({length:n+1},()=>new Uint16Array(m+1));for(let i=n-1;i>=0;i--)for(let j=m-1;j>=0;j--)d[i][j]=x[i]===y[j]?d[i+1][j+1]+1:Math.max(d[i+1][j],d[i][j+1]);
 let i=0,j=0,o=[];while(i<n&&j<m){if(x[i]===y[j]){o.push(esc(x[i]));i++;j++}else if(d[i+1][j]>=d[i][j+1])o.push(`<del>${esc(x[i++])}</del>`);else o.push(`<ins>${esc(y[j++])}</ins>`)}
 while(i<n)o.push(`<del>${esc(x[i++])}</del>`);while(j<m)o.push(`<ins>${esc(y[j++])}</ins>`);return o.join(' ').replace(/<\/del> <del>/g,' ').replace(/<\/ins> <ins>/g,' ')}

/* ---------- tabs ---------- */
function tab(t){box.querySelectorAll('[role=tab]').forEach(b=>b.setAttribute('aria-selected',b.dataset.t===t));box.querySelectorAll('[data-p]').forEach(p=>p.hidden=p.dataset.p!==t)}
box.querySelector('.tabs').onclick=e=>{const b=e.target.closest('[data-t]');if(b)tab(b.dataset.t)};
document.addEventListener('ras-doc',()=>{if(ctrl)ctrl.abort();review=null;edits=null;base=null;['rv-out','ed-out','rv-tip','ed-tip','ex-tip'].forEach(id=>$(id).innerHTML='');$('ed-acts').hidden=true;$('rv-bar').style.width=$('ed-bar').style.width='0'});

/* ---------- REVIEW ---------- */
$('rv-go').onclick=async()=>{
 if(ctrl){ctrl.abort();return}
 const secs=R.fullSecs(),us=units(secs,3000).slice(0,12);if(!us.length){$('rv-tip').textContent='Open a document first.';return}
 const I=info(),focus=$('rv-focus').value,aud=$('rv-aud').value.trim();
 ctrl=new AbortController();busy($('rv-go'),true,'Review this document');review={parts:[],overall:null};$('rv-out').innerHTML='';
 try{for(let k=0;k<us.length;k++){const u=us[k];$('rv-tip').textContent=`Reviewing part ${k+1} of ${us.length}: ${u.label}…`;$('rv-bar').style.width=(k/(us.length+1)*100)+'%';
   const t=await ai(`You are an expert writing reviewer. Review this part of a ${I.type} document called "${I.docName}". Focus on ${focus}. Audience or purpose: ${aud||'not stated, infer it'}.
Return ONLY valid JSON, no other text, in exactly this shape:
{"summary":"one sentence on how this part reads","suggestions":[{"type":"clarity|structure|grammar|style|evidence|citation|tone","quote":"exact words copied from the text, at most 12 words","issue":"what is wrong, in one short sentence","fix":"a concrete rewrite or action"}]}
Give at most 5 suggestions, most important first. Only point out real problems; if this part is strong, give fewer.
TEXT (${u.label}):
${u.paras.join('\n\n')}`);
   const j=parse(t)||{summary:'',suggestions:[]};review.parts.push({u,summary:j.summary||'',items:(j.suggestions||[]).filter(s=>s&&s.issue).slice(0,5)});renderReview()}
  $('rv-tip').textContent='Writing the overall assessment…';
  const digest=review.parts.map(p=>`${p.u.label}: ${p.summary} Issues: ${p.items.map(s=>s.type+' – '+s.issue).join('; ')}`).join('\n').slice(0,9000);
  const o=parse(await ai(`You are an expert writing reviewer. Based on these section notes for "${I.docName}" (focus: ${focus}; audience: ${aud||'not stated'}), give an overall assessment.
Return ONLY valid JSON: {"verdict":"2-3 sentences","strengths":["...","..."],"priorities":["most important fix","second","third"],"rating":7}
rating is 1-10 for how ready the document is for its audience.
NOTES:
${digest}`))||{};
  review.overall=o;$('rv-bar').style.width='100%';$('rv-tip').textContent=`Done: ${review.parts.reduce((a,p)=>a+p.items.length,0)} recommendations. Use "Edit → Apply review advice" to fix them, or export the report.`;renderReview()}
 catch(e){$('rv-tip').textContent=errMsg(e);if(review&&review.parts.length)renderReview()}
 finally{ctrl=null;busy($('rv-go'),false,'Review this document')}};
function renderReview(){if(!review)return;const o=review.overall;let h='';
 if(o&&o.verdict)h+=`<div class="card ok"><h3>Overall assessment${o.rating?` · <span class="big">${esc(o.rating)}</span>/10`:''}</h3><p style="margin:0">${esc(o.verdict)}</p>
 ${o.strengths&&o.strengths.length?`<p style="margin:8px 0 0"><b>Strengths</b></p><ul>${o.strengths.map(s=>`<li>${esc(s)}</li>`).join('')}</ul>`:''}
 ${o.priorities&&o.priorities.length?`<p style="margin:8px 0 0"><b>Fix first</b></p><ul>${o.priorities.map(s=>`<li>${esc(s)}</li>`).join('')}</ul>`:''}</div>`;
 review.parts.forEach(p=>{h+=`<div class="card"><h3>${esc(p.u.label)}</h3>${p.summary?`<p style="margin:0 0 6px;color:var(--soft)">${esc(p.summary)}</p>`:''}${p.items.length?p.items.map(s=>`<div style="margin:10px 0"><span class="tag">${esc(s.type||'note')}</span>${s.quote?`<button class="cite rv-find" data-q="${esc(s.quote)}">Find</button> <span class="q">“${esc(s.quote)}”</span>`:''}<div>${esc(s.issue)}</div>${s.fix?`<div class="fix"><b>Suggestion:</b> ${esc(s.fix)}</div>`:''}</div>`).join(''):'<p style="margin:0">No issues found here.</p>'}</div>`});
 $('rv-out').innerHTML=h}
$('rv-out').onclick=e=>{const b=e.target.closest('.rv-find');if(b&&!R.find(b.dataset.q))$('rv-tip').textContent='Could not find that exact wording in the current reading view.'};

/* ---------- EDIT ---------- */
const STY={proof:'Fix only spelling, grammar and punctuation errors. Change as little as possible.',clear:'Improve clarity and flow: simpler sentence structure, clear transitions, no jargon where avoidable.',concise:'Make it more concise: remove redundancy and filler, keep every key point.',formal:'Make the tone more formal and academic: precise vocabulary, no contractions or casual phrases.',plain:'Rewrite in plain language a general reader understands easily: short sentences, everyday words.'};
$('ed-style').onclick=e=>{const b=e.target.closest('[data-s]');if(!b)return;style=b.dataset.s;[...$('ed-style').children].forEach(c=>c.setAttribute('aria-pressed',c===b));$('ed-custom').hidden=style!=='custom';if(style==='review'&&!review)$('ed-tip').textContent='Run a review first (Review tab), then come back.';else $('ed-tip').textContent=''};
$('ed-go').onclick=async()=>{
 if(ctrl){ctrl.abort();return}
 if(style==='review'&&!review){$('ed-tip').textContent='Run a review first (Review tab), then come back.';return}
 const ins=style==='custom'?($('ed-custom').value.trim()||STY.clear):style==='review'?'Apply the reviewer recommendations listed below where they fit; otherwise only fix errors.':STY[style];
 base=R.fullSecs();const us=units(base,2200);if(!us.length){$('ed-tip').textContent='Open a document first.';return}
 if(us.length>25&&!confirmBig(us.length))return;
 ctrl=new AbortController();busy($('ed-go'),true,'Edit this document');edits=[];$('ed-out').innerHTML='';$('ed-acts').hidden=true;
 try{for(let k=0;k<us.length;k++){const u=us[k];$('ed-tip').textContent=`Editing part ${k+1} of ${us.length}: ${u.label}…`;$('ed-bar').style.width=(k/us.length*100)+'%';
   const rp=style==='review'&&review?review.parts.find(p=>p.u.si===u.si&&p.u.start<=u.start+u.paras.length-1)||null:null;
   const advice=rp&&rp.items.length?'\nReviewer recommendations for this part:\n'+rp.items.map(s=>`- ${s.issue} Suggestion: ${s.fix}`).join('\n'):'';
   const t=await ai(`You are a careful professional editor. ${ins}${advice}
Keep the author's meaning, facts, numbers, names, quotations and citations exactly. Do not add new claims.
Return ONLY valid JSON: {"paragraphs":["edited paragraph 1", ...]} with exactly ${u.paras.length} items, one per input paragraph, in the same order.
PARAGRAPHS:
${u.paras.map((p,i)=>`[${i+1}] ${p}`).join('\n\n')}`,true);
   const j=parse(t),out=j&&Array.isArray(j.paragraphs)?j.paragraphs:[];
   u.paras.forEach((p,i)=>{const n=typeof out[i]==='string'?out[i].replace(/^\[\d+\]\s*/,'').trim():'';if(n&&n!==p.trim())edits.push({si:u.si,pi:u.start+i,label:u.label,orig:p,edit:n,state:'pending'})});
   renderEdits()}
  $('ed-bar').style.width='100%';$('ed-tip').textContent=edits.length?`${edits.length} paragraphs changed. Accept or reject each one, then listen or export.`:'No changes were needed.'}
 catch(e){$('ed-tip').textContent=errMsg(e)}
 finally{ctrl=null;busy($('ed-go'),false,'Edit this document');renderEdits()}};
function confirmBig(n){$('ed-tip').textContent=`This is a long document (${n} parts). Editing all of it may use up today's free AI allowance. Press Edit again to continue.`;if($('ed-go').dataset.big==='1'){$('ed-go').dataset.big='';return true}$('ed-go').dataset.big='1';return false}
function renderEdits(){if(!edits)return;$('ed-acts').hidden=!edits.length;
 $('ed-out').innerHTML=edits.map((x,i)=>`<div class="card ${x.state==='accepted'?'ok':x.state==='rejected'?'no':''}"><h3>${esc(x.label)} <span class="tag">${x.state}</span></h3><div>${diff(x.orig,x.edit)}</div><div class="acts2"><button data-i="${i}" data-a="accepted">Accept</button><button data-i="${i}" data-a="rejected">Reject</button></div></div>`).join('')}
$('ed-out').onclick=e=>{const b=e.target.closest('[data-a]');if(!b)return;edits[+b.dataset.i].state=b.dataset.a;renderEdits()};
$('ed-all').onclick=()=>{edits.forEach(x=>x.state='accepted');renderEdits()};
$('ed-none').onclick=()=>{edits.forEach(x=>x.state='rejected');renderEdits()};
function editedSecs(){if(!base)return null;const secs=base.map(s=>({label:s.label,paras:s.paras.slice()}));(edits||[]).forEach(x=>{if(x.state==='accepted')secs[x.si].paras[x.pi]=x.edit});return secs}
$('ed-listen').onclick=()=>{if(!edits.some(x=>x.state==='accepted')){$('ed-tip').textContent='Accept at least one change first.';return}const keep={review,edits,base};R.load(stem()+' (edited)',editedSecs());review=keep.review;edits=keep.edits;base=keep.base;renderReview();renderEdits();$('reader').scrollIntoView({behavior:'smooth'})};
$('ed-export').onclick=()=>{$('ex-src').value='edit';tab('export')};

/* ---------- EXPORT ---------- */
function content(src){const I=info();
 if(src==='review'){if(!review||!review.parts.length)throw new Error('Run a review first.');const o=review.overall||{},rows=[];
  const sec=[{heading:'Overall assessment',paras:[o.verdict||'',o.rating?`Rating: ${o.rating}/10`:'',...(o.strengths||[]).map(s=>'Strength: '+s),...(o.priorities||[]).map((s,i)=>`Priority ${i+1}: ${s}`)].filter(Boolean)}];
  review.parts.forEach(p=>{sec.push({heading:p.u.label,paras:[p.summary,...p.items.map(s=>`[${s.type}] ${s.quote?`"${s.quote}" — `:''}${s.issue} Suggestion: ${s.fix}`)].filter(Boolean)});p.items.forEach(s=>rows.push([p.u.label,s.type,s.quote||'',s.issue,s.fix||'']))});
  return{title:'Review: '+I.docName,sections:sec,table:{headers:['Section','Type','Quote','Issue','Suggestion'],rows}}}
 const secs=src==='edit'?editedSecs():R.fullSecs();if(!secs)throw new Error('Edit the document first.');
 let table=null;if(I.kind==='xlsx'&&src==='orig'){const sh=(R.raw()||[])[0];if(sh&&sh.rows.length)table={headers:sh.rows[0].map(String),rows:sh.rows.slice(1)}}
 const nm=I.docName.replace(/ \(edited\)$/,'');return{title:src==='edit'?nm+' (edited)':I.docName,sections:secs.map(s=>({heading:s.label,paras:s.paras})),table}}
async function restructure(c,fmt){const text=c.sections.map(s=>(s.heading?`## ${s.heading}\n`:'')+s.paras.join('\n')).join('\n\n').slice(0,14000);
 if(fmt==='pptx'){const j=parse(await ai(`Turn this document into a clear presentation. Return ONLY valid JSON: {"title":"deck title","slides":[{"title":"slide title","bullets":["short point","..."],"notes":"what the speaker says, 2-3 sentences"}]}
Use 6 to 12 slides, at most 5 bullets each, each bullet under 15 words. Keep all key facts and numbers accurate.
DOCUMENT "${c.title}":
${text}`,true));if(!j||!Array.isArray(j.slides))throw new Error('The AI could not build slides this time. Try again.');return{deck:j}}
 const j=parse(await ai(`Rewrite this material as a polished, well-structured written document with an introduction, clear headed sections and a conclusion. Keep every fact and number accurate; do not invent new facts.
Return ONLY valid JSON: {"title":"document title","sections":[{"heading":"section heading","paragraphs":["...","..."]}]}
MATERIAL "${c.title}":
${text}`,true));if(!j||!Array.isArray(j.sections))throw new Error('The AI could not restructure it this time. Try again.');return{title:j.title||c.title,sections:j.sections.map(s=>({heading:s.heading||'',paras:(s.paragraphs||[]).map(String)})),table:c.table}}
function save(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},4000)}
const plain=c=>[c.title,'',...c.sections.flatMap(s=>[s.heading?s.heading.toUpperCase():'',...s.paras,''])].filter((x,i,a)=>x||a[i-1]).join('\n');
const mdcell=v=>String(v??'').replace(/\|/g,'\\|').replace(/\n/g,' ');
function md(c){let o=`# ${c.title}\n\n`;c.sections.forEach(s=>{if(s.heading)o+=`## ${s.heading}\n\n`;s.paras.forEach(p=>o+=p+'\n\n')});if(c.table&&c.table.rows.length){o+=`| ${c.table.headers.map(mdcell).join(' | ')} |\n|${c.table.headers.map(()=>'---').join('|')}|\n`+c.table.rows.slice(0,1000).map(r=>`| ${c.table.headers.map((_,i)=>mdcell(r[i])).join(' | ')} |`).join('\n')+'\n'}return o}
function html(c){return`<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(c.title)}</title><style>body{font:17px/1.6 Georgia,serif;max-width:720px;margin:40px auto;padding:0 20px;color:#1a1a1a}h1{font:600 30px/1.2 system-ui,sans-serif}h2{font:600 21px/1.3 system-ui,sans-serif;margin-top:32px}table{border-collapse:collapse;font:14px system-ui,sans-serif;width:100%}td,th{border:1px solid #ccc;padding:6px;text-align:left;vertical-align:top}</style></head><body><h1>${esc(c.title)}</h1>${c.sections.map(s=>(s.heading?`<h2>${esc(s.heading)}</h2>`:'')+s.paras.map(p=>`<p>${esc(p)}</p>`).join('')).join('')}${c.table&&c.table.rows.length?`<table><tr>${c.table.headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr>${c.table.rows.slice(0,2000).map(r=>`<tr>${c.table.headers.map((_,i)=>`<td>${esc(r[i]??'')}</td>`).join('')}</tr>`).join('')}</table>`:''}</body></html>`}
const tableOf=c=>c.table&&c.table.rows.length?c.table:{headers:['Section','Text'],rows:c.sections.flatMap(s=>s.paras.map(p=>[s.heading,p]))};
async function docxOut(c){await lib(CDN.docx);const D=window.docx,kids=[new D.Paragraph({text:c.title,heading:D.HeadingLevel.TITLE})];
 c.sections.forEach(s=>{if(s.heading)kids.push(new D.Paragraph({text:s.heading,heading:D.HeadingLevel.HEADING_1}));s.paras.forEach(p=>kids.push(new D.Paragraph({children:[new D.TextRun({text:p,size:24})],spacing:{after:160}})))});
 if(c.table&&c.table.rows.length){const rows=[c.table.headers,...c.table.rows.slice(0,500)].map((r,ri)=>new D.TableRow({children:c.table.headers.map((_,i)=>new D.TableCell({children:[new D.Paragraph({children:[new D.TextRun({text:String(r[i]??''),bold:ri===0,size:20})]})]}))}));kids.push(new D.Table({rows,width:{size:100,type:D.WidthType.PERCENTAGE}}))}
 return D.Packer.toBlob(new D.Document({creator:'Read Aloud Studio',title:c.title,sections:[{children:kids}]}))}
async function pdfOut(c){await lib(CDN.pdf);const doc=new window.jspdf.jsPDF({unit:'pt',format:'letter'}),M=60,W=doc.internal.pageSize.getWidth()-2*M,H=doc.internal.pageSize.getHeight();let y=M;
 const put=(t,size,bold,gap)=>{doc.setFont('helvetica',bold?'bold':'normal');doc.setFontSize(size);for(const line of doc.splitTextToSize(String(t).replace(/[^\x00-ɏ‐-‧]/g,''),W)){if(y+size>H-M){doc.addPage();y=M}doc.text(line,M,y);y+=size*1.4}y+=gap};
 put(c.title,20,true,10);c.sections.forEach(s=>{if(s.heading)put(s.heading,14,true,4);s.paras.forEach(p=>put(p,11,false,8))});
 if(c.table&&c.table.rows.length){put(c.table.headers.join('  |  '),10,true,2);c.table.rows.slice(0,300).forEach(r=>put(c.table.headers.map((_,i)=>r[i]??'').join('  |  '),10,false,2))}
 return doc.output('blob')}
async function pptxOut(c,deck){await lib(CDN.pptx);const P=new window.PptxGenJS();P.layout='LAYOUT_WIDE';const ink='0A2540',sky='1E88C8';
 const t=P.addSlide();t.background={color:ink};t.addText(deck?deck.title||c.title:c.title,{x:0.7,y:2.4,w:12,h:1.5,fontSize:40,bold:true,color:'FFFFFF',fontFace:'Arial'});t.addText('Made with Read Aloud Studio',{x:0.7,y:4,w:12,h:0.5,fontSize:16,color:'8FD3FF',fontFace:'Arial'});
 const slides=deck?deck.slides:c.sections.filter(s=>s.paras.length).slice(0,40).map(s=>({title:s.heading||c.title,bullets:s.paras.slice(0,6).map(p=>{const f=R.split(p)[0]||p;return f.length>160?f.slice(0,157)+'…':f}),notes:s.paras.join(' ').slice(0,1500)}));
 slides.forEach(s=>{const sl=P.addSlide();sl.addText(String(s.title||''),{x:0.6,y:0.4,w:12.1,h:0.9,fontSize:28,bold:true,color:ink,fontFace:'Arial'});sl.addShape(P.ShapeType.rect,{x:0.6,y:1.3,w:1.2,h:0.06,fill:{color:sky},line:{color:sky}});
  const b=(s.bullets||[]).map(String).filter(Boolean);if(b.length)sl.addText(b.map(x=>({text:x,options:{bullet:true,breakLine:true}})),{x:0.6,y:1.6,w:12.1,h:5.4,fontSize:20,color:'1A1A1A',fontFace:'Arial',valign:'top',paraSpaceAfter:10});if(s.notes)sl.addNotes(String(s.notes))});
 return P.write({outputType:'blob'})}
async function xlsxOut(c){await lib(CDN.xlsx);const t=tableOf(c),wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet([t.headers,...t.rows]);ws['!cols']=t.headers.map((h,i)=>({wch:Math.min(60,Math.max(10,...[h,...t.rows.slice(0,200).map(r=>r[i])].map(v=>String(v??'').length)))}));XLSX.utils.book_append_sheet(wb,ws,'Content');return new Blob([XLSX.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
const csvOut=c=>{const t=tableOf(c),q=v=>{v=String(v??'');return/[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v};return new Blob(['﻿'+[t.headers,...t.rows].map(r=>t.headers.map((_,i)=>q(r[i])).join(',')).join('\r\n')],{type:'text/csv'})};
$('ex-go').onclick=async()=>{
 if(ctrl){ctrl.abort();return}
 const src=$('ex-src').value,fmt=$('ex-fmt').value,useAI=$('ex-ai').checked;let c;
 try{c=content(src)}catch(e){$('ex-tip').textContent=e.message;return}
 ctrl=new AbortController();busy($('ex-go'),true,'Create file');
 try{let deck=null;if(useAI&&src!=='review'){$('ex-tip').textContent=fmt==='pptx'?'AI is designing the slides…':'AI is restructuring the document…';const r=await restructure(c,fmt);if(r.deck)deck=r.deck;else c=r}
  $('ex-tip').textContent='Creating the file…';
  const name=`${stem()}${src==='edit'?'-edited':src==='review'?'-review':''}${useAI?'-ai':''}.${fmt}`;
  const blob=fmt==='docx'?await docxOut(c):fmt==='pdf'?await pdfOut(c):fmt==='pptx'?await pptxOut(c,deck):fmt==='xlsx'?await xlsxOut(c):fmt==='csv'?csvOut(c):new Blob([fmt==='md'?md(c):fmt==='html'?html(c):plain(c)],{type:fmt==='html'?'text/html':fmt==='md'?'text/markdown':'text/plain'});
  save(blob,name);$('ex-tip').textContent=`Saved ${name} (${Math.max(1,Math.round(blob.size/1024))} KB). Check your downloads.`}
 catch(e){$('ex-tip').textContent=e&&e.message&&!e.code?e.message:errMsg(e)}
 finally{ctrl=null;busy($('ex-go'),false,'Create file')}};
})();
