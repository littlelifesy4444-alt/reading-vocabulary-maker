
let DATA=null;
const $=s=>document.querySelector(s);
function esc(s){return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
function render(){
  if(!DATA)return;
  $("#result").classList.remove("hidden");
  $("#detected").textContent="추정 난이도: "+DATA.detected_level;
  $("#note").textContent=DATA.source_note||"";
  $("#vocab").innerHTML=`<table><thead><tr><th>No.</th><th>Word</th><th>발음·품사</th><th>뜻</th><th>Syn./Ant.</th><th class="example">New Example</th></tr></thead><tbody>${
    DATA.vocabulary.map((v,i)=>`<tr><td>${i+1}</td><td><b>${esc(v.word)}</b></td><td>${esc(v.pronunciation)}<br>${esc(v.pos)}</td><td>${esc(v.meaning_ko)}</td><td>S: ${esc(v.synonym)}<br>A: ${esc(v.antonym)}</td><td>${esc(v.example)}</td></tr>`).join("")
  }</tbody></table>`;
  let last=0, html='<div class="qgrid">';
  DATA.questions.forEach(q=>{
    if(q.section!==last){html+=`<div class="qsection">${q.section}</div>`;last=q.section}
    html+=`<div class="q"><b>${q.number}.</b> ${esc(q.prompt)}${q.choices?.length?`<div class="choices">${q.choices.map(esc).join("　")}</div>`:""}
    <div class="qtools"><button onclick="regen(${q.number},'easier')">이 문제 쉽게</button><button onclick="regen(${q.number},'harder')">이 문제 어렵게</button></div></div>`;
  }); html+='</div>'; $("#test").innerHTML=html;
  $("#answer").innerHTML=`<div class="answer-row answer-head"><div>No.</div><div>정답</div><div>뜻</div><div>문제 해설</div></div>`+
    DATA.questions.map(q=>`<div class="answer-row"><div>${q.number}</div><div><b>${esc(q.answer)}</b></div><div>${esc(q.meaning_ko)}</div><div>${esc(q.explanation_ko)}</div></div>`).join("");
}
async function regen(number,difficulty){
  const r=await fetch("/api/regenerate-question",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({number,difficulty})});
  const j=await r.json(); if(!r.ok){alert(j.error||"오류");return}
  DATA.questions=DATA.questions.map(q=>q.number===number?j:q);render();
}
$("#generate").onclick=async()=>{
  const files=$("#files").files;if(!files.length){alert("파일을 선택해 주세요.");return}
  const fd=new FormData(); for(const f of files)fd.append("files",f);
  fd.append("title",$("#title").value);fd.append("page_range",$("#pageRange").value);fd.append("difficulty",$("#difficulty").value);
  $("#status").textContent="분석 중...";
  const r=await fetch("/api/generate",{method:"POST",body:fd});const j=await r.json();
  $("#status").textContent="";
  if(!r.ok){alert(j.error||"생성 실패");return} DATA=j;render();
}
$("#shuffle").onclick=async()=>{
  const r=await fetch("/api/reshuffle",{method:"POST"});const j=await r.json();if(!r.ok){alert(j.error);return}DATA=j;render();
}
document.querySelectorAll(".tabs button").forEach(b=>b.onclick=()=>{
  document.querySelectorAll(".tabs button").forEach(x=>x.classList.remove("active"));b.classList.add("active");
  document.querySelectorAll(".tab").forEach(x=>x.classList.add("hidden"));$("#"+b.dataset.tab).classList.remove("hidden");
});
