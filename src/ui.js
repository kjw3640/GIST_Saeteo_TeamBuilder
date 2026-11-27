"use strict";
const $=id=>document.getElementById(id);
const el=(tag,value)=>{const node=document.createElement(tag);if(value!==undefined)node.textContent=value;return node;};
let state=null,worker=null,workerURL="",busy=false,version=0;
const urls=new Set();
function message(s,error=false){$("status").textContent=s;$("status").classList.toggle("error",error);}
function stage(id){
  ["setup","review","result"].forEach((s,i)=>{const current=s===id;$(s).hidden=!current;const li=document.querySelectorAll(".steps li")[i];if(current)li.setAttribute("aria-current","step");else li.removeAttribute("aria-current");});
}
function stopWorker(){worker?.terminate();worker=null;if(workerURL)URL.revokeObjectURL(workerURL);workerURL="";busy=false;}
function manualMode(){return document.querySelector('input[name="mode"]:checked').value==="manual";}
function ready(){
  const manual=manualMode();
  $("manual").hidden=!manual;$("leaderPick").hidden=manual;
  for(const input of $("setup").querySelectorAll("input, #add, #manualRows button"))input.disabled=busy;
  $("read").disabled=busy||!$("people").files.length || (manual?!$("manualRows").children.length:!$("leaders").files.length);
}
function addLeader(){
  const body=$("manualRows");
  if(body.children.length>=120)return;
  const row=el("tr");
  const used=Array.from(body.rows,r=>Number(r.cells[0].firstChild.value));
  let id=1;while(used.includes(id))id++;
  for(const [i,label] of ["조 번호","조장1 이름","조장1 생년","조장2 이름","조장2 생년"].entries()){
    const td=el("td"),input=el("input");input.setAttribute("aria-label",label);
    td.dataset.label=label;input.autocomplete="off";input.spellcheck=false;input.maxLength=40;
    if([0,2,4].includes(i)){input.inputMode="numeric";input.maxLength=4;}
    if(!i)input.value=id;
    td.append(input);row.append(td);
  }
  const td=el("td"),del=el("button","삭제");del.type="button";del.className="secondary";del.onclick=()=>{row.remove();ready();};td.append(del);row.append(td);body.append(row);ready();
}
function manualData(){return [{number:1,cells:["조 번호","조장1 이름","조장1 생년","조장2 이름","조장2 생년"]},...Array.from($("manualRows").rows,(r,i)=>({number:i+2,cells:Array.from(r.querySelectorAll("input"),t=>t.value.trim())}))];}
function clearResult(){for(const id of ["counts","alerts","edits","summary","stats","scope","exportHint"])$(id).replaceChildren();}
function reset(){
  version++;stopWorker();state=null;
  for(const url of urls)URL.revokeObjectURL(url);urls.clear();
  for(const pick of document.querySelectorAll(".pick"))pick.classList.remove("selected");
  $("people").value="";$("leaders").value="";
  $("leaderLabel").textContent=$("peopleLabel").textContent="파일 선택";
  document.querySelector('input[value="file"]').checked=true;
  $("manualRows").replaceChildren();addLeader();
  $("year").value=new Date().getFullYear();clearResult();stage("setup");ready();message("");
}
$("reset").onclick=()=>{reset();message("명단과 편성 결과를 비웠습니다.");};
window.addEventListener("pagehide",reset);
for(const radio of document.querySelectorAll('input[name="mode"]'))radio.onchange=ready;
$("add").onclick=addLeader;
for(const [id,label] of [["leaders","leaderLabel"],["people","peopleLabel"]]){
  $(id).onchange=()=>{$(id).parentElement.classList.toggle("selected",!!$(id).files.length);$(label).textContent=$(id).files[0]?.name||"파일 선택";message("");ready();};
}
function table(headers,rows) {
  const t=el("table"),head=el("thead"),tr=el("tr");
  for(const h of headers)tr.append(el("th",h));head.append(tr);t.append(head);
  const body=el("tbody");for(const values of rows){const row=el("tr");for(const v of values)row.append(el("td",v));body.append(row);}t.append(body);return t;
}
function details(title){const d=el("details");d.append(el("summary",title));return d;}
function review(){
  const {people,leaders,notices}=state;
  $("counts").textContent=leaders.length+"개 조 · 배정 대상 "+people.length+"명";
  const alerts=$("alerts"),edits=$("edits");alerts.replaceChildren();edits.replaceChildren();
  if(notices.length){const d=details("입력 확인 사항 "+notices.length+"건");const list=el("ul");for(const s of notices)list.append(el("li",s));d.append(list);alerts.append(d);}
  const uncertain=people.filter(p=>p.needsEdit);
  if(uncertain.length){
    edits.append(el("p","학교·이름 확인이 필요한 "+uncertain.length+"개 행입니다. 학교를 모르면 비워둘 수 있습니다."));
    for(const p of uncertain){const box=el("div");box.className="edit";box.append(el("b",p.row+"행"));
      for(const [key,label] of [["name","이름"],["school","출신고교"]]){const l=el("label",label),inp=el("input");inp.autocomplete="off";inp.spellcheck=false;inp.value=p[key];inp.maxLength=100;inp.oninput=()=>{p[key]=inp.value.trim();};l.append(inp);box.append(l);}edits.append(box);}
  }
  const d=details("인식한 참가자 명단 보기");d.ontoggle=()=>{if(!d.open)return;d.querySelector(".roster")?.remove();const div=el("div");div.className="roster";div.append(table(["행","이름","출신고교","학과","신캠 조"],people.map(p=>[p.row,p.name,p.school||"미기재",p.major,p.camp])));d.append(div);};edits.append(d);
  const confirm=el("label");confirm.className="confirm";const check=el("input");check.type="checkbox";check.id="confirmed";check.onchange=()=>{$("run").disabled=!check.checked||busy;};
  confirm.append(check,document.createTextNode(" 행사 운영을 위해 명단과 제외·확인 사항을 검토했습니다."));edits.append(confirm);
  $("run").disabled=true;$("back").disabled=false;stage("review");message("");
}
$("read").onclick=async()=>{
  const year=Number($("year").value), token=++version;
  if(!Number.isInteger(year)||year<2000||year>2100){message("행사 연도를 네 자리로 입력하세요.",true);return;}
  const lf=$("leaders").files[0],pf=$("people").files[0],manual=manualMode();
  busy=true;ready();message("이 기기에서 엑셀을 읽고 있습니다.");
  try{
    const leaderRows=manual?manualData():await readBook(lf);
    const leaders=readLeaders(leaderRows,year), data=readPeople(await readBook(pf),year);
    if(token!==version)return;
    state={year,leaders,...data,groups:null};review();
  }catch(e){if(token===version)message(e instanceof InputIssue?e.message:(e.message&&/[가-힣]/.test(e.message)?e.message:"파일을 읽을 수 없습니다. 첫 시트와 .xlsx 형식을 확인하세요."),true);}
  finally{if(token===version){busy=false;ready();}}
};
$("back").onclick=()=>{state=null;clearResult();stage("setup");ready();message("");};
function results(rows){
  const {leaders,people}=state, byRow=new Map(people.map(p=>[p.row,p]));
  state.groups=rows.map(group=>group.map(r=>byRow.get(r)).sort((a,b)=>a.name.localeCompare(b.name,"ko")));
  const oldest=Math.min(...leaders.map(g=>g.year)),exceptions=people.filter(p=>p.year<oldest).length;
  $("summary").textContent=people.length+"명 · "+leaders.length+"개 조 편성 완료. 필수 조건을 모두 확인했습니다."+(exceptions?" 최고령 조장 조 배정 "+exceptions+"명.":"");
  const cams=[...new Set(people.map(p=>p.camp))].sort((a,b)=>+a-+b);
  $("stats").replaceChildren(table(["새터 조","인원","남","여",...cams.map(c=>"신캠 "+c+"조")],state.groups.map((g,i)=>[leaders[i].id+"조",g.length,g.filter(p=>p.sex==="남").length,g.filter(p=>p.sex==="여").length,...cams.map(c=>g.filter(p=>p.camp===c).length)])));
  $("scope").replaceChildren();
  leaders.forEach((g,i)=>$("scope").append(new Option(g.id+"조 조장용",String(i))));
  $("scope").append(new Option("운영진용 전체 명단 (주민등록번호 포함)","all"));
  $("scope").onchange();stage("result");message("");
}
$("run").onclick=()=>{
  if(!state||busy||!$("confirmed")?.checked)return;
  if(state.people.some(p=>!p.name)){message("빈 이름을 입력한 뒤 편성하세요.",true);return;}
  const keys=new Map(),token=name=>{const k=nameKey(name);if(!keys.has(k))keys.set(k,keys.size);return keys.get(k);};
  const leaders=state.leaders.map(g=>({id:g.id,year:g.year,keys:g.names.map(token)}));
  const people=state.people.map(({row,name,sex,year,camp,major})=>({row,key:token(name),sex,year,camp,major}));
  const failed=s=>{stopWorker();for(const input of $("review").querySelectorAll("input"))input.disabled=false;$("run").disabled=!$("confirmed").checked;$("back").disabled=false;message(s,true);};
  try{
    workerURL=URL.createObjectURL(new Blob(["'use strict';\n"+solve.toString()+"\nonmessage=e=>{try{postMessage({rows:solve(e.data,n=>{if(n===1||n%10===0)postMessage({attempt:n})})});}catch(e){postMessage({error:e.message});}};"],{type:"text/javascript"}));
    worker=new Worker(workerURL);busy=true;for(const input of $("review").querySelectorAll("input"))input.disabled=true;$("run").disabled=true;$("back").disabled=true;
    worker.onmessage=e=>{if(e.data.rows){stopWorker();results(e.data.rows);}else if(e.data.error)failed(e.data.error);else message("필수 조건을 확인하며 편성 중입니다. 시도 "+e.data.attempt+"회");};
    worker.onerror=()=>failed("브라우저에서 편성 작업을 실행할 수 없습니다.");worker.postMessage({leaders,people});
  }catch{failed("브라우저에서 편성 작업을 시작할 수 없습니다.");}
};
$("again").onclick=()=>{state.groups=null;review();};
$("scope").onchange=()=>{$("exportHint").textContent=$("scope").value==="all"?"전체 시트에 보험용 주민등록번호와 전화번호를 포함합니다. 조별 공유에는 조장용 파일을 사용하세요.":"선택한 조의 조장·조원과 연락처를 담습니다. 주민등록번호는 포함하지 않습니다.";};
function groupSheet(i){
  const l=state.leaders[i];
  return {name:l.id+"조",widths:[12,18,23,12,21,20,20],rows:[["구분","이름","학과","성별","전화번호","담당 문행위원","위원 전화번호"],...l.names.map(n=>["조장",n]),...state.groups[i].map((p,j)=>[j+1,p.name,p.major,p.sex+"자",p.phone])]};
}
$("download").onclick=()=>{
  if(!state?.groups)return;
  const all=$("scope").value==="all",i=Number($("scope").value);
  const sheets=all?[{name:"전체",widths:[10,16,24,22,10,12,20,23],rows:[["새터 조","이름","출신고교","학과","성별","신캠 조","전화번호","주민번호"],...state.groups.flatMap((g,i)=>g.map(p=>[state.leaders[i].id,p.name,p.school||"미기재",p.major,p.sex+"자",p.camp,p.phone,p.id]))]},...state.groups.map((_,i)=>groupSheet(i))]:[groupSheet(i)];
  const url=URL.createObjectURL(makeWorkbook(sheets));urls.add(url);
  const a=el("a");a.href=url;a.download=state.year+"새터_"+(all?"운영진전체":state.leaders[i].id+"조")+".xlsx";document.body.append(a);a.click();a.remove();
  setTimeout(()=>{URL.revokeObjectURL(url);urls.delete(url);},2000);message("선택한 명단의 다운로드를 요청했습니다.");
};
reset();

$("openOffline").onclick = () => $("offline").showModal();
$("closeOffline").onclick = () => $("offline").close();

if(location.protocol==="file:"){
  $("offlineLink").hidden=true;
  $("cliLink").hidden=true;
  document.querySelector(".offline p").textContent="내려받은 오프라인 화면을 사용 중입니다. 터미널판 압축 파일은 배포 페이지에서 받을 수 있습니다.";
}
