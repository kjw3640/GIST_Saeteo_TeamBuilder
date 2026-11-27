class InputIssue extends Error {}
function issue(s){throw new InputIssue(s);}
const nameKey=name=>Array.from(name.trim()).slice(name.trim().length>1?1:0).join("");
const cleanSchool=s=>s.replace(/고등학교/g,"고").replace(/과고/g,"과학고").replace(/여자고/g,"여고").replace(/교원대학교부설고/g,"교대부고").replace(/대학교사범대학부설고/g,"사대부고");
const cleanDigits=s=>s.replace(/\.0$/,"").replace(/[\s-]/g,"");
function birthInfo(id){
  const m=/^(\d{2})(\d{2})(\d{2})([1-8])\d{6}$/.exec(id);
  if(!m) return null;
  const year=Number(m[1])+("3478".includes(m[4])?2000:1900);
  const date=new Date(Date.UTC(year,+m[2]-1,+m[3]));
  return date.getUTCFullYear()===year && date.getUTCMonth()===+m[2]-1 && date.getUTCDate()===+m[3]?year:null;
}
function readLeaders(rows,eventYear) {
  const [headers,...rest]=rows, h=headers.cells;
  const names=[], years=[];
  h.forEach((v,i)=>{if(/이름|성명/.test(v))names.push(i);if(/생년|year/i.test(v))years.push(i);});
  const group=h.findIndex(v=>v&&/조/.test(v)&&/번호/.test(v));
  const slots=group<0||names.length!==2||years.length!==2?[0,1,2,3,4]:[group,names[0],years[0],names[1],years[1]];
  const ids=new Set(),leaders=[];
  for(const {cells,number} of rest) {
    const [g,a,y,b,z]=slots.map(i=>cells[i]||"");
    if(!g&&!a&&!b&&!y&&!z) continue;
    const id=Number(g.replace(/\s*조$/,""));
    const parseYear=v=>/^\d{4}$/.test(v)?+v:birthInfo(cleanDigits(v))||(/^\d{4}[-/.]/.test(v)?Number(v.slice(0,4)):0);
    const yrs=[parseYear(y),parseYear(z)];
    if(!Number.isSafeInteger(id)||id<1||id>9999||ids.has(id)||!a||!b||yrs.some(v=>v<1900||v>eventYear))
      issue("조장 명단 "+number+"행의 조 번호, 두 조장 이름과 네 자리 생년을 확인하세요.");
    ids.add(id);leaders.push({id,names:[a,b],keys:[nameKey(a),nameKey(b)],year:Math.min(...yrs)});
  }
  if(!leaders.length||leaders.length>120)issue("조장 명단에는 1~120개 조가 필요합니다.");
  return leaders.sort((a,b)=>a.id-b.id);
}
function readPeople(rows,eventYear) {
  const [header,...rest]=rows;
  function find(pattern,label,optional=false){
    const cols=[];header.cells.forEach((v,i)=>{if(pattern.test(v)&&!/동의|보호자|비상연락/.test(v))cols.push(i);});
    if(cols.length>1 || (!cols.length&&!optional))issue(label+" 열을 하나로 식별할 수 없습니다. 첫 행의 제목을 확인하세요.");
    return cols[0];
  }
  const name=find(/이름|성명/,"이름"),school=header.cells.findIndex((v,i)=>i!==name&&v&&/고교|고등학교|출신고/.test(v));
  const cam=find(/신캠|신입생\s*캠프/,"신캠 조"),major=find(/학과|학부/,"학과"),gender=find(/성별/,"성별");
  const phone=find(/전화|휴대폰|휴대전화|연락처/,"전화번호"), idCol=find(/주민/,"주민등록번호");
  const people=[],notices=[], phones=new Set(), ids=new Set();
  for(const {cells,number} of [...rest].reverse()){
    if(![name,cam,major,gender,phone,idCol].some(i=>cells[i]))continue;
    const get=i=>(cells[i]||"").trim(), rawName=get(name);
    const split=school>=0 ? null : rawName.match(/^(.+?(?:고등학교|학교|고))\s*[+,_/\-]?\s*([가-힣]{2,4})$/)||rawName.match(/^(.+?)[\s+,_/\-]+([가-힣]{2,4})$/);
    const real=split?split[2]:rawName, hs=cleanSchool(get(school)||(split?split[1]:""));
    let p=cleanDigits(get(phone)), id=cleanDigits(get(idCol));
    if(/^1\d{9}$/.test(p)) p="0"+p;
    if(/^\d{12}$/.test(id))id="0"+id;
    if(!/^01\d{8,9}$/.test(p))issue("새내기 명단 "+number+"행의 휴대전화 번호를 확인하세요.");
    const year=birthInfo(id);
    if(!year||year>eventYear)issue("새내기 명단 "+number+"행의 주민등록번호 자릿수와 생년월일을 확인하세요.");
    if(phones.has(p)||ids.has(id)){notices.push(number+"행 · 중복 제출로 제외 (아래쪽 응답 유지)");continue;}
    phones.add(p);ids.add(id);
    const checksum=(11-[2,3,4,5,6,7,8,9,2,3,4,5].reduce((a,w,i)=>a+w*Number(id[i]),0)%11)%10===Number(id[12]);
    if(year>=eventYear-18&&checksum){notices.push(number+"행 · 기존 행사 연령 기준으로 제외");continue;}
    if(!checksum)notices.push(number+"행 · "+(year>=eventYear-18?"연령 및 ":"")+"보험용 번호 원본 확인 필요 (명단에 포함)");
    const rawSex=get(gender).replace(/[\s?\u200b\ufeff]/g,"").toLowerCase();
    const sex=["남","남자","남성","m","male","man"].includes(rawSex)?"남":["여","여자","여성","f","female","woman"].includes(rawSex)?"여":"";
    const camp=get(cam).replace(/\s*조$/,""),dept=get(major);
    if(!real||!sex||!dept||!/^\d+$/.test(camp))issue("새내기 명단 "+number+"행의 이름·성별·학과·신캠 조를 확인하세요.");
    people.push({row:number,name:real,school:hs,needsEdit:!hs||!/^([가-힣]{2,4})$/.test(real),major:dept.includes("도전탐색과정")?"도전탐색과정":dept,sex,camp:String(Number(camp)),year,phone:p.replace(/^(\d{3})(\d{3,4})(\d{4})$/,"$1-$2-$3"),id:id.slice(0,6)+"-"+id.slice(6)});
  }
  if(!people.length)issue("배정할 참가자가 없습니다. 행사 연도와 명단을 확인하세요.");
  return {people:people.reverse(),notices:notices.reverse()};
}
