// Worker receives only anonymous row ids and the attributes needed for allocation.
function solve({leaders, people}, report=()=>{}) {
  const n=leaders.length, male=people.filter(p=>p.sex==="남").length, female=people.length-male;
  const oldest=Math.min(...leaders.map(g=>g.year)), names=new Map(), camps=new Map();
  const bump=(map,key,d)=>map.set(key,(map.get(key)||0)+d);
  const shuffle=a=>{for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}return a;};
  for(const p of people){bump(names,p.key,1);bump(camps,p.camp,1);}
  const eligible=new Map();
  for(const p of people) {
    const ids=leaders.flatMap((g,i)=>!g.keys.includes(p.key) && g.year<=Math.max(p.year,oldest)?[i]:[]);
    if(!ids.length) throw new Error(p.row+"행은 조장 이름과 나이 규칙을 만족하는 조가 없습니다.");
    eligible.set(p.row,ids);
  }
  for(const [key,count] of names) {
    if(count>leaders.filter(g=>!g.keys.includes(key)).length) throw new Error("같은 이름의 참가자 수가 배정 가능한 조 수를 초과합니다. 이름과 조장 명단을 확인하세요.");
  }
  const deadline=performance.now()+14000;
  const leaderById=new Map(leaders.map((g,i)=>[g.id,i]));
  const neighbors=leaders.map(g=>[leaderById.get(g.id-1),leaderById.get(g.id+1)].filter(i=>i!==undefined));
  for(let attempt=1;performance.now()<deadline;attempt++) {
    report(attempt);
    const order=shuffle(leaders.map((_,i)=>i));
    // Opposite ends give disjoint extras when possible, exactly the necessary overlap otherwise.
    const limits=leaders.map(()=>({남:Math.floor(male/n),여:Math.floor(female/n)}));
    for(let i=0;i<male%n;i++) limits[order[i]].남++;
    for(let i=0;i<female%n;i++) limits[order[n-1-i]].여++;
    const groups=leaders.map(()=>({list:new Set(),names:new Set(),sex:{남:0,여:0},camps:new Map(),majors:new Map(),years:new Map()}));
    const extra=new Map();
    function update(i,p,d) {
      const g=groups[i], before=g.camps.get(p.camp)||0, floor=Math.floor(camps.get(p.camp)/n);
      if(d>0){g.list.add(p);g.names.add(p.key);}else{g.list.delete(p);g.names.delete(p.key);}
      g.sex[p.sex]+=d;bump(g.camps,p.camp,d);bump(g.majors,p.major,d);bump(g.years,p.year,d);
      bump(extra,p.camp,Number(before+d>floor)-Number(before>floor));
    }
    function fits(i,p) {
      const g=groups[i], size=g.camps.get(p.camp)||0, total=camps.get(p.camp);
      return g.sex[p.sex]<limits[i][p.sex] && !g.names.has(p.key) &&
        (size<Math.floor(total/n) || (size===Math.floor(total/n) && (extra.get(p.camp)||0)<total%n));
    }
    function score(i,p) {
      const g=groups[i];
      const adjacent=neighbors[i].filter(j=>groups[j].names.has(p.key)||leaders[j].keys.includes(p.key)).length;
      return adjacent*240+(g.camps.get(p.camp)||0)*90+g.list.size*4+(g.majors.get(p.major)||0)*8+(g.years.get(p.year)||0)*5+Math.random()*12;
    }
    function place(p,depth,blocked,budget) {
      if(--budget.left<0 || performance.now()>=deadline) return false;
      const ids=eligible.get(p.row).filter(i=>!blocked.has(i));
      let best=-1, bestCost=Infinity;
      for(const i of ids) if(fits(i,p)){const cost=score(i,p);if(cost<bestCost){best=i;bestCost=cost;}}
      if(best>=0){update(best,p,1);return true;}
      if(depth===0) return false;
      for(const i of shuffle(ids)) {
        blocked.add(i);
        for(const q of shuffle([...groups[i].list])) {
          if(budget.left<=0 || performance.now()>=deadline) break;
          update(i,q,-1);
          if(fits(i,p)) {
            update(i,p,1);
            if(place(q,depth-1,blocked,budget)) return true;
            update(i,p,-1);
          }
          update(i,q,1);
        }
        blocked.delete(i);
      }
      return false;
    }
    const sequence=shuffle([...people]).sort((a,b)=>eligible.get(a.row).length-eligible.get(b.row).length || names.get(b.key)-names.get(a.key));
    if(sequence.every(p=>place(p,3,new Set(),{left:2000}))) {
      const rows=groups.map(g=>[...g.list]);
      const spread=xs=>Math.max(...xs)-Math.min(...xs);
      const flat=rows.flat();
      const ok=flat.length===people.length && new Set(flat.map(p=>p.row)).size===people.length &&
        spread(rows.map(r=>r.length))<=1 &&
        ["남","여"].every(s=>spread(rows.map(r=>r.filter(p=>p.sex===s).length))<=1) &&
        [...camps.keys()].every(c=>spread(rows.map(r=>r.filter(p=>p.camp===c).length))<=1) &&
        rows.every((r,i)=>new Set(r.map(p=>p.key)).size===r.length && r.every(p=>!leaders[i].keys.includes(p.key)&&leaders[i].year<=Math.max(oldest,p.year)));
      if(!ok) throw new Error("배정 결과 검증에 실패했습니다.");
      return rows.map(r=>r.map(p=>p.row));
    }
  }
  throw new Error("시간 안에 조건을 만족하는 배정을 찾지 못했습니다. 다시 편성하거나 이름·조장 나이·신캠 조 분포를 확인하세요.");
}
