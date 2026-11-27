const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const source = fs.readFileSync(path.join(__dirname, "../src/solver.js"), "utf8");
const {solve} = vm.runInNewContext(source + "\n;({solve})", {performance});
const leaders = Array.from({length:20}, (_,i)=>({id:i+1, keys:[500+i], year:2000+i%4}));
const people = Array.from({length:200}, (_,i)=>({
  row:i+2, key:i%150, sex:i<110?"남":"여", year:i===0?1999:2006,
  camp:String(i%17), major:i%3
}));
const rows = solve({leaders, people});
assert.equal(new Set(rows.flat()).size,200);
for(const [i,group] of rows.entries()){
  const members=Array.from(group, row=>people[row-2]);
  assert.equal(group.length,10);
  assert.ok([5,6].includes(members.filter(p=>p.sex==="남").length));
  assert.equal(new Set(members.map(p=>p.key)).size,10);
  if(group.includes(2))assert.equal(leaders[i].year,2000);
}
for(const camp of new Set(people.map(p=>p.camp))){
  const counts=rows.map(group=>group.filter(row=>people[row-2].camp===camp).length);
  assert.ok(Math.max(...counts)-Math.min(...counts)<=1);
}
assert.throws(()=>solve({leaders:[{id:1,keys:[1],year:2000}],people:[{...people[0],key:1}]}),/없습니다/);
console.log("Passed 20-team quota, camp, age, and name checks.");
