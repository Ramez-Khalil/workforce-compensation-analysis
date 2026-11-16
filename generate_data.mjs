import fs from 'node:fs/promises';
let state=3100;
const rand=()=>((state=(Math.imul(state,1664525)+1013904223)>>>0)/4294967296);
const jobs=['Administration','Analyst','Engineering','Operations','Sales','Support'];
const bases=[48000,65000,88000,56000,62000,43000];
const rows=[];
for(let i=0;i<747;i++){
 const j=Math.floor(rand()*6), gender=rand()<0.49?'Women':'Men';
 const years=Math.floor(rand()*15), hire=Date.UTC(2010+years,Math.floor(rand()*12),1);
 const leave=rand()<0.14?Date.UTC(2025,Math.floor(rand()*12),1):null;
 const salary=Math.round((bases[j]+rand()*24000)/100)*100;
 rows.push([`SYN-${String(i+1).padStart(4,'0')}`,gender,jobs[j],salary,new Date(hire),leave?new Date(leave):null,...Array.from({length:5},()=>1+Math.floor(rand()*5))]);
}

const headers=['Record ID','Gender group','Occupation','Annual salary USD','Hire date','Termination date','Knowledge','Communication','Teamwork','Leadership','Quality'];
const csv=[headers,...rows].map(r=>r.map(v=>v instanceof Date?v.toISOString().slice(0,10):v??'').join(',')).join('\n')+'\n';
await fs.writeFile('synthetic-workforce.csv',csv);
console.log('Generated 747 synthetic records, seed 3100.');
