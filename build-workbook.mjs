import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const out=process.cwd();
const outputDir=out+'/outputs/compensation_bi';
await fs.mkdir(outputDir,{recursive:true});
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
const wb=Workbook.create();
const summary=wb.worksheets.add('Compensation');
const data=wb.worksheets.add('Synthetic data');
for(const s of [summary,data]){s.showGridLines=false;s.getRange('A1:N800').format.font={name:'Arial',size:10,color:'#172B4D'};s.getRange('A1:N800').format.rowHeight=21;}
summary.tabColor='#18334D';
summary.getRange('A2').values=[['Workforce compensation analysis']];summary.getRange('A2').format.font={name:'Arial',size:16,bold:true,color:'#18334D'};
summary.getRange('A3').values=[['Synthetic portfolio case. No real employee records.']];
summary.getRange('A5:B6').values=[['Snapshot date',new Date(Date.UTC(2025,11,31))],['Population','All 747 synthetic records']];
summary.getRange('B5').setNumberFormat('mmm d, yyyy');
summary.getRange('A8:D8').values=[['Records','Active at snapshot','Mean annual salary','Total annual salary rates']];
summary.getRange('A9:D9').formulas=[['=COUNTA(\'Synthetic data\'!A6:A752)','=COUNTIF(\'Synthetic data\'!L6:L752,"Active")','=AVERAGE(\'Synthetic data\'!D6:D752)','=SUM(\'Synthetic data\'!D6:D752)']];
summary.getRange('C9:D9').setNumberFormat('$#,##0');
summary.getRange('A11').values=[['Salary rates include former employees; the total is not payroll expenditure.']];
summary.getRange('A13:F13').values=[['Occupation','Records','Women mean USD','Men mean USD','Difference %','Mean salary USD']];
for(let i=0;i<6;i++){const r=14+i;summary.getRange(`A${r}`).values=[[jobs[i]]];summary.getRange(`B${r}:F${r}`).formulas=[[
`=COUNTIF('Synthetic data'!$C$6:$C$752,A${r})`,
`=IF(COUNTIFS('Synthetic data'!$C$6:$C$752,A${r},'Synthetic data'!$B$6:$B$752,"Women")=0,"not available",AVERAGEIFS('Synthetic data'!$D$6:$D$752,'Synthetic data'!$C$6:$C$752,A${r},'Synthetic data'!$B$6:$B$752,"Women"))`,
`=IF(COUNTIFS('Synthetic data'!$C$6:$C$752,A${r},'Synthetic data'!$B$6:$B$752,"Men")=0,"not available",AVERAGEIFS('Synthetic data'!$D$6:$D$752,'Synthetic data'!$C$6:$C$752,A${r},'Synthetic data'!$B$6:$B$752,"Men"))`,
`=IF(AND(ISNUMBER(C${r}),ISNUMBER(D${r}),D${r}<>0),C${r}/D${r}-1,"not available")`,
`=AVERAGEIF('Synthetic data'!$C$6:$C$752,A${r},'Synthetic data'!$D$6:$D$752)`]];}
summary.getRange('C14:D19').setNumberFormat('$#,##0');summary.getRange('F14:F19').setNumberFormat('$#,##0');summary.getRange('E14:E19').setNumberFormat('0.0%');
summary.getRange('A22:D22').values=[['Annual salary band','Women','Men','Total']];
const bands=[['Below $50,000',0,50000],['$50,000–$74,999',50000,75000],['$75,000–$99,999',75000,100000],['$100,000 or more',100000,1000000000]];
for(let i=0;i<4;i++){const r=23+i;const [label,lo,hi]=bands[i];summary.getRange(`A${r}`).values=[[label]];summary.getRange(`B${r}:D${r}`).formulas=[['Women','Men'].map(g=>`=COUNTIFS('Synthetic data'!$D$6:$D$752,">=${lo}",'Synthetic data'!$D$6:$D$752,"<${hi}",'Synthetic data'!$B$6:$B$752,"${g}")`).concat(`=SUM(B${r}:C${r})`)];}
summary.getRange('A29').values=[['Interpretation']];summary.getRange('A30').values=[['Difference = women’s mean / men’s mean − 1 within occupation.']];summary.getRange('A31').values=[['These descriptive differences do not establish equal pay or discrimination.']];summary.getRange('A32').values=[['The generator assigns salaries without gender as an input. Differences arise by sampling.']];
summary.getRange('A35:B35').values=[['Reconciliation','Expected: 0']];
summary.getRange('A36:A39').values=[['Occupation counts − records'],['Salary bands − records'],['Missing salary cells'],['Invalid salary values (≤0)']];
summary.getRange('B36:B39').formulas=[['=SUM(B14:B19)-A9'],['=SUM(D23:D26)-A9'],["=COUNTBLANK('Synthetic data'!D6:D752)"],["=COUNTIF('Synthetic data'!D6:D752,\"<=0\")"]];
data.getRange('A2').values=[['Synthetic workforce records']];data.getRange('A2').format.font={name:'Arial',size:16,bold:true};
data.getRange('A3').values=[['Generated with seed 3100. Original coursework records are not included.']];
data.getRange('A5:N5').values=[['Record ID','Gender group','Occupation','Annual salary USD','Hire date','Termination date','Knowledge','Communication','Teamwork','Leadership','Quality','Status at snapshot','Tenure years','Mean evaluation']];
data.getRange('A6:K752').values=rows;
const formulas=rows.map((_,i)=>{let r=i+6;return [`=IF(E${r}>Compensation!$B$5,"Not yet hired",IF(AND(F${r}<>"",F${r}<=Compensation!$B$5),"Former","Active"))`,`=MAX(0,(IF(AND(F${r}<>"",F${r}<Compensation!$B$5),F${r},Compensation!$B$5)-E${r})/365.25)`,`=ROUND(AVERAGE(G${r}:K${r}),1)`];});
data.getRange('L6:N752').formulas=formulas;
data.getRange('D6:D752').setNumberFormat('$#,##0');data.getRange('E6:F752').setNumberFormat('yyyy-mm-dd');data.getRange('M6:N752').setNumberFormat('0.0');
data.tables.add('A5:N752',true,'SyntheticWorkforce');data.freezePanes.freezeRows(5);data.freezePanes.freezeColumns(1);
data.getRange('A5:N5').format.font={name:'Arial',size:10,bold:true,color:'#FFFFFF'};
data.getRange('G6:K752').format.horizontalAlignment='center';
data.getRange('A1:N752').format.columnWidth=16;data.getRange('C1:C752').format.columnWidth=21;data.getRange('L1:L752').format.columnWidth=22;data.getRange('A5:N5').format.wrapText=true;data.getRange('A5:N5').format.rowHeight=34;
summary.getRange('A1:A39').format.columnWidth=33;summary.getRange('B1:F39').format.columnWidth=22;summary.getRange('G1:L39').format.columnWidth=14;
for(const r of ['A8:D8','A13:F13','A22:D22','A35:B35']){summary.getRange(r).format.fill='#18334D';summary.getRange(r).format.font={name:'Arial',size:10,bold:true,color:'#FFFFFF'};summary.getRange(r).format.wrapText=true;summary.getRange(r).format.rowHeight=32;}
summary.getRange('A9:D9').format.font={name:'Arial',size:14,bold:true,color:'#18334D'};
summary.getRange('B36:B39').conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:'#FEE2E2',font:{color:'#991B1B',bold:true}}});
const chart=summary.charts.add('bar',[summary.getRange('A13:A19'),summary.getRange('C13:C19'),summary.getRange('D13:D19')]);
chart.title='Mean annual salary by occupation';chart.setPosition('H5','P21');chart.titleTextStyle.typeface='Arial';chart.titleTextStyle.fontSize=14;chart.legend={position:'bottom'};chart.yAxis={numberFormatCode:'$#,##0',numberFormatSourceLinked:false};chart.series.items[0].fill='#236E91';chart.series.items[1].fill='#C68539';
wb.recalculate();
const values=summary.getRange('A9:D9').values[0];
const total=rows.reduce((s,r)=>s+r[3],0);assert.equal(values[0],747);assert.equal(values[1],rows.filter(r=>r[5]===null).length);assert.ok(Math.abs(values[2]-total/747)<1e-7);assert.equal(values[3],total);
assert.deepEqual(summary.getRange('B36:B39').values,[[0],[0],[0],[0]]);
for(let i=0;i<6;i++){const rr=rows.filter(r=>r[2]===jobs[i]);assert.equal(summary.getRange(`B${14+i}`).values[0][0],rr.length);assert.ok(Math.abs(summary.getRange(`F${14+i}`).values[0][0]-rr.reduce((s,r)=>s+r[3],0)/rr.length)<1e-7);}
// Recalculation test: a controlled salary edit changes the linked mean, then is restored.
const before=values[2], old=data.getRange('D6').values[0][0];data.getRange('D6').values=[[old+7470]];wb.recalculate();assert.ok(Math.abs(summary.getRange('C9').values[0][0]-before-10)<1e-7);data.getRange('D6').values=[[old]];wb.recalculate();
data.getRange('D6').values=[[null]];wb.recalculate();assert.equal(summary.getRange('B38').values[0][0],1);data.getRange('D6').values=[[0]];wb.recalculate();assert.equal(summary.getRange('B39').values[0][0],1);data.getRange('D6').values=[[old]];
// Employment date boundaries are explicit and do not depend on the computer clock.
const oldHire=data.getRange('E6').values[0][0], oldLeave=data.getRange('F6').values[0][0];
data.getRange('E6').values=[[new Date(Date.UTC(2026,0,1))]];wb.recalculate();assert.equal(data.getRange('L6').values[0][0],'Not yet hired');assert.equal(data.getRange('M6').values[0][0],0);
data.getRange('E6').values=[[oldHire]];data.getRange('F6').values=[[new Date(Date.UTC(2025,11,31))]];wb.recalculate();assert.equal(data.getRange('L6').values[0][0],'Former');data.getRange('F6').values=[[oldLeave]];wb.recalculate();
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:20}});
await fs.writeFile(out+'/validation.json',JSON.stringify({records:747,seed:3100,active:values[1],salaryTotal:total,meanSalary:values[2],occupationCounts:jobs.map(j=>[j,rows.filter(r=>r[2]===j).length]),tests:'Passed independent counts, sums, means, reconciliations and salary-edit recalculation',errorScan:errors.ndjson},null,2));
for(const [sheet,range,name] of [['Compensation','A1:P39','dashboard'],['Synthetic data','A1:N18','data-preview']]){const img=await wb.render({sheetName:sheet,range,scale:1.5,format:'png'});await fs.writeFile(out+`/${name}.png`,new Uint8Array(await img.arrayBuffer()));}
await(await SpreadsheetFile.exportXlsx(wb)).save(outputDir+'/workforce-compensation.xlsx');
await fs.writeFile(out+'/synthetic-workforce.csv',data.getRange('A5:K752').values.map(r=>r.map(v=>v instanceof Date?v.toISOString().slice(0,10):v??'').join(',')).join('\n')+'\n');
console.log(JSON.stringify({outputDir,values,total,errorScan:errors.ndjson}));
