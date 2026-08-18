import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(.:)/, "$1")), "..");
const dataDir = path.join(root, "integration-data", "summit-ridge");
const data = JSON.parse(await fs.readFile(path.join(dataDir, "workbook-data.json"), "utf8"));
const wb = Workbook.create();

const specs = [
  ["_Meta", ["key","value"], data.meta],
  ["Projects", ["native_id","code","name","status","scenario"], data.projects],
  ["CostCodes", ["code","name"], data.cost_codes],
  ["Activities", ["project_code","cost_code","name","quantity","unit","planned_start","planned_finish","progress_method"], data.activities],
  ["Progress", ["project_code","cost_code","as_of_date","pct_complete","recorded_by"], data.progress],
  ["Budgets", ["project_code","budget_version","status","cost_code","amount","label"], data.budgets],
  ["Contracts", ["project_code","client","contract_value","payment_terms_days","claim_frequency","retention_pct","retention_release","median_days_late"], data.contracts],
  ["ContractChanges", ["project_code","seq_no","description","value","status","approved_date"], data.contract_changes],
  ["Commitments", ["project_code","subcontractor","cost_code","original_value","payment_terms_days","retention_pct"], data.commitments],
  ["CashInputs", ["key","value","note"], data.cash_inputs],
  ["Sources", ["item","source","reference","notes"], data.sources],
];

function colName(n) { let s=""; while(n){ n--; s=String.fromCharCode(65+n%26)+s; n=Math.floor(n/26); } return s; }
for (const [name, headers, rows] of specs) {
  const sh = wb.worksheets.add(name); sh.showGridLines = false;
  const matrix = [headers, ...rows];
  const end = `${colName(headers.length)}${matrix.length}`;
  sh.getRange(`A1:${end}`).values = matrix;
  const hdr=sh.getRange(`A1:${colName(headers.length)}1`);
  hdr.format.fill="#1F4E78"; hdr.format.font={bold:true,color:"#FFFFFF"}; hdr.format.rowHeight=24;
  sh.getRange(`A1:${end}`).format.font={name:"Aptos",size:10};
  hdr.format.font={name:"Aptos Display",size:10,bold:true,color:"#FFFFFF"};
  sh.freezePanes.freezeRows(1);
  sh.getRange(`A1:${end}`).format.autofitColumns();
  sh.getRange(`A1:${end}`).format.autofitRows();
  for(let c=0;c<headers.length;c++){
    const h=headers[c]; const rng=sh.getRange(`${colName(c+1)}2:${colName(c+1)}${matrix.length}`);
    if(h.includes("date")||h.includes("start")||h.includes("finish")) rng.setNumberFormat("yyyy-mm-dd");
    if(["amount","contract_value","original_value"].includes(h) || (name==="ContractChanges" && h==="value")) {
      rng.setNumberFormat('"$"#,##0.00;[Red]("$"#,##0.00);-');
      rng.format.columnWidth=16;
    }
    if(h.includes("pct")) rng.setNumberFormat("0.0");
  }
  const textCols=headers.map((h,i)=>({h,i})).filter(x=>["name","description","label","note","notes","source","reference"].includes(x.h));
  for(const x of textCols) sh.getRange(`${colName(x.i+1)}1:${colName(x.i+1)}${matrix.length}`).format.columnWidth=32;
  if(name==="_Meta") sh.getRange(`B1:B${matrix.length}`).format.columnWidth=32;
  if(name==="CashInputs") {
    sh.getRange(`B1:B${matrix.length}`).format.columnWidth=18;
    for (const r of [2,4,5]) sh.getRange(`B${r}`).setNumberFormat('"$"#,##0.00;[Red]("$"#,##0.00);-');
  }
}

await fs.mkdir(dataDir,{recursive:true});
const out=await SpreadsheetFile.exportXlsx(wb);
await out.save(path.join(dataDir,"project-input.xlsx"));
const inspect=await wb.inspect({kind:"table",range:"Projects!A1:E8",include:"values,formulas",tableMaxRows:10,tableMaxCols:8});
console.log(inspect.ndjson);
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},summary:"formula scan"});
console.log(errors.ndjson);
const previewDir=path.join(dataDir,"verification-previews");
await fs.mkdir(previewDir,{recursive:true});
for (const [name, headers, rows] of specs) {
  const lastRow=Math.min(rows.length+1,25);
  const preview=await wb.render({sheetName:name,range:`A1:${colName(headers.length)}${lastRow}`,scale:1.2});
  await fs.writeFile(path.join(previewDir,`${name}.png`),new Uint8Array(await preview.arrayBuffer()));
}
