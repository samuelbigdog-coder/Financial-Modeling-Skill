import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { styleMiniModel } from "./style-mini-model.mjs";
import { fileURLToPath, pathToFileURL } from "node:url";

const local = path.join(os.homedir(), ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs");
const { Workbook, SpreadsheetFile } = await import(process.env.FIN_MODEL_ARTIFACT_TOOL || pathToFileURL(local).href);
const folder = path.dirname(fileURLToPath(import.meta.url));
const book = Workbook.create();
const tabs = [
  "Inputs", "Mini Model", "Cases", "M&A Schedule", "Output -->", "Equity Bridge",
  "Fan of Outcomes", "CF Walk", "Returns Scenarios", "LR Performance", "Segment Margins",
  "Exec Summary", "Sources--->", "P&L Detailed", "DS Backlog", "M&A Targets", "EM Backlog",
  "Hist BS Detailed", "SOFR", "Contr. Margin", "LR Data", "Margin Explored", "Comps",
  "Deal Comps", "Capex", "Checks",
];
const sheets = Object.fromEntries(tabs.map(name => [name, book.worksheets.add(name)]));
const navy = "#002060", blue = "#0000FF", green = "#00B050", grey = "#595959";
const yellow = "#FFFFCC", light = "#DCE6F1", black = "#000000";
const dollar = '#,##0.0;(#,##0.0);"-"', pct = '0.0%;(0.0%);"-"', mult = '0.0"x";(0.0"x");"-"';
const history = ["E","F","G"], future = ["H","I","J","K","L"], periods = [...history, ...future];
const years = [2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030];

function r(name, address) { return sheets[name].getRange(address); }
function value(name, address, data) { r(name, address).values = [[data]]; }
function formula(name, address, data) {
  const cell = r(name, address);
  cell.formulas = [[data]];
  if (/^='[^']+'!\$?[A-Z]+\$?\d+$/.test(data)) cell.format.font.color = green;
}
function link(name, cell) { return "'" + name + "'!" + cell; }
function put(name, address, data) {
  if (typeof data === "string" && data.startsWith("=")) formula(name, address, data);
  else value(name, address, data);
}
function block(name, row, title) {
  const band = r(name, "C" + row + ":L" + row);
  band.format.fill = navy;
  band.format.font.color = "#FFFFFF";
  band.format.font.bold = true;
  value(name, "C" + row, title);
}
function row(name, n, label, data, cols = periods, format = dollar) {
  value(name, "C" + n, label);
  if (format) r(name, cols[0] + n + ":" + cols.at(-1) + n).setNumberFormat(format);
  data.forEach((entry, i) => put(name, cols[i] + n, entry));
}
function input(name, address, data, fmt = dollar) {
  value(name, address, data);
  const cell = r(name, address);
  cell.format.font.color = blue;
  cell.format.fill = yellow;
  if (fmt) cell.setNumberFormat(fmt);
}
function total(name, ...rows) {
  for (const n of rows) {
    const range = r(name, "C" + n + ":L" + n);
    range.format.font.bold = true;
    range.format.borders = { top: { style: "thin", color: navy } };
  }
}
function annual(name, cols = periods) {
  for (const col of cols) {
    const i = periods.indexOf(col);
    value(name, col + "6", String(years[i]) + (i < 3 ? "A" : "E"));
    const cell = r(name, col + "6");
    cell.format.fill = i < 3 ? "#E7E6E6" : light;
    cell.format.font.bold = true;
  }
}
function lines(name, entries, cols = periods) {
  for (const [n, label, fn, format] of entries) row(name, n, label, cols.map(fn), cols, format);
}
for (const name of tabs) {
  const sheet = sheets[name];
  sheet.showGridLines = false;
  const div = name === "Output -->" || name === "Sources--->";
  const group = tabs.indexOf(name);
  sheet.tabColor = div ? navy : group < 4 ? "#4472C4" : group < 12 ? "#70AD47" : "#A5A5A5";
  if (div) continue;
  r(name, "A1:L110").format.font.name = "Arial";
  r(name, "A1:L110").format.font.size = 10;
  r(name, "A1:L110").format.font.color = black;
  r(name, "A:B").format.columnWidth = 3;
  r(name, "C:C").format.columnWidth = 39;
  r(name, "D:D").format.columnWidth = 3;
  r(name, "E:L").format.columnWidth = 13;
  r(name, "A1:L110").format.rowHeight = 13;
  sheet.freezePanes.freezeRows(6);
  block(name, 3, "NORTHSTAR MOTION SYSTEMS  |  " + name.toUpperCase());
  value(name, "C4", "Fictional teaching example | USD millions unless noted");
  r(name, "C4").format.font.color = grey;
  if (name !== "Inputs") formula(name, "L5", "=" + link("Inputs", "E8"));
}

// Controls and evidence. Every value in the workbook is invented.
const I = "Inputs";
r(I, "E:E").format.columnWidth = 17;
r(I, "F:F").format.columnWidth = 4;
r(I, "G:G").format.columnWidth = 55;
r(I, "H:H").format.columnWidth = 29;
block(I, 7, "Scenario controls and transaction assumptions");
const inputs = [
  [8,"Operating case","Base","Select the organic growth and gross margin profile.",null],
  [9,"Acquisition plan","Selective","Select annual completed bolt-ons.",null],
  [10,"Entry EV / EBITDA",8.25,"Illustrative platform price.",mult],
  [11,"Exit EV / EBITDA",8.5,"Illustrative hold-end price.",mult],
  [12,"Maximum entry leverage",3.25,"Debt sizing ceiling as a multiple of LTM EBITDA.",mult],
  [13,"Maximum entry loan-to-value",.48,"Debt sizing ceiling as a share of entry EV.",pct],
  [14,"Minimum cash",3,"Cash reserve at close and through the forecast.",dollar],
  [15,"Transaction fee / entry EV",.018,"Expensed at close in this simplified purchase.",pct],
  [16,"Annual term amortization",.07,"Share of original principal amortized each year.",pct],
  [17,"Term loan spread",.045,"Margin above annual illustrative benchmark.",pct],
  [18,"Revolver spread",.055,"Margin above benchmark on beginning draws.",pct],
  [19,"Maximum revolver commitment",16,"Absolute borrowing commitment.",dollar],
  [20,"Maximum funded leverage",3.75,"Limits additional revolver draws.",mult],
  [21,"Cash tax rate",.26,"Current cash tax only. Loss carryforwards excluded.",pct],
  [22,"DSO",48,"Days of total revenue held in receivables.",'0'],
  [23,"DIO",47,"Days of total direct cost held in inventory.",'0'],
  [24,"DPO",37,"Days of total direct cost held in payables.",'0'],
  [25,"Annual fixed cost inflation",.035,"Wage and occupancy assumption.",pct],
  [26,"Variable operating cost / revenue",.09,"Applies only to organic revenue.",pct],
  [27,"Maintenance capex / total revenue",.024,"Replacement investment.",pct],
  [28,"Growth capex / organic revenue",.012,"Growth investment in platform capacity.",pct],
  [29,"D&A / opening PP&E",.15,"Simplified depreciation with no tax/book split.",pct],
  [30,"Acquired EBITDA margin",.23,"Stand-alone margin after operating expense.",pct],
  [31,"Add-on EV / EBITDA",6.1,"Flat purchase multiple for all add-ons.",mult],
  [32,"Average acquired EBITDA",1.4,"Same annual run rate for each completed deal.",dollar],
  [33,"Add-on net tangible allocation",.20,"Purchase price share assigned to PP&E.",pct],
  [34,"Multiple sensitivity step",.5,"Symmetric increments around selected multiples.",mult],
  [35,"Reconciliation tolerance",.000001,"Absolute tolerance in model units.",'0.000000'],
  [36,"Days convention",365,"Annual working-capital day count.",'0'],
];
for (const [n, label, val, why, fmt] of inputs) {
  value(I, "C" + n, label);
  input(I, "E" + n, val, fmt);
  value(I, "G" + n, why);
  value(I, "H" + n, "Synthetic design assumption");
}
r(I, "E8").dataValidation = { rule: { type:"list", values:["Base","Upside","Downside","Management"] } };
r(I, "E9").dataValidation = { rule: { type:"list", values:["Selective","Accelerated","Limited","None"] } };
value(I, "C38", "Blue/yellow = editable input | green = direct internal link | black = calculation");
value(I, "C39", "Entry is at 2025 year-end. Bolt-ons contribute a full year of EBITDA.");



const P = "P&L Detailed";
annual(P, history);
block(P, 8, "Synthetic historical operating data");
const pnl = [
  [9,"Motion components revenue",[46,50,55]],[10,"Service kits revenue",[26,29,31]],
  [11,"Automation projects revenue",[18,21,24]],[13,"Motion components direct cost",[29,31,34]],
  [14,"Service kits direct cost",[14,15,16]],[15,"Automation projects direct cost",[13,15,17]],
  [17,"Fixed operating expense",[9,9.5,10.1]],[18,"Variable operating expense",[8,9,9.9]],
  [21,"Nonrecurring owner expense add-back",[.3,.4,.5]],[22,"Nonrecurring legal adjustment",[.2,.1,.2]],
];
for (const [n, label, vals] of pnl) {
  value(P,"C"+n,label);
  vals.forEach((v,i) => input(P, history[i]+n, v));
}
lines(P, [
  [12,"Total revenue",c=>"=SUM("+c+"9:"+c+"11)"],
  [16,"Gross profit",c=>"="+c+"12-SUM("+c+"13:"+c+"15)"],
  [19,"Reported EBITDA",c=>"="+c+"16-SUM("+c+"17:"+c+"18)"],
  [23,"Adjusted EBITDA",c=>"="+c+"19+SUM("+c+"21:"+c+"22)"],
  [24,"Adjusted EBITDA margin",c=>"="+c+"23/"+c+"12",pct],
], history);
total(P, 12, 16, 19, 23);

const B = "Hist BS Detailed";
annual(B, history);
block(B,8,"Synthetic pre-close balance sheets");
const historicalBalances = [
  [9,"Cash",[3.8,4.2,4.7]], [10,"Receivables",[12,13,14.5]], [11,"Inventory",[7,7.6,8.5]],
  [12,"Net PP&E",[12,13,14]], [13,"Goodwill",[2,2,2]], [15,"Payables",[5.6,6.1,6.7]],
  [16,"Debt",[9,8,7]], [17,"Historical equity",[22.2,25.7,30]],
];
for (const [n,label,vals] of historicalBalances) {
  value(B,"C"+n,label);
  vals.forEach((v,i)=>input(B,history[i]+n,v));
}
lines(B,[
  [14,"Total assets",c=>"=SUM("+c+"9:"+c+"13)"],
  [18,"Liabilities and equity",c=>"=SUM("+c+"15:"+c+"17)"],
  [19,"Historical balance residual",c=>"="+c+"14-"+c+"18"],
],history);
total(B,14,18,19);

const C = "Cases";
annual(C);
block(C,8,"Organic case bank");
const drivers = [
  [9,"Motion components growth",9,[.065,.105,-.04,.09]],
  [16,"Service kits growth",10,[.075,.11,.015,.095]],
  [23,"Automation projects growth",11,[.09,.14,-.06,.12]],
  [30,"Motion components gross margin",13,[.39,.42,.32,.41]],
  [37,"Service kits gross margin",14,[.49,.52,.43,.51]],
  [44,"Automation projects gross margin",15,[.30,.34,.22,.33]],
];
for (const [n,label,src,assumptions] of drivers) {
  value(C,"C"+n,label+" | active");
  r(C,"E"+n+":L"+n).setNumberFormat(pct);
  history.forEach((col,i)=>{
    if (src >= 13) formula(C,col+n,"=1-"+link(P,col+src)+"/"+link(P,col+(src-4)));
    else if(i>0) formula(C,col+n,"="+link(P,col+src)+"/"+link(P,history[i-1]+src)+"-1");
  });
  const cases = ["Base","Upside","Downside","Management"];
  for (let i=0;i<cases.length;i++) {
    value(C,"C"+(n+i+1),cases[i]);
    future.forEach((c,j)=>input(C,c+(n+i+1),assumptions[i]+(i===2 && j>0 ? .008*j : 0),pct));
  }
  future.forEach(c=>formula(C,c+n,"=INDEX("+c+(n+1)+":"+c+(n+4)+",MATCH("+link(I,"$E$8")+",$C$"+(n+1)+":$C$"+(n+4)+",0))"));
  total(C,n);
}

const A = "M&A Schedule";
annual(A,future);
block(A,8,"Acquisition cadence and purchase price allocation");
const plans = [["Selective",[0,1,0,1,0]],["Accelerated",[1,1,1,1,1]],["Limited",[0,0,1,0,0]],["None",[0,0,0,0,0]]];
for (const [i,[label,vals]] of plans.entries()) {
  value(A,"C"+(10+i),label);
  vals.forEach((v,j)=>input(A,future[j]+(10+i),v,'0'));
}
lines(A,[
  [9,"Active completed deals",c=>"=INDEX("+c+"10:"+c+"13,MATCH("+link(I,"$E$9")+", $C$10:$C$13,0))",'0'],
  [15,"New run-rate acquired EBITDA",c=>"="+c+"9*"+link(I,"$E$32")],
  [16,"Cumulative acquired EBITDA",c=>"=SUM($H15:"+c+"15)"],
  [17,"Acquired business revenue",c=>"="+c+"16/"+link(I,"$E$30")],
  [18,"Cash acquisition consideration",c=>"="+c+"15*"+link(I,"$E$31")],
  [19,"PP&E acquired",c=>"="+c+"18*"+link(I,"$E$33")],
  [20,"Goodwill acquired",c=>"="+c+"18-"+c+"19"],
],future);
total(A,9,16,18,20);

const rate = "SOFR";
annual(rate,future);
block(rate,8,"Synthetic annual benchmark curve");
value(rate,"C9","Illustrative SOFR assumption");
[.041,.037,.035,.035,.036].forEach((v,j)=>input(rate,future[j]+9,v,pct));
value(rate,"C11","Fabricated curve for instruction.");

const cap = "Capex";
annual(cap);
block(cap,8,"Historical and forecast capital expenditure");
for (const [n,label,vals] of [[9,"Maintenance capital spending",[2,2.2,2.5]],[10,"Growth capital spending",[1.1,1.3,1.6]]]) {
  value(cap,"C"+n,label);
  vals.forEach((v,i)=>input(cap,history[i]+n,v));
}
const M = "Mini Model";
lines(cap,[
  [9,"Maintenance capital spending",c=>"="+link(M,c+"28")+"*"+link(I,"$E$27")],
  [10,"Growth capital spending",c=>"="+link(M,c+"25")+"*"+link(I,"$E$28")],
],future);
lines(cap,[[11,"Total capital spending",c=>"=SUM("+c+"9:"+c+"10)"]]);
total(cap,11);

// Close and forecast. Cash before financing is computed before revolver and equity funding.
annual(M);
block(M,8,"Transaction capitalization");
const close = [
  [9,"LTM adjusted EBITDA","="+link(P,"G23")],
  [10,"Entry enterprise value","=G9*"+link(I,"$E$10")],
  [11,"Minimum cash funded","="+link(I,"$E$14")],
  [12,"Transaction fees expensed at close","=G10*"+link(I,"$E$15")],
  [13,"Total uses","=SUM(G10:G12)"],
  [14,"New term loan","=MIN(G9*"+link(I,"$E$12")+",G10*"+link(I,"$E$13")+")"],
  [15,"Sponsor equity at entry","=G13-G14"],
  [16,"Total sources","=SUM(G14:G15)"],
  [17,"Acquired operating NWC","="+link(B,"G10")+"+"+link(B,"G11")+"-"+link(B,"G15")],
  [18,"Acquired net PP&E","="+link(B,"G12")],
  [19,"Entry goodwill","=G10-G17-G18"],
];
for(const [n,label,expr] of close) row(M,n,label,[expr],["G"]);
block(M,21,"Income statement");
lines(M,[
  [22,"Motion components revenue",c=>"="+link(P,c+"9")],[23,"Service kits revenue",c=>"="+link(P,c+"10")],
  [24,"Automation projects revenue",c=>"="+link(P,c+"11")],[25,"Organic revenue",c=>"=SUM("+c+"22:"+c+"24)"],
  [28,"Total revenue",c=>"="+c+"25"],[29,"Motion components direct cost",c=>"="+link(P,c+"13")],
  [30,"Service kits direct cost",c=>"="+link(P,c+"14")],[31,"Automation projects direct cost",c=>"="+link(P,c+"15")],
  [33,"Total direct cost",c=>"=SUM("+c+"29:"+c+"31)"],[34,"Gross profit",c=>"="+c+"28-"+c+"33"],
  [35,"Fixed operating cost",c=>"="+link(P,c+"17")],[36,"Variable operating cost",c=>"="+link(P,c+"18")],
  [38,"Adjusted EBITDA",c=>"="+link(P,c+"23")],[39,"Adjusted EBITDA margin",c=>"="+c+"38/"+c+"28",pct],
],history);
const prev = c=>periods[periods.indexOf(c)-1];
const ass = n=>link(I,"$E$"+n);
lines(M,[
  [22,"Motion components revenue",c=>"="+prev(c)+"22*(1+"+link(C,c+"9")+")"],
  [23,"Service kits revenue",c=>"="+prev(c)+"23*(1+"+link(C,c+"16")+")"],
  [24,"Automation projects revenue",c=>"="+prev(c)+"24*(1+"+link(C,c+"23")+")"],
  [25,"Organic revenue",c=>"=SUM("+c+"22:"+c+"24)"],
  [26,"Acquired revenue",c=>"="+link(A,c+"17")],
  [28,"Total revenue",c=>"=SUM("+c+"25:"+c+"26)"],
  [29,"Motion components direct cost",c=>"="+c+"22*(1-"+link(C,c+"30")+")"],
  [30,"Service kits direct cost",c=>"="+c+"23*(1-"+link(C,c+"37")+")"],
  [31,"Automation projects direct cost",c=>"="+c+"24*(1-"+link(C,c+"44")+")"],
  [32,"Acquired direct and operating cost",c=>"="+c+"26-"+link(A,c+"16")],
  [33,"Total direct cost",c=>"=SUM("+c+"29:"+c+"32)"],
  [34,"Gross profit before organic opex",c=>"="+c+"28-"+c+"33"],
  [35,"Fixed operating cost",c=>"="+prev(c)+"35*(1+"+ass(25)+")"],
  [36,"Variable operating cost",c=>"="+c+"25*"+ass(26)],
  [37,"Organic EBITDA",c=>"="+c+"25-SUM("+c+"29:"+c+"31)-SUM("+c+"35:"+c+"36)"],
  [38,"Adjusted EBITDA",c=>"="+c+"34-SUM("+c+"35:"+c+"36)"],
  [39,"Adjusted EBITDA margin",c=>"="+c+"38/"+c+"28",pct],
  [40,"Depreciation",c=>"="+prev(c)+"54*"+ass(29)],
  [41,"EBIT",c=>"="+c+"38-"+c+"40"],
  [42,"Cash interest",c=>"="+c+"82"],
  [43,"Pre-tax income",c=>"="+c+"41-"+c+"42"],
  [44,"Cash taxes",c=>"=MAX(0,"+c+"43*"+ass(21)+")"],
  [45,"Net income",c=>"="+c+"43-"+c+"44"],
],future);
block(M,50,"Balance sheet | 2025A column shown post-close");
const postClose = [
  [51,"Cash","=G11"],[52,"Receivables","="+link(B,"G10")],[53,"Inventory","="+link(B,"G11")],
  [54,"Net PP&E","=G18"],[55,"Goodwill","=G19"],[56,"Total assets","=SUM(G51:G55)"],
  [57,"Payables","="+link(B,"G15")],[58,"Term debt","=G14"],[59,"Revolver","=G14-G14"],
  [60,"Sponsor book equity","=G15-G12"],[61,"Liabilities and equity","=SUM(G57:G60)"],
  [62,"Operating net working capital","=SUM(G52:G53)-G57"],
];
for (const [n,label,expr] of postClose) row(M,n,label,[expr],["G"]);
lines(M,[
  [51,"Cash",c=>"="+c+"72"],[52,"Receivables",c=>"="+c+"28*"+ass(22)+"/"+ass(36)],
  [53,"Inventory",c=>"="+c+"33*"+ass(23)+"/"+ass(36)],
  [54,"Net PP&E",c=>"="+prev(c)+"54+"+c+"67-"+c+"40+"+link(A,c+"19")],
  [55,"Goodwill",c=>"="+prev(c)+"55+"+link(A,c+"20")],
  [56,"Total assets",c=>"=SUM("+c+"51:"+c+"55)"],
  [57,"Payables",c=>"="+c+"33*"+ass(24)+"/"+ass(36)],
  [58,"Term debt",c=>"="+c+"77"],[59,"Revolver",c=>"="+c+"81"],
  [60,"Sponsor book equity",c=>"="+prev(c)+"60+"+c+"45+"+c+"88"],
  [61,"Liabilities and equity",c=>"=SUM("+c+"57:"+c+"60)"],
  [62,"Operating net working capital",c=>"=SUM("+c+"52:"+c+"53)-"+c+"57"],
],future);
block(M,64,"Cash flow");
lines(M,[
  [65,"Cash flow before NWC",c=>"="+c+"45+"+c+"40"],
  [66,"Investment in NWC",c=>"="+c+"62-"+prev(c)+"62"],
  [67,"Capex",c=>"="+link(cap,c+"11")],
  [68,"Acquisition consideration",c=>"="+link(A,c+"18")],
  [69,"Opening cash",c=>"="+prev(c)+"51"],
  [70,"Cash before discretionary financing",c=>"="+c+"69+"+c+"65-SUM("+c+"66:"+c+"68)-"+c+"76"],
  [71,"Revolver draws less repayments plus equity",c=>"="+c+"79-"+c+"80+"+c+"88"],
  [72,"Ending cash",c=>"=SUM("+c+"70:"+c+"71)"],
],future);
block(M,74,"Debt and liquidity");
lines(M,[
  [75,"Opening term debt",c=>"="+prev(c)+"58"],
  [76,"Scheduled term amortization",c=>"=MIN("+c+"75,$G$14*"+ass(16)+")"],
  [77,"Ending term debt",c=>"="+c+"75-"+c+"76"],
  [78,"Opening revolver",c=>"="+prev(c)+"59"],
  [79,"Revolver draw",c=>"=MIN("+c+"85,"+c+"86)"],
  [80,"Revolver sweep",c=>"=MIN("+c+"78,MAX(0,"+c+"70-"+ass(14)+"))"],
  [81,"Ending revolver",c=>"=SUM("+c+"78:"+c+"79)-"+c+"80"],
  [82,"Interest on term loan and opening revolver",c=>"=AVERAGE("+c+"75,"+c+"77)*"+c+"83+"+c+"78*"+c+"84"],
  [83,"Term loan cash rate",c=>"="+link(rate,c+"9")+"+"+ass(17),pct],
  [84,"Revolver cash rate",c=>"="+link(rate,c+"9")+"+"+ass(18),pct],
  [85,"Borrowing capacity for incremental draw",c=>"=MAX(0,MIN("+ass(19)+"-"+c+"78,"+ass(20)+"*"+c+"38-"+c+"77-"+c+"78))"],
  [86,"Cash shortfall before discretionary financing",c=>"=MAX(0,"+ass(14)+"-"+c+"70)"],
  [87,"Ending funded debt / EBITDA",c=>"=IF("+c+"38>0,("+c+"77+"+c+"81)/"+c+"38,"+"0)",mult],
  [88,"Sponsor equity top-up",c=>"="+c+"86-"+c+"79"],
  [89,"Discretionary debt / equity funding residual",c=>"="+c+"86-"+c+"79-"+c+"88"],
],future);
block(M,91,"Sponsor returns | hypothetical exit by year");
row(M,92,"Annual period",[0],["G"],'0');
row(M,97,"Sponsor equity contributions",["=G15"],["G"]);
row(M,98,"Cash flows to sponsor",["=-G97"],["G"]);
lines(M,[
  [92,"Annual period",c=>"="+prev(c)+"92+1",'0'],
  [93,"Exit EV / EBITDA",c=>"="+ass(11),mult],
  [94,"Exit enterprise value",c=>"="+c+"38*"+c+"93"],
  [95,"Net debt at exit",c=>"=SUM("+c+"58:"+c+"59)-"+c+"51"],
  [96,"Exit equity proceeds",c=>"=MAX(0,"+c+"94-"+c+"95)"],
  [97,"Sponsor equity contributions",c=>"="+c+"88"],
  [98,"Cash flows to sponsor",c=>c==="L"?"="+c+"96-"+c+"97":"=-"+c+"97"],
  [99,"Exit MOIC",c=>"="+c+"96/SUM($G97:"+c+"97)",mult],
],future);
row(M,100,"2030 sponsor annual IRR",["=IF(L98>0,IRR(G98:L98),IF(L96=0,-1,\"N/M\"))"],["L"],pct);
total(M,13,16,28,34,38,45,56,61,72,77,81,88,96,99,100);


// Operating evidence and market context. All records are intentionally fictional.
const D = "DS Backlog", E = "EM Backlog", T = "M&A Targets", CM = "Contr. Margin";
function table(name, n, headers, records) {
  block(name,n-1,"Synthetic source records");
  const cols = ["C","E","F","G","H","I","J","K","L"];
  headers.forEach((h,i)=>{
    value(name,cols[i]+n,h);
    r(name,cols[i]+n).format.font.bold=true;
    r(name,cols[i]+n).format.fill=light;
  });
  records.forEach((rec,j)=>rec.forEach((v,i)=>{
    const adr = cols[i]+(n+j+1);
    if(typeof v === "number") input(name,adr,v);
    else put(name,adr,v);
  }));
}
table(D,9,["End market","Awarded","Pipeline","Win rate","Weighted"],[
  ["Industrial service",7.8,4.5,.55,"=E10+F10*G10"],
  ["Fleet maintenance",5.3,2.4,.6,"=E11+F11*G11"],
  ["Utility customers",3.7,1.6,.45,"=E12+F12*G12"],
]);
r(D,"G10:G12").setNumberFormat(pct);
row(D,13,"Weighted service backlog",["=SUM(H10:H12)"],["H"]);
row(D,15,"Coverage of first forecast service revenue",["=H13/"+link(M,"H23")],["H"],pct);
total(D,13);
table(E,9,["Project","Value","Booked?","Probability","Weighted"],[
  ["Project Sequoia",6,"Yes",1,"=E10*G10"],
  ["Project Flint",3.4,"Yes",1,"=E11*G11"],
  ["Project Quartz",4.6,"No",.5,"=E12*G12"],
  ["Project Willow",3.2,"No",.25,"=E13*G13"],
]);
r(E,"G10:G13").setNumberFormat(pct);
row(E,15,"Weighted project backlog",["=SUM(H10:H13)"],["H"]);
row(E,16,"Coverage of first forecast project revenue",["=H15/"+link(M,"H24")],["H"],pct);
total(E,15);
table(T,9,["Illustrative target","Revenue","EBITDA","Offer EV","EV/EBITDA"],[
  ["Alto Drive Works",6.4,1.4,8.54,"=G10/F10"],
  ["Mariner Controls",5.8,1.3,7.2,"=G11/F11"],
  ["Juniper Automation",7.2,1.5,9.15,"=G12/F12"],
]);
r(T,"H10:H12").setNumberFormat(mult);
value(T,"C15","Pipeline evidence only. The M&A Schedule owns the standardized case assumptions.");
table(CM,9,["Representative job","Revenue","Direct cost","Contribution","Margin"],[
  ["Motion platform",6.7,4.05,"=E10-F10","=G10/E10"],
  ["Service kit batch",4.2,2.0,"=E11-F11","=G11/E11"],
  ["Automation cell",3.5,2.5,"=E12-F12","=G12/E12"],
]);
r(CM,"H10:H12").setNumberFormat(pct);
row(CM,13,"Weighted contribution margin",["=SUM(G10:G12)/SUM(E10:E12)"],["H"],pct);
total(CM,13);
const LR = "LR Data";
annual(LR,history);
block(LR,8,"Fictional operating metrics");
for(const [n,label,vals] of [[9,"Orders booked",[94,106,117]],[10,"Active customers",[76,82,86]],[11,"Direct headcount",[145,157,168]]]){
  value(LR,"C"+n,label);
  vals.forEach((v,i)=>input(LR,history[i]+n,v,n>9?'0':dollar));
}
const MX="Margin Explored";
annual(MX,history);
block(MX,8,"Reported-to-adjusted margin and cost mix");
lines(MX,[
  [9,"Reported EBITDA margin",c=>"="+link(P,c+"19")+"/"+link(P,c+"12"),pct],
  [10,"Owner expense adjustment",c=>"="+link(P,c+"21")+"/"+link(P,c+"12"),pct],
  [11,"Legal expense adjustment",c=>"="+link(P,c+"22")+"/"+link(P,c+"12"),pct],
  [12,"Adjusted EBITDA margin",c=>"=SUM("+c+"9:"+c+"11)",pct],
  [14,"Fixed cost / revenue",c=>"="+link(P,c+"17")+"/"+link(P,c+"12"),pct],
  [15,"Variable cost / revenue",c=>"="+link(P,c+"18")+"/"+link(P,c+"12"),pct],
],history);
total(MX,12);

const PC="Comps", DC="Deal Comps";
table(PC,9,["Fictional trading peer","EV","Revenue","EBITDA","EV/revenue","EV/EBITDA"],[
  ["Aster Industrial",340,174,38,"=E10/F10","=E10/G10"],
  ["Rook Motion",470,215,51,"=E11/F11","=E11/G11"],
  ["Cinder Controls",238,142,25,"=E12/F12","=E12/G12"],
  ["Summit Drive",525,290,64,"=E13/F13","=E13/G13"],
]);
r(PC,"H10:I14").setNumberFormat(mult);
row(PC,14,"Median trading multiples",["=MEDIAN(H10:H13)","=MEDIAN(I10:I13)"],["H","I"],mult);
total(PC,14);
table(DC,9,["Fictional control transaction","EV","Revenue","EBITDA","EV/EBITDA","Include?"],[
  ["Atlas Components",126,102,15,"=E10/G10","Yes"],
  ["Pioneer Service",80,64,11,"=E11/G11","Yes"],
  ["Bridgeport Cells",54,38,7,"=E12/G12","Yes"],
  ["Meridian Megadeal",870,530,87,"=E13/G13","No"],
]);
r(DC,"H10:H13").setNumberFormat(mult);
row(DC,15,"Included control median",["=MEDIAN(J10:J13)"],["H"],mult);
row(DC,16,"Included sample size",['=COUNTIF(I10:I13,"Yes")'],["H"],'0');
value(DC,"J9","Eligible");
[10,11,12,13].forEach(n=>formula(DC,"J"+n,'=IF(I'+n+'="Yes",H'+n+',"")'));
r(DC,"J10:J13").setNumberFormat(mult);
value(DC,"C18","Meridian is excluded from the small-company precedent set.");
total(DC,15);

// Reporting tabs own their interpretations rather than repeating whole schedules.
const EB="Equity Bridge", CF="CF Walk", SP="Segment Margins", LP="LR Performance", XS="Exec Summary";
block(EB,8,"Entry enterprise-to-equity and hold-end value creation");
r(EB,"C:C").format.columnWidth=46;
const bridge = [
 [9,"Entry sponsor equity","="+link(M,"G15")],
 [10,"Fee leakage at entry","=-"+link(M,"G12")],
 [11,"Organic EBITDA growth at entry multiple","=("+link(M,"L37")+"-"+link(M,"G9")+")*"+ass(10)],
 [12,"Acquired EBITDA at entry multiple","="+link(A,"L16")+"*"+ass(10)],
 [13,"Multiple change on total EBITDA","="+link(M,"L38")+"*("+ass(11)+"-"+ass(10)+")"],
 [14,"Change in net debt","="+link(M,"G58")+"-"+link(M,"G51")+"-"+link(M,"L95")],
 [15,"Unfloored exit equity","=SUM(E9:E14)"],
 [16,"Limited liability floor","=MAX(0,-E15)"],
 [17,"Exit sponsor equity","=SUM(E15:E16)"],
 [19,"Additional sponsor investment","=SUM("+link(M,"H97:L97")+")"],
 [20,"Sponsor value creation","=E17-E9-E19"],
];
for(const [n,label,f] of bridge) row(EB,n,label,[f],["E"]);
total(EB,15,17,20);
annual(CF,future);
block(CF,8,"Cash generation and financing walk");
lines(CF,[
  [9,"Beginning cash",c=>"="+link(M,c+"69")],
  [10,"EBITDA",c=>"="+link(M,c+"38")],
  [11,"Cash interest",c=>"=-"+link(M,c+"42")],
  [12,"Cash taxes",c=>"=-"+link(M,c+"44")],
  [13,"Working capital",c=>"=-"+link(M,c+"66")],
  [14,"Maintenance capex",c=>"=-"+link(cap,c+"9")],
  [15,"Growth capex",c=>"=-"+link(cap,c+"10")],
  [16,"Levered free cash flow",c=>"=SUM("+c+"10:"+c+"15)"],
  [17,"Acquisition consideration",c=>"=-"+link(M,c+"68")],
  [18,"Term amortization",c=>"=-"+link(M,c+"76")],
  [19,"Revolver activity",c=>"="+link(M,c+"79")+"-"+link(M,c+"80")],
  [20,"Sponsor equity",c=>"="+link(M,c+"88")],
  [21,"Ending cash from walk",c=>"="+c+"9+SUM("+c+"16:"+c+"20)"],
],future);
total(CF,16,21);
annual(SP);
block(SP,8,"Segment gross margins and consolidated mix");
lines(SP,[
  [9,"Motion components gross margin",c=>"=1-"+link(M,c+"29")+"/"+link(M,c+"22"),pct],
  [10,"Service kits gross margin",c=>"=1-"+link(M,c+"30")+"/"+link(M,c+"23"),pct],
  [11,"Automation projects gross margin",c=>"=1-"+link(M,c+"31")+"/"+link(M,c+"24"),pct],
  [12,"Motion components revenue mix",c=>"="+link(M,c+"22")+"/"+link(M,c+"28"),pct],
  [13,"Service kits revenue mix",c=>"="+link(M,c+"23")+"/"+link(M,c+"28"),pct],
  [14,"Automation projects revenue mix",c=>"="+link(M,c+"24")+"/"+link(M,c+"28"),pct],
  [15,"Acquired revenue mix",c=>"="+link(M,c+"26")+"/"+link(M,c+"28"),pct],
  [16,"Mix total",c=>"=SUM("+c+"12:"+c+"15)",pct],
  [17,"Consolidated adjusted EBITDA margin",c=>"="+link(M,c+"39"),pct],
]);
total(SP,16,17);
annual(LP,history);
block(LP,8,"Operating efficiency and reported performance");
lines(LP,[
  [9,"Book-to-bill",c=>"="+link(LR,c+"9")+"/"+link(P,c+"12"),mult],
  [10,"Revenue per active customer",c=>"="+link(P,c+"12")+"/"+link(LR,c+"10")],
  [11,"Revenue per direct FTE ($000)",c=>"="+link(P,c+"12")+"*1000/"+link(LR,c+"11")],
  [12,"Reported EBITDA margin",c=>"="+link(MX,c+"9"),pct],
  [13,"Adjusted EBITDA margin",c=>"="+link(P,c+"24"),pct],
],history);
row(LP,15,"Historical revenue CAGR",["=("+link(P,"G12")+"/"+link(P,"E12")+")^(1/2)-1"],["G"],pct);
total(LP,13,15);
annual(XS);
block(XS,8,"Operating and investment summary");
lines(XS,[
  [9,"Motion components revenue",c=>"="+link(M,c+"22")],
  [10,"Service kits revenue",c=>"="+link(M,c+"23")],
  [11,"Automation projects revenue",c=>"="+link(M,c+"24")],
  [12,"Total revenue",c=>"="+link(M,c+"28")],
  [13,"Adjusted EBITDA",c=>"="+link(M,c+"38")],
  [14,"Adjusted EBITDA margin",c=>"="+link(M,c+"39"),pct],
  [15,"Maintenance capital spending",c=>"="+link(cap,c+"9")],
]);
lines(XS,[
  [17,"Levered free cash flow",c=>"="+link(CF,c+"16")],
  [18,"Ending cash",c=>"="+link(M,c+"51")],
  [19,"Funded leverage",c=>"="+link(M,c+"87"),mult],
  [20,"Sponsor MOIC if exited",c=>"="+link(M,c+"99"),mult],
],future);
row(XS,22,"Sponsor IRR at final exit",["="+link(M,"L100")],["L"],pct);
row(XS,23,"Sponsor value created",["="+link(EB,"E20")],["L"]);
row(XS,25,"Entry uses",["="+link(M,"G13")],["G"]);
row(XS,26,"Entry debt",["="+link(M,"G14")],["G"]);
row(XS,27,"Entry sponsor equity",["="+link(M,"G15")],["G"]);
total(XS,12,13,17,22,23,25,27);

// Sensitivity revalues each exit multiple against the same live operating and funding model.
// Changing entry pricing also changes financing and cash interest; show a full-engine entry
// price sensitivity in the snapshot table, avoiding an entry-price-only denominator shortcut.
const RS="Returns Scenarios";
annual(RS,future);
block(RS,8,"Live exit year and multiple sensitivity");
lines(RS,[
  [9,"EBITDA at exit",c=>"="+link(M,c+"38")],
  [10,"Net debt at exit",c=>"="+link(M,c+"95")],
  [11,"Cumulative sponsor contributions",c=>"=SUM("+link(M,"$G$97:"+c+"97")+")"],
  [12,"Base MOIC at exit",c=>"="+link(M,c+"99"),mult],
],future);
const horizonCash = ["G",...future];
for(const [i,c] of future.entries()){
  const n=29+i;
  value(RS,"C"+n,"Exit in "+years[i+3]);
  for(const [j,d] of horizonCash.entries()){
    if(j>i+1) continue;
    const f=j===i+1 ? "="+link(M,d+"96")+"-"+link(M,d+"97") : "=-"+link(M,d+"97");
    formula(RS,d+n,f);
  }
  formula(RS,c+"13",'=IF('+c+n+'>0,IRR(G'+n+':'+c+n+'),IF('+link(M,c+"96")+'=0,-1,"N/M"))');
}
value(RS,"C13","Base annual IRR at exit");
r(RS,"H13:L13").setNumberFormat(pct);
block(RS,15,"Live exit MOIC by exit multiple and year");
for(const [i,offset] of [-2,-1,0,1,2].entries()){
  const n=17+i;
  row(RS,n,"Exit multiple",["="+ass(11)+"+"+offset+"*"+ass(34)],["E"],mult);
  future.forEach(c=>formula(RS,c+n,"=MAX(0,"+c+"$9*$E"+n+"-"+c+"$10)/"+c+"$11"));
  r(RS,"H"+n+":L"+n).setNumberFormat(mult);
}
total(RS,12,13,19);
block(RS,27,"Dated at equal annual intervals | cash flows supporting each hypothetical exit");
value(RS,"C36","Valuation scope");
value(RS,"C37","This buyout example values exit EBITDA with market multiples and tests sponsor equity cash flows.");
value(RS,"C38","Standalone intrinsic valuation would require a separate unlevered cash flow and terminal-value case.");
value(RS,"C39","IRR is N/M when the last net cash flow is nonpositive despite positive exit proceeds.");

const FO="Fan of Outcomes";
block(FO,8,"Full-engine release scenario captures | rerun builder to refresh");
value(FO,"C25","Example release 1.0 | 2026-11-03 | Each capture re-runs the active model with the indicated controls.");
value(FO,"C26","Captures are fixed release outputs. Input edits update the active model and live exit sensitivity.");
r(FO,"C:C").format.columnWidth=41;
for (const [c,h] of [["C","Capture"],["E","Operations"],["F","M&A"],["G","Entry x"],["H","2030 EBITDA"],["I","2030 cash"],["J","Equity added"],["K","2030 MOIC"],["L","2030 IRR"]]) value(FO,c+"9",h);
r(FO,"C9:L9").format.fill=light;
r(FO,"C9:L9").format.font.bold=true;

const CH="Checks";
annual(CH);
block(CH,8,"Independent accounting and sensitivity diagnostics");
const ch = [
  [9,"Pre-close historical balance",c=>"="+link(B,c+"14")+"-"+link(B,c+"18")],
  [10,"Purchase sources and uses",c=>"="+link(M,c+"16")+"-"+link(M,c+"13")],
  [11,"Balance sheet residual",c=>"="+link(M,c+"56")+"-"+link(M,c+"61")],
  [12,"Cash flow and cash walk",c=>"="+link(CF,c+"21")+"-"+link(M,c+"51")],
  [13,"Acquisition consideration allocation",c=>"="+link(A,c+"18")+"-"+link(A,c+"19")+"-"+link(A,c+"20")],
  [14,"Organic plus acquired EBITDA",c=>"="+link(M,c+"37")+"+"+link(A,c+"16")+"-"+link(M,c+"38")],
  [15,"Cash floor shortfall",c=>"=MAX(0,"+ass(14)+"-"+link(M,c+"51")+")"],
  [16,"Revolver commitment excess",c=>"=MAX(0,"+link(M,c+"59")+"-"+ass(19)+")"],
  [17,"Operating NWC roll-forward",c=>"="+link(M,c+"62")+"-"+link(M,prev(c)+"62")+"-"+link(M,c+"66")],
  [18,"Term debt roll-forward",c=>"="+link(M,prev(c)+"58")+"-"+link(M,c+"76")+"-"+link(M,c+"58")],
  [19,"Exit sensitivity center vs live MOIC",c=>"="+link(RS,c+"19")+"-"+link(M,c+"99")],
  [20,"Sensitivity monotonicity violations",c=>"=MAX(0,"+link(RS,c+"17")+"-"+link(RS,c+"18")+")+MAX(0,"+link(RS,c+"18")+"-"+link(RS,c+"19")+")+MAX(0,"+link(RS,c+"19")+"-"+link(RS,c+"20")+")+MAX(0,"+link(RS,c+"20")+"-"+link(RS,c+"21")+")"],
  [21,"Revenue mix residual",c=>"="+link(SP,c+"16")+"-1"],
  [22,"Funding requirement vs draw and equity",c=>"="+link(M,c+"86")+"-"+link(M,c+"79")+"-"+link(M,c+"88")],
  [23,"Equity bridge vs exit sponsor proceeds",c=>"="+link(EB,"E17")+"-"+link(M,"L96")],
  [24,"Current cash draw and sweep overlap",c=>"=MIN("+link(M,c+"79")+","+link(M,c+"80")+")"],
];
const formatCheck = '0.000000;(0.000000);"-"';
for(const [n,label,f] of ch) {
  const cols=n===9?history:n===10?["G"]:n===11?["G",...future]:n===23?["L"]:future;
  row(CH,n,label,cols.map(f),cols,formatCheck);
}
lines(CH,[[26,"Period status",c=>'=IF(MAX('+c+'9:'+c+'24)<= '+ass(35)+',IF(MIN('+c+'9:'+c+'24)>=-'+ass(35)+',"PASS","FAIL"),"FAIL")',null]]);
r(CH,"C:C").format.columnWidth=46;
r(CH,"E9:L24").conditionalFormats.add("cellIs",{operator:"greaterThan",formula:0.000001,format:{fill:"#FFC7CE"}});
r(CH,"E9:L24").conditionalFormats.add("cellIs",{operator:"lessThan",formula:-0.000001,format:{fill:"#FFC7CE"}});
total(CH,26);

function scalar(name, address) { return r(name,address).values[0][0]; }
function runCase(op, ma, entry) {
  value(I,"E8",op); value(I,"E9",ma); value(I,"E10",entry);
  book.recalculate();

  if (periods.some(c=>scalar(CH,c+"26")!=="PASS")) {
    throw new Error("Reconciliation failed for "+op+" / "+ma+" / "+entry);
  }
  if (op==="Downside" && ma==="Limited") {
    for (const fundingRow of [79,80,88]) {
      if (!future.some(c=>scalar(M,c+fundingRow)>0)) {
        throw new Error("Downside funding path was not exercised at row "+fundingRow);
      }
    }
  }
  return [scalar(M,"L38"),scalar(M,"L51"),scalar(EB,"E19"),scalar(M,"L99"),scalar(M,"L100")];
}
const captures = [
  ["Base", "Base","Selective",8.25],
  ["Upside", "Upside","Accelerated",8.25],
  ["Downside", "Downside","Limited",8.25],
  ["Management", "Management","Selective",8.25],
  ["No deals", "Base","None",8.25],
  ["Lower entry price", "Base","Selective",7.75],
  ["Higher entry price", "Base","Selective",8.75],
];
for(const [i,[label,op,ma,entry]] of captures.entries()){
  const n=10+i;
  const captured=runCase(op,ma,entry);
  if(!captured.every(Number.isFinite)) throw new Error("Invalid capture "+label+": "+JSON.stringify(captured));
  const vals=[label,op,ma,entry,...captured];
  const dest=["C","E","F","G","H","I","J","K","L"];
  vals.forEach((v,j)=>value(FO,dest[j]+n,v));
}
for(const [c,fmt] of [["G",mult],["H",dollar],["I",dollar],["J",dollar],["K",mult],["L",pct]]) r(FO,c+"10:"+c+"16").setNumberFormat(fmt);
runCase("Base","Selective",8.25);
total(FO,10);
if (periods.some(c=>scalar(CH,c+"26")!=="PASS")) throw new Error("Release reconciliation failure");
const baseExit = scalar(M,"L99");
value(I,"E11",9);
book.recalculate();
if (!(scalar(M,"L99")>baseExit)) throw new Error("Exit pricing did not flow to the active return");
value(I,"E11",8.5);
styleMiniModel(sheets);
book.recalculate();
const file = await SpreadsheetFile.exportXlsx(book);
const output = path.join(folder,"Northstar_Motion_MiniModel.xlsx");
await file.save(output);
await fs.rm(output+".inspect.ndjson",{force:true});
if(process.env.FIN_MODEL_PREVIEWS) {
  await fs.mkdir(process.env.FIN_MODEL_PREVIEWS,{recursive:true});
  for(const [name,area] of [["Inputs","A1:H41"],["Mini Model","A1:L46"],["Exec Summary","A1:L27"],["DS Backlog","A1:L16"],["Returns Scenarios","A1:L38"]]){
    const png=await book.render({sheetName:name,range:area,scale:1,format:"png"});
    await fs.writeFile(path.join(process.env.FIN_MODEL_PREVIEWS,name+".png"),new Uint8Array(await png.arrayBuffer()));
  }
}





