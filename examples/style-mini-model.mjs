// Reusable visual layer for the synthetic example. This changes no formula or model cell.
export function styleMiniModel(sheets) {
  const navy = "#002060", medium = "#1F497D", ink = "#000000";
  const gray = "#F2F2F2", lavender = "#E6E0EC", lightBlue = "#DCE6F1";
  const white = "#FFFFFF";
  const layouts = {
    "Inputs":           { title: "B2", width: 34, years: 12 },
    "Mini Model":       { title: "D5", width: 42, years: 11 },
    "Cases":            { title: "B2", width: 34, years: 10 },
    "M&A Schedule":     { title: "B2", width: 39, years: 10 },
    "Equity Bridge":    { title: "B2", width: 44, years: 13 },
    "Fan of Outcomes":  { title: "C2", width: 30, years: 14, open: true },
    "CF Walk":          { title: "B2", width: 28, years: 11, header: medium, titleSize: 14 },
    "Returns Scenarios":{ title: "B2", width: 31, years: 12 },
    "LR Performance":   { title: "B2", width: 30, years: 9, open: true },
    "Segment Margins":  { title: "B2", width: 38, years: 12, open: true },
    "Exec Summary":     { title: "B2", width: 32, years: 11, header: "#4472C4" },
    "P&L Detailed":     { title: "C3", width: 38, years: 11, raw: "Arial", size: 8 },
    "DS Backlog":       { title: "B2", width: 33, years: 12, raw: "Century Gothic", size: 10, grid: true, header: "#00B0F0" },
    "M&A Targets":      { title: "A1", width: 33, years: 18, raw: "Century Gothic", size: 10, grid: true, header: "#595959" },
    "EM Backlog":       { title: "B2", width: 28, years: 13, raw: "Century Gothic", size: 10, grid: true, header: "#156297" },
    "Hist BS Detailed": { title: "C3", width: 38, years: 12, raw: "Arial", size: 8 },
    "SOFR":             { title: "C3", width: 34, years: 11, header: medium },
    "Contr. Margin":    { title: "B2", width: 34, years: 12, raw: "Century Gothic", size: 10, header: "#156297" },
    "LR Data":          { title: "B2", width: 30, years: 9, open: true },
    "Margin Explored":  { title: "B2", width: 32, years: 11, header: "#156297" },
    "Comps":            { title: "B1", width: 40, years: 13 },
    "Deal Comps":       { title: "B2", width: 46, years: 12, open: true },
    "Capex":            { title: "B3", width: 38, years: 10 },
    "Checks":           { title: "B2", width: 44, years: 12 },
  };
  const range = (name, address) => sheets[name].getRange(address);
  const tint = (name, address, fill) => { range(name,address).format.fill = fill; };
  const line = (name, address) => {
    range(name,address).format.borders = { bottom: { style:"thin", color:ink } };
  };
  const outputStart = Object.keys(sheets).indexOf("Output -->");
  const sourcesStart = Object.keys(sheets).indexOf("Sources--->");

  for (const [index, [name, sheet]] of Object.entries(sheets).entries()) {
    const style = layouts[name];
    sheet.tabColor = index < 2 ? medium : index < 4 ? "#C0504D"
      : !style ? white : index > outputStart && index < sourcesStart ? "#FF0000"
      : name === "Checks" ? "#808080" : "#F79646";
    sheet.showGridLines = style?.grid ?? false;
    sheet.freezePanes.unfreeze();
    if (!style) continue;
    const canvas = range(name,"A1:L110");
    canvas.format.font.name = "Calibri";
    canvas.format.font.size = 11;
    canvas.format.rowHeight = 15;
    range(name,"A:B").format.columnWidth = 2;
    range(name,"C:C").format.columnWidth = style.width;
    range(name,"D:D").format.columnWidth = 2;
    range(name,"E:L").format.columnWidth = style.years;
    const oldTitle = range(name,"C3").values[0][0];
    tint(name,"C3:L3",white);
    range(name,"C3:L3").format.font.color = ink;
    range(name,"C3:L3").format.font.bold = false;
    if (style.title !== "C3") range(name,"C3").values = [[null]];
    const title = range(name,style.title);
    title.values = [[oldTitle]];
    title.format.font.name = "Calibri";
    title.format.font.size = style.titleSize ?? 12;
    title.format.font.bold = true;
    title.format.font.color = name === "CF Walk" ? navy : ink;
    const notes = range(name,"C4");
    notes.format.font.size = 10;
    notes.format.font.italic = true;
    if (style.raw) {
      const body = range(name,"C9:L21");
      body.format.font.name = style.raw;
      body.format.font.size = style.size;
    }
    const main = range(name,"C8:L8");
    if (style.open) {
      main.format.fill = white;
      main.format.font.color = ink;
      line(name,"C8:L8");
    } else if (style.header) {
      main.format.fill = style.header;
    }
    if (name === "CF Walk") {
      tint(name,"H6:L6",medium);
      range(name,"H6:L6").format.font.color = white;
    }
  }
  range("Inputs","E:E").format.columnWidth = 14;
  range("Inputs","F:F").format.columnWidth = 2;
  range("Inputs","G:G").format.columnWidth = 76;
  range("Inputs","H:H").format.columnWidth = 36;
  tint("Inputs","G8:H36",lightBlue);
  range("Inputs","G8:H36").format.wrapText = true;
  range("Inputs","C8:H36").format.rowHeight = 30;
  range("Inputs","C8:H36").format.verticalAlignment = "center";
  range("Inputs","E10:E12").setNumberFormat('0.00"x"');
  range("Fan of Outcomes","G10:G16").setNumberFormat('0.00"x"');
  for (let n=8;n<=36;n++) {
    const cell=range("Inputs","E"+n);
    if (cell.values[0][0] === null) continue;
    const edge={style:"thin",color:"#A6A6A6"};
    cell.format.borders={top:edge,right:edge,bottom:edge,left:edge};
  }
  for (const row of [12,13,17,22,23,25,27]) tint("Exec Summary","C"+row+":L"+row,gray);
  for (const row of [16,21]) tint("CF Walk","C"+row+":L"+row,gray);
  for (const row of [9,17,20]) tint("Equity Bridge","C"+row+":E"+row,lavender);
  for (const row of [10,12,14]) tint("Comps","C"+row+":I"+row,gray);
  tint("Comps","H14:I14",lightBlue);
  for (const name of ["DS Backlog","M&A Targets","EM Backlog","Contr. Margin","Comps","Deal Comps","Fan of Outcomes"]) {
    range(name,"C9:L9").format.wrapText=true;
    range(name,"C9:L9").format.rowHeight=29;
  }
  range("Mini Model","M:M").format.columnWidth=2;
  range("Mini Model","N:N").format.columnWidth=30;
  range("Fan of Outcomes","E:F").format.columnWidth=16;
  range("Fan of Outcomes","G:L").format.columnWidth=14;
  const rateSheet = sheets.SOFR;
  range("SOFR","H9:L9").setNumberFormat("0.00%");
  const curve = range("SOFR","C14:D19");
  curve.values = [["Year","SOFR"],[null,null],[null,null],[null,null],[null,null],[null,null]];
  ["H","I","J","K","L"].forEach((column,i) => {
    range("SOFR","C"+(i+15)+":D"+(i+15)).formulas = [["="+column+"6","="+column+"9"]];
  });
  range("SOFR","D:D").format.columnWidth = 10;
  range("SOFR","D15:D19").setNumberFormat("0.00%");
  range("SOFR","C14:D14").format.fill = medium;
  range("SOFR","C14:D14").format.font.color = white;
  range("SOFR","C14:D14").format.font.bold = true;
  const chart = rateSheet.charts.add("line",curve);
  chart.title = "Illustrative SOFR curve";
  chart.titleTextStyle.typeface = "Calibri";
  chart.hasLegend = false;
  chart.xAxis = {axisType:"textAxis",textStyle:{typeface:"Calibri"}};
  chart.yAxis = {numberFormatCode:"0.00%",numberFormatSourceLinked:false,textStyle:{typeface:"Calibri"}};
  chart.setPosition("F13","L28");
}



