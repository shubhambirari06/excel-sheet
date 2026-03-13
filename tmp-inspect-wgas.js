const ExcelJS = require("exceljs");

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("public/8300 Meadowbrook Ln_WGCPPS1550333380_Final ECM Calculations.xlsx");
  const ws = wb.getWorksheet("Input Sheet -W'Gas");
  if (!ws) {
    console.log("sheet missing");
    return;
  }

  const addresses = [
    "A1","V1","A2","E2","I2","L2","O2","Q2","S2","A3","E3","G3","H3","L3","N3",
    "B5","C5","D5","E5","F5","G5","E10","F10","G10","H10","E11","F11","G11","H11",
    "C15","D15","E15","F15","K15"
  ];

  for (const a of addresses) {
    const c = ws.getCell(a);
    const v = c.value;
    const printable = v && typeof v === "object" && "formula" in v
      ? `FORMULA:${v.formula}`
      : JSON.stringify(v);
    console.log(`${a} => ${printable} | align=${JSON.stringify(c.alignment || {})} | numFmt=${c.numFmt || ""}`);
  }

  console.log("MERGES:");
  for (const m of ws.model.merges || []) {
    console.log(m);
  }
})();
