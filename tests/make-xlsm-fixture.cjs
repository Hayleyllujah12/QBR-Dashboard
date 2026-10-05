// Builds tests/SAMPLE_Lenovo_Inventory_SYNTH.xlsm from the synthetic .xlsx sample plus a
// fake VBA project, so the linked "Save to Excel" path can be tested on a macro workbook
// (2026-10-05 .xlsm corruption fix). Usage: node tests/make-xlsm-fixture.cjs
const path = require("path"), fs = require("fs");
const LIB = [path.resolve(__dirname, "../qbr-app/libs/xlsx.full.min.js"), path.resolve(__dirname, "../libs/xlsx.full.min.js")].find(fs.existsSync);
const XLSX = require(LIB);
const src = path.resolve(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx");
const wb = XLSX.read(fs.readFileSync(src), { type: "buffer", cellStyles: true });
wb.vbaraw = Buffer.from("SYNTHETIC-VBA-PROJECT-FOR-TESTS-" + "0".repeat(256));
fs.writeFileSync(path.resolve(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsm"),
  XLSX.write(wb, { bookType: "xlsm", type: "buffer", bookVBA: true, cellStyles: true }));
console.log("wrote tests/SAMPLE_Lenovo_Inventory_SYNTH.xlsm");
