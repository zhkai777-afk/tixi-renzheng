import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = "/Users/zhazhakai777/Documents/体系认证/outputs/materials-confirmation/体系认证-材料归档需求待确认问题清单.xlsx";
const input = await FileBlob.load(path);
const workbook = await SpreadsheetFile.importXlsx(input);
const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  range: "1-需求冲突!A1:I8",
  include: "values",
  tableMaxRows: 8,
  tableMaxCols: 9,
  maxChars: 12000,
});
console.log(summary.ndjson);
