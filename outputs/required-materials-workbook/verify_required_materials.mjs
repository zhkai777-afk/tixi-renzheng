import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = path.resolve(".");
const workbookPath = path.join(outputDir, "业务类型必传文件清单.xlsx");
const input = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheetSummary = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 4000,
});
console.log(sheetSummary.ndjson);

for (const sheetName of ["说明", "规则口径", "初次认证", "监督", "再认证", "证书变更", "认证转换", "临时审核"]) {
  const preview = await workbook.inspect({
    kind: "region",
    sheetId: sheetName,
    range: sheetName === "说明" || sheetName === "规则口径" ? "A1:D12" : "A1:G12",
    maxChars: 3000,
  });
  console.log(preview.ndjson);
  const blob = await workbook.render({
    sheetName,
    range: sheetName === "说明" || sheetName === "规则口径" ? "A1:D12" : "A1:G12",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), new Uint8Array(await blob.arrayBuffer()));
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);
