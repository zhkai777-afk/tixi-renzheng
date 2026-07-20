import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve("../..");
const outputDir = path.resolve(".");
const catalogPath = path.join(root, "materials-catalog.json");
const catalog = JSON.parse(await fs.readFile(catalogPath, "utf8"));

const codeBySortOrder = {
  1: "MSF11-19",
  2: "MSF11-18",
  3: "MSF11-01",
  5: "MSF11-23",
  6: "MSF11-25",
  7: "MSF11-16",
  8: "MSF11-02",
  12: "MSF11-04",
  14: "MSF11-05",
  15: "MSF11-03",
  32: "MSF11-27",
  33: "MSF11-08",
  34: "MSF11-07",
  35: "MSF11-14",
  36: "MSF11-11",
  37: "MSF11-10",
  38: "MSF11-09",
  39: "MSF11-21",
  40: "MSF11-07",
  41: "MSF11-12",
  42: "MSF11-14",
  43: "MSF11-11",
  44: "MSF11-15",
  45: "MSF11-13",
  46: "MSF11-17",
};

const recommendationAuditKeys = new Set([
  "initialCertification",
  "recertification",
  "certificateChange",
  "certificationTransfer",
]);

const workbook = Workbook.create();

function displayName(material) {
  const code = codeBySortOrder[material.sortOrder];
  if (!code) return material.materialName;
  return `${material.materialName.replace(/（.*?）$/, "")}（${code}）`;
}

function systemDisplay(material) {
  if (material.rawApplicableSystems) return material.rawApplicableSystems;
  if ((material.applicableSystems || []).join("/") === "Q/E/S/F") return "通用";
  return (material.applicableSystems || []).join("/") || "通用";
}

function systemRule(display) {
  switch (display) {
    case "通用":
      return "项目包含任一 Q/E/S/F 体系时适用";
    case "Q除外":
      return "项目包含 E/S/F 任一体系时适用；仅 Q 体系项目不校验";
    case "Q/F/H":
      return "项目包含 Q 或 F 体系时适用；H 如后续启用，按 HACCP/相关食品安全扩展体系处理";
    case "E/S":
      return "项目包含 E 或 S 体系时适用";
    case "E":
      return "项目包含 E 体系时适用";
    case "S":
      return "项目包含 S 体系时适用";
    case "F/H":
      return "项目包含 F 体系时适用；H 如后续启用，按 HACCP/相关食品安全扩展体系处理";
    case "IS":
      return "仅信息安全 IS 体系项目适用；本批 Q/E/S/F 项目不校验";
    default:
      return `按适用体系“${display}”命中项目体系时校验`;
  }
}

function markerMeaning(raw) {
  if (raw.includes("◆")) return "适用时或进行现场审核时必须上报";
  if (raw.includes("★")) return "适用时必须上报";
  if (raw.includes("●")) return "由各地公司归档保留；若平台统一管控，作为归档必有项校验";
  if (raw.includes("系统上填报")) return "系统上填报，按截图口径作为系统必有文件";
  return "按截图口径作为系统必有文件";
}

function validationRule(auditLabel, material, raw) {
  const display = systemDisplay(material);
  return `当业务类型=${auditLabel}，且${systemRule(display)}时，${markerMeaning(raw)}。`;
}

function applySheetStyle(sheet, rowCount) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(3);
  sheet.getRange("A1:G1").merge();
  sheet.getRange("A1:G1").format.fill.color = "#E8F1FF";
  sheet.getRange("A1:G1").format.font.bold = true;
  sheet.getRange("A1:G1").format.font.size = 15;
  sheet.getRange("A1:G1").format.font.color = "#0052D9";
  sheet.getRange("A1:G1").format.rowHeight = 28;
  sheet.getRange("A2:G2").format.fill.color = "#FAFAFA";
  sheet.getRange("A2:G2").format.font.bold = true;
  sheet.getRange("A2:G2").format.font.size = 11;
  sheet.getRange("A2:G2").format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };
  sheet.getRange(`A3:G${Math.max(rowCount, 3)}`).format.borders = { preset: "all", style: "thin", color: "#E7E7E7" };
  sheet.getRange(`A3:G${Math.max(rowCount, 3)}`).format.font.size = 11;
  sheet.getRange(`A3:G${Math.max(rowCount, 3)}`).format.wrapText = true;
  sheet.getRange(`A1:G${Math.max(rowCount, 3)}`).format.verticalAlignment = "center";
  sheet.getRange("A:A").format.columnWidth = 8;
  sheet.getRange("B:B").format.columnWidth = 46;
  sheet.getRange("C:C").format.columnWidth = 14;
  sheet.getRange("D:D").format.columnWidth = 14;
  sheet.getRange("E:E").format.columnWidth = 12;
  sheet.getRange("F:F").format.columnWidth = 24;
  sheet.getRange("G:G").format.columnWidth = 72;
  sheet.getRange("A:A").format.horizontalAlignment = "center";
  sheet.getRange("C:E").format.horizontalAlignment = "center";
}

const summarySheet = workbook.worksheets.add("说明");
summarySheet.getRange("A1:D1").values = [["业务类型", "规则行数", "整理口径", "来源"]];
summarySheet.getRange("A1:D1").format.fill.color = "#FAFAFA";
summarySheet.getRange("A1:D1").format.font.bold = true;
summarySheet.getRange("A1:D1").format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };
summarySheet.showGridLines = false;
summarySheet.freezePanes.freezeRows(1);
summarySheet.getRange("A:A").format.columnWidth = 18;
summarySheet.getRange("B:B").format.columnWidth = 12;
summarySheet.getRange("C:C").format.columnWidth = 52;
summarySheet.getRange("D:D").format.columnWidth = 48;
summarySheet.getRange("A:D").format.wrapText = true;

const summaryRows = [];

const ruleSheet = workbook.worksheets.add("规则口径");
const ruleRows = [
  ["字段", "说明"],
  ["适用体系=通用", "项目包含任一 Q/E/S/F 体系时，该材料对当前业务类型参与必传校验。"],
  ["适用体系=Q除外", "项目包含 E/S/F 任一体系时参与校验；仅 Q 体系项目不校验。"],
  ["适用体系=Q/F/H", "项目包含 Q 或 F 体系时参与校验；H 如后续启用，按 HACCP/相关食品安全扩展体系处理。"],
  ["适用体系=E/S", "项目包含 E 或 S 体系时参与校验。"],
  ["适用体系=E / S / F/H / IS", "分别按对应体系命中时参与校验；IS 不属于本批 Q/E/S/F，只有信息安全体系项目启用。"],
  ["原表标识=★", "在适用时必须上报。"],
  ["原表标识=◆", "适用时或进行现场审核时必须上报。"],
  ["原表标识=●", "由各地公司归档保留；若平台统一纳管归档完整性，可作为归档必有项校验。"],
  ["认证发证推荐表", "MSF11-20 原表描述为系统上填报；按用户截图口径纳入发证相关业务类型的系统必有项。"],
];
ruleSheet.getRangeByIndexes(0, 0, ruleRows.length, 2).values = ruleRows;
ruleSheet.showGridLines = false;
ruleSheet.freezePanes.freezeRows(1);
ruleSheet.getRange("A1:B1").format.fill.color = "#FAFAFA";
ruleSheet.getRange("A1:B1").format.font.bold = true;
ruleSheet.getRange(`A1:B${ruleRows.length}`).format.borders = { preset: "all", style: "thin", color: "#E7E7E7" };
ruleSheet.getRange(`A1:B${ruleRows.length}`).format.wrapText = true;
ruleSheet.getRange("A:A").format.columnWidth = 26;
ruleSheet.getRange("B:B").format.columnWidth = 86;

for (const auditType of catalog.auditTypes) {
  const rows = catalog.materials
    .filter((material) => {
      if (material.sortOrder === 1) return recommendationAuditKeys.has(auditType.key);
      return Boolean(material.requirements[auditType.key]?.required);
    })
    .map((material) => {
      const raw = material.sortOrder === 1
        ? "系统上填报"
        : material.requirements[auditType.key]?.raw || "";
      return [
        material.sortOrder,
        displayName(material),
        systemDisplay(material),
        raw,
        "是",
        "",
        validationRule(auditType.label, material, raw),
      ];
    });

  const sheet = workbook.worksheets.add(auditType.label);
  sheet.getRange("A1:G1").values = [[auditType.label]];
  sheet.getRange("A2:G2").values = [[
    "序号",
    "文件名称",
    "适用体系",
    "原表标识",
    "是否必传",
    "系统中数据获取节点",
    "校验规则说明",
  ]];
  if (rows.length) {
    sheet.getRangeByIndexes(2, 0, rows.length, 7).values = rows;
  }
  applySheetStyle(sheet, rows.length + 2);
  summaryRows.push([
    auditType.label,
    rows.length,
    "按业务类型提取原表中 ★/◆/● 标识项，并补充适用体系命中条件；按截图口径将“认证发证推荐表（MSF11-19）”作为发证相关业务类型的系统必有项；获取节点列留空。",
    "materials-catalog.json；MSF11-20 审核材料归档/上报清单（2025.12.31修订）",
  ]);
}

summarySheet.getRangeByIndexes(1, 0, summaryRows.length, 4).values = summaryRows;
summarySheet.getRange(`A2:D${summaryRows.length + 1}`).format.borders = { preset: "all", style: "thin", color: "#E7E7E7" };
summarySheet.getRange(`A1:D${summaryRows.length + 1}`).format.verticalAlignment = "center";
summarySheet.getRange("B:B").format.horizontalAlignment = "center";

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "业务类型必传文件清单.xlsx"));
