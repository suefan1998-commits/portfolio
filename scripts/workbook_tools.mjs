import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const HEADERS = [
  "作品ID",
  "标题",
  "来源类型",
  "链接/文件名",
  "撰稿类型",
  "主题标签",
  "发布媒体",
  "发布日期",
  "是否公开展示",
  "是否为精选作品",
  "处理状态",
  "备注",
  "作品图片",
];

const STATUS_VALUES = ["待处理", "已入库", "需手动补充", "错误"];
const SOURCE_VALUES = ["新媒体", "纸刊"];
const PUBLIC_VALUES = ["是", "否"];
const WRITING_TYPE_VALUES = [
  "人物与访谈（Profiles & Interviews）",
  "城市、旅行与生活方式（Travel & Lifestyle）",
  "设计、建筑与文化（Design Architecture & Culture）",
  "商业、科技与社会（Business Technology & Society）",
  "品牌特稿（Brand Features）",
];

async function createTemplate(outputPath) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("作品台账");
  sheet.showGridLines = false;

  sheet.getRange("A1:M1").values = [HEADERS];
  sheet.getRange("A1:M1").format = {
    fill: { color: "#EAE7DE" },
    font: { color: "#111827", bold: true },
    wrapText: true,
  };
  sheet.getRange("A1:M1").format.borders = {
    top: { style: "thin", color: "#111111" },
    bottom: { style: "thin", color: "#111111" },
  };

  sheet.getRange("A2:M201").format = {
    font: { color: "#111827" },
    wrapText: true,
  };
  sheet.getRange("A2:M201").format.borders = {
    insideHorizontal: { style: "thin", color: "#E5E7EB" },
  };

  sheet.getRange("A:A").format.columnWidth = 14;
  sheet.getRange("B:B").format.columnWidth = 24;
  sheet.getRange("C:C").format.columnWidth = 12;
  sheet.getRange("D:D").format.columnWidth = 42;
  sheet.getRange("E:E").format.columnWidth = 16;
  sheet.getRange("F:F").format.columnWidth = 20;
  sheet.getRange("G:G").format.columnWidth = 18;
  sheet.getRange("H:H").format.columnWidth = 14;
  sheet.getRange("I:I").format.columnWidth = 14;
  sheet.getRange("J:J").format.columnWidth = 16;
  sheet.getRange("K:K").format.columnWidth = 14;
  sheet.getRange("L:L").format.columnWidth = 28;
  sheet.getRange("M:M").format.columnWidth = 28;
  sheet.getRange("H2:H201").setNumberFormat("yyyy-mm-dd");

  sheet.dataValidations.add({
    range: "C2:C201",
    rule: { type: "list", values: SOURCE_VALUES },
  });
  sheet.dataValidations.add({
    range: "E2:E201",
    rule: { type: "list", values: WRITING_TYPE_VALUES },
  });
  sheet.dataValidations.add({
    range: "I2:I201",
    rule: { type: "list", values: PUBLIC_VALUES },
  });
  sheet.dataValidations.add({
    range: "J2:J201",
    rule: { type: "list", values: PUBLIC_VALUES },
  });
  sheet.dataValidations.add({
    range: "K2:K201",
    rule: { type: "list", values: STATUS_VALUES },
  });

  sheet.freezePanes.freezeRows(1);

  const guide = workbook.worksheets.add("填写说明");
  guide.showGridLines = false;
  const guideRows = [
    ["字段", "填写说明"],
    ["作品ID", "可留空，整理入库时自动生成。"],
    ["标题", "可留空；新媒体链接会尽量从网页标题提取，纸刊文件名会作为兜底标题。"],
    ["来源类型", "填写“新媒体”或“纸刊”。"],
    ["链接/文件名", "新媒体填链接；纸刊填 inbox 文件夹内的文件名。"],
    ["撰稿类型", "主分类，例如文旅、人物、品牌、评论、报道。"],
    ["主题标签", "辅助筛选标签，多个标签用顿号、逗号或分号分隔。"],
    ["发布媒体", "填写发布平台、刊物或客户名称。"],
    ["发布日期", "建议使用 yyyy-mm-dd。"],
    ["是否公开展示", "填写“是”或“否”。"],
    ["是否为精选作品", "填写“是”或“否”；首页精选作品优先读取这一列。"],
    ["处理状态", "新作品填“待处理”；脚本会更新为“已入库”、“需手动补充”或“错误”。"],
    ["作品图片", "可选。填写本地图片路径，例如 20260629-男人风尚.jpg；用于首页精选作品、作品总览和单篇文章页。"],
  ];
  guide.getRangeByIndexes(0, 0, guideRows.length, 2).values = guideRows;
  guide.getRange("A1:B1").format = {
    fill: { color: "#263238" },
    font: { color: "#FFFFFF", bold: true },
  };
  guide.getRange("A:B").format.columnWidth = 28;
  guide.getRange("B:B").format.columnWidth = 72;
  guide.getRange("A1:B14").format = { wrapText: true };
  guide.freezePanes.freezeRows(1);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);

  const preview = await workbook.render({
    sheetName: "作品台账",
    range: "A1:M12",
    scale: 1,
    format: "png",
  });
  const previewBytes = new Uint8Array(await preview.arrayBuffer());
  await fs.mkdir(path.join(path.dirname(outputPath), "tmp"), { recursive: true });
  await fs.writeFile(path.join(path.dirname(outputPath), "tmp", "registry_preview.png"), previewBytes);
}

async function updateStatus(workbookPath, updatesPath) {
  const input = await FileBlob.load(workbookPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheet = workbook.worksheets.getItem("作品台账");
  const used = sheet.getUsedRange(true);
  const values = used.values || [];
  const headers = values[0] || [];
  const idCol = headers.indexOf("作品ID");
  const titleCol = headers.indexOf("标题");
  const statusCol = headers.indexOf("处理状态");
  const noteCol = headers.indexOf("备注");
  if (idCol < 0 || statusCol < 0 || noteCol < 0) {
    throw new Error("台账缺少必要列：作品ID、处理状态、备注");
  }

  const updates = JSON.parse(await fs.readFile(updatesPath, "utf8"));
  for (const update of updates) {
    const rowIndex = Number(update.rowIndex);
    if (!Number.isInteger(rowIndex) || rowIndex < 2) continue;
    if (update.id) {
      sheet.getCell(rowIndex - 1, idCol).values = [[update.id]];
    }
    if (update.title && titleCol >= 0) {
      sheet.getCell(rowIndex - 1, titleCol).values = [[update.title]];
    }
    if (update.status) {
      sheet.getCell(rowIndex - 1, statusCol).values = [[update.status]];
    }
    if (update.note !== undefined) {
      sheet.getCell(rowIndex - 1, noteCol).values = [[update.note]];
    }
  }

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(workbookPath);
}

async function main() {
  const [cmd, ...args] = process.argv.slice(2);
  if (cmd === "create-template") {
    await createTemplate(args[0] || "portfolio_registry.xlsx");
    return;
  }
  if (cmd === "update-status") {
    await updateStatus(args[0], args[1]);
    return;
  }
  throw new Error(`未知工作簿操作：${cmd || ""}`);
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exit(1);
});
