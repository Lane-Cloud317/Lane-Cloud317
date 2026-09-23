import fs from "node:fs/promises";
import JSZip from "jszip";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const base = "D:/手工站位及表單優化";
const source = `${base}/企劃部手動作業站位風險排查(2).xlsx`;
const outputDir = `${base}/outputs/手動作業站位風險排查_優化版`;
const outputPath = `${outputDir}/企劃部手動作業站位風險排查_優化版.xlsx`;
const original = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const sourceSheet = original.worksheets.getItemAt(0);
const data = sourceSheet.getRange("B12:T30").values;

const measures = [
  "推進送貨預約系統上線，串接到貨資料與收貨掃碼；AI輔助辨識送貨單欄位及差異，異常由人員確認。",
  "導入儲位管理與掃碼過帳，保留定期盤點；AI分析帳實差異並提示異常儲位。",
  "導入掃碼發料，核對料號、批次與數量；AI彙整錯發、短發異常供主管追蹤。",
  "導入儲位系統及掃碼調撥，核對過帳與實物移動；AI提示未閉環調撥。",
  "按計畫導入WMS及PDA出貨點檢，保留雙人複核；AI比對訂單、箱單與掃碼結果。",
  "維持每週PR清單複核；以規則比對需求、庫存與採購數量，AI提示異常供物控確認。",
  "定期維護BOM清單，清除EOL機種並新增機種；AI比對版本差異，人工核准後更新。",
  "以腳本按料件彙總WDN交期並按固定模板輸出；AI提示交期衝突與異常變動。",
  "維持SAP庫存容差比對；AI依歷史偏差提示需覆盤料件，差異由人員核實。",
  "保留MPS與WDN總量交叉複核；以公式或程式檢查漏連、錯連，AI輔助說明差異。",
  "建立Config基表及批量開單模板，料號由公式帶入；SAP校驗料號與數量，SFC比對MPS配置，並優化WIP流程。AI提示配置異常，人工核准。",
  "倉庫隔離並鎖定不可用物料；發料前校驗物料狀態與批次，AI提示異常批次供人工複核。",
  "從SAP匯出工單資料，以公式帶入實計畫；AI提示投入量與原料結餘異常，物控複核。",
  "維持交叉檢查與ASN卡控；AI擷取並比對報關單、承運單、箱單、發票及ASN欄位，異常由人員確認。",
  "保留Relabel系統下達與人工複核；AI比對托運、供應商及進口單證欄位，異常由人員確認。",
  "保留逐項複核；AI擷取備案欄位並提供商編候選，歸類由關務審核確認。",
  "維持電子簽核；AI比對費用申請、付款憑證與成本分類，提示差異供審核。",
  "保留核銷資料逐項複核；AI擷取申報欄位並比對報關、帳冊及庫存數量，差異由關務確認。",
  "保留搬運進度表交叉檢查；AI彙整搬運節點及逾期異常，人工確認後更新進度。",
];

const compact = [
  ["碼頭收貨／貨物資訊點檢", "漏收或資訊缺漏", "收貨掃描與系統驗收", "系統防呆"],
  ["原料入庫／儲存與過帳", "漏過帳、管制卡錯誤或實物難尋", "定期盤點", "人工交叉複核"],
  ["原料發料／人工配料", "多發、少發或發錯料", "與產線逐項對點", "人工交叉複核"],
  ["庫存調撥／儲位異動", "僅過帳而未移動實物", "每日盤點", "人工交叉複核"],
  ["碼頭出貨／資訊點檢", "出貨資訊錯誤", "倉庫與物流雙人點檢", "人工交叉複核"],
  ["PR開立／數量輸入", "輸入錯誤致欠料或呆滯", "每週PR清單主管複核", "人工交叉複核"],
  ["基礎資訊／BOM維護", "物料資訊錯誤致欠料或呆滯", "BOM用量與機種數複核", "人工交叉複核"],
  ["交期管理／WDN更新", "交期貼錯致齊套誤判", "按料件彙總WDN交期", "人工交叉複核"],
  ["庫存實際化／數量維護", "輸入錯誤致欠料或呆滯", "與SAP庫存容差比對", "人工交叉複核"],
  ["需求核算／公式連結", "漏連或錯連需求", "MPS與WDN總量複核", "人工交叉複核"],
  ["開立工單／料號與數量", "料號或數量錯誤致呆滯", "MFG複查並比對MPS", "人工交叉複核"],
  ["開立發料單／批次輸入", "不可用物料發至產線", "倉庫隔離不可用物料", "人工交叉複核"],
  ["實計畫／投入數量", "實計畫數量錯誤影響採購", "物控交叉複核原料結餘", "人工交叉複核"],
  ["出口／單證與ASN", "錄入錯誤影響報關及交付", "交叉檢查、出貨點檢與ASN", "系統＋人工複核"],
  ["進口／單證與Relabel", "錄入錯誤影響報關及運輸", "交叉檢查與Relabel系統", "系統＋人工複核"],
  ["備案／資料及商編", "錄入或歸類錯誤", "關務逐項交叉檢查", "人工交叉複核"],
  ["結報／費用與成本", "錄入錯誤影響付款或統計", "電子簽核系統", "系統防呆"],
  ["核銷／報關與帳冊", "錄入錯誤致核銷異常", "申報表與帳冊逐項複核", "人工交叉複核"],
  ["設備搬運／進度表", "錄入錯誤影響統計", "搬運進度交叉檢查", "人工交叉複核"],
];

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("手動作業站位風險排查");
sheet.showGridLines = false;
sheet.mergeCells("A1:G1");
sheet.getRange("A1").values = [["企劃部手動作業站位風險排查"]];
sheet.mergeCells("A2:G2");
sheet.getRange("A2").values = [["共19項｜綠燈19項｜判定：系統卡控或人工交叉複核均為有效卡控；AI措施為建議方向，實施仍需驗證。"]];
sheet.getRange("A3:G3").values = [["序號", "課別／站位／手動作業", "潛在風險", "現行卡控措施／方式", "燈號", "檢討整改措施", "責任人／期限／進度"]];

for (let i = 0; i < 19; i++) {
  const src = data[i];
  let deadline = src[15] ?? "";
  if (typeof deadline === "number" && deadline > 40000) {
    const dt = new Date(Date.UTC(1899, 11, 30) + deadline * 86400000);
    deadline = `${dt.getUTCFullYear()}/${String(dt.getUTCMonth() + 1).padStart(2, "0")}/${String(dt.getUTCDate()).padStart(2, "0")}`;
  }
  if (i === 10) deadline = "2026/10/15、10/30";
  const progress = String(src[16] ?? "").replace("进行中", "進行中").replace("未开始", "未開始");
  sheet.getRange(`A${i + 4}:G${i + 4}`).values = [[
    i + 1,
    `${String(src[1] ?? "").replace("关", "關")}｜${compact[i][0]}`,
    compact[i][1], `${compact[i][3]}｜${compact[i][2]}`, null,
    measures[i], [String(src[14] ?? "").replace(/\n/g, "、"), String(deadline).replace(/\n/g, "、"), progress].filter(Boolean).join("｜"),
  ]];
  sheet.getRange(`E${i + 4}`).formulas = [[`=IF(D${i + 4}="","待完善",IF(ISNUMBER(SEARCH("無防呆",D${i + 4})),"紅燈",IF(OR(ISNUMBER(SEARCH("系統",D${i + 4})),ISNUMBER(SEARCH("人工",D${i + 4}))),"綠燈","待完善")))`]];
}

sheet.getRange("A1:G22").format.font = { name: "Microsoft JhengHei", size: 9, color: "#17324A" };
sheet.getRange("A1:G1").format = { fill: "#17324A", font: { name: "Microsoft JhengHei", size: 13, bold: true, color: "#FFFFFF" } };
sheet.getRange("A1:G1").format.rowHeight = 28;
sheet.getRange("A2:G2").format = { fill: "#E8F0F6", font: { name: "Microsoft JhengHei", size: 9, color: "#17324A" } };
sheet.getRange("A2:G2").format.rowHeight = 24;
sheet.getRange("A3:G3").format = { fill: "#285A78", font: { name: "Microsoft JhengHei", size: 9, bold: true, color: "#FFFFFF" } };
sheet.getRange("A3:G3").format.rowHeight = 26;
sheet.getRange("A4:G22").format.rowHeight = 43;
sheet.getRange("A14:G14").format.rowHeight = 72;
sheet.getRange("A4:G22").format.verticalAlignment = "center";
sheet.getRange("B4:G22").format.wrapText = true;
sheet.getRange("A4:A22").format.columnWidth = 5;
sheet.getRange("B4:B22").format.columnWidth = 24;
sheet.getRange("C4:C22").format.columnWidth = 17;
sheet.getRange("D4:D22").format.columnWidth = 23;
sheet.getRange("E4:E22").format.columnWidth = 7;
sheet.getRange("F4:F22").format.columnWidth = 49;
sheet.getRange("G4:G22").format.columnWidth = 17;
sheet.getRange("A3:G3").format.borders = { preset: "outside", style: "thin", color: "#FFFFFF" };
sheet.getRange("A4:G22").format.borders = { preset: "all", style: "thin", color: "#D9E2E9" };
sheet.getRange("E4:E22").format = { fill: "#DDF3DF", font: { name: "Microsoft JhengHei", size: 9, bold: true, color: "#1D7135" } };
sheet.getRange("E4:E22").format.verticalAlignment = "center";
sheet.getRange("A4:A22").format.font = { name: "Microsoft JhengHei", size: 8, bold: true, color: "#285A78" };

const preview = await workbook.render({ sheetName: sheet.name, range: "A1:G22", scale: 1.3, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(`${outputDir}/版面預覽.png`, new Uint8Array(await preview.arrayBuffer()));
const inspection = await workbook.inspect({ kind: "table", range: "A1:G22", include: "values,formulas", tableMaxRows: 22, tableMaxCols: 7, maxChars: 28000 });
console.log(inspection.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan" });
console.log(errors.ndjson);
const cellValues = sheet.getRange("A1:G22").values;
const lampFormulas = sheet.getRange("E4:E22").formulas;
const blob = await SpreadsheetFile.exportXlsx(workbook);
await blob.save(outputPath);

const zip = await JSZip.loadAsync(await fs.readFile(outputPath));
const wsName = "xl/worksheets/sheet1.xml";
let xml = await zip.file(wsName).async("string");
const xmlEscape = value => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");
const colIndex = letter => letter.charCodeAt(0) - 65;
xml = xml.replace(/<x:c r="([A-G])(\d+)"([^>]*)>([\s\S]*?)<\/x:c>/g, (full, col, rowText, attrs, body) => {
  const row = Number(rowText);
  const value = cellValues[row - 1]?.[colIndex(col)];
  if (value === undefined || value === null) return full;
  let repaired = body;
  if (col === "E" && row >= 4 && row <= 22) {
    const formula = String(lampFormulas[row - 4]?.[0] ?? "").replace(/^=/, "");
    repaired = repaired.replace(/<x:f>[\s\S]*?<\/x:f>/, `<x:f>${xmlEscape(formula)}</x:f>`);
  }
  repaired = repaired.replace(/<x:v>[\s\S]*?<\/x:v>/, `<x:v>${xmlEscape(value)}</x:v>`);
  return `<x:c r="${col}${rowText}"${attrs}>${repaired}</x:c>`;
});
xml = xml.replace(/<x:worksheet([^>]*)>/, '<x:worksheet$1><x:sheetPr><x:pageSetUpPr fitToPage="1"/></x:sheetPr>');
xml = xml.replace(/<x:pageMargins[^>]*\/>/, '<x:printOptions horizontalCentered="1"/><x:pageMargins left="0.18" right="0.18" top="0.20" bottom="0.20" header="0.10" footer="0.10"/><x:pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1"/>');
zip.file(wsName, xml);
let wbXml = await zip.file("xl/workbook.xml").async("string");
wbXml = wbXml.replace(/<x:sheet name="[^"]+"/, `<x:sheet name="${xmlEscape(sheet.name)}"`);
const printArea = '<x:definedNames><x:definedName name="_xlnm.Print_Area" localSheetId="0">\'手動作業站位風險排查\'!$A$1:$G$22</x:definedName></x:definedNames>';
wbXml = wbXml.replace(/<\/x:workbook>/, `${printArea}</x:workbook>`);
zip.file("xl/workbook.xml", wbXml);
await fs.writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
console.log(`OUTPUT ${outputPath}`);
