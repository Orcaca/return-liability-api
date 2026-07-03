const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "80mb" }));

const fileStore = new Map();

app.get("/", (req, res) => {
  res.send("반품충당부채 XLSX 재투입용 결과파일 생성 API 서버가 실행 중입니다.");
});

function normalize(value) {
  return String(value ?? "")
    .replace(/\n/g, "")
    .replace(/\s/g, "")
    .trim();
}

function parseCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    const next = line[i + 1];

    if (ch === '"' && inQuotes && next === '"') {
      current += '"';
      i++;
    } else if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (ch === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += ch;
    }
  }

  result.push(current);
  return result;
}

function csvToObjects(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .filter((line) => line.trim() !== "");

  if (lines.length === 0) {
    return [];
  }

  const headers = parseCsvLine(lines[0]).map((h) => String(h || "").trim());

  return lines.slice(1).map((line) => {
    const values = parseCsvLine(line);
    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = values[index] ?? "";
    });

    return obj;
  });
}

function csvToRows(csvText) {
  return String(csvText || "")
    .split(/\r?\n/)
    .filter((line) => line.trim() !== "")
    .map(parseCsvLine);
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return value;
  }

  const text = String(value).replace(/,/g, "").trim();

  if (text === "") {
    return 0;
  }

  const num = Number(text);

  if (Number.isFinite(num)) {
    return num;
  }

  return 0;
}

function toCellValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "number") {
    return value;
  }

  const text = String(value).trim();

  if (text === "") {
    return "";
  }

  const num = Number(text.replace(/,/g, ""));

  if (Number.isFinite(num) && /^-?\d+(\.\d+)?(e-?\d+)?$/i.test(text.replace(/,/g, ""))) {
    return num;
  }

  return text;
}

function cleanBase64(base64Text) {
  return String(base64Text || "").replace(/^data:.*?;base64,/, "");
}

function cloneValue(value) {
  if (value === null || value === undefined) {
    return value;
  }

  if (typeof value === "object") {
    return JSON.parse(JSON.stringify(value));
  }

  return value;
}

function getCellValue(worksheet, rowNo, colNo) {
  return worksheet.getRow(rowNo).getCell(colNo).value;
}

function setCellValue(worksheet, rowNo, colNo, value) {
  worksheet.getRow(rowNo).getCell(colNo).value = value;
}

function findLiabilityBlocks(worksheet) {
  const blocks = [];

  for (let rowNo = 1; rowNo <= worksheet.rowCount; rowNo++) {
    const row = worksheet.getRow(rowNo);

    const b = normalize(row.getCell(2).value);
    const c = normalize(row.getCell(3).value);

    if (b === "구분" && c.includes("총액매출")) {
      blocks.push({
        headerRow: rowNo,
        dataStartRow: rowNo + 1,
        dataEndRow: null,
        totalRow: null
      });
    }
  }

  if (blocks.length < 3) {
    throw new Error("반품충당부채 시트에서 계산표 3개를 찾지 못했습니다.");
  }

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const nextHeaderRow = i + 1 < blocks.length ? blocks[i + 1].headerRow : worksheet.rowCount + 1;

    let endRow = block.dataStartRow;

    for (let rowNo = block.dataStartRow; rowNo < nextHeaderRow; rowNo++) {
      const row = worksheet.getRow(rowNo);

      let hasValue = false;

      for (let colNo = 1; colNo <= 20; colNo++) {
        const value = row.getCell(colNo).value;

        if (value !== null && value !== undefined && String(value).trim() !== "") {
          hasValue = true;
          break;
        }
      }

      if (!hasValue) {
        break;
      }

      endRow = rowNo;
    }

    block.dataEndRow = endRow;

    let totalRow = null;

    for (let rowNo = block.dataEndRow; rowNo >= block.dataStartRow; rowNo--) {
      const row = worksheet.getRow(rowNo);
      const a = normalize(row.getCell(1).value);
      const b = normalize(row.getCell(2).value);
      const c = row.getCell(3).value;

      if (a === "" && b === "" && c !== null && c !== undefined && String(c).trim() !== "") {
        totalRow = rowNo;
        break;
      }
    }

    if (!totalRow) {
      totalRow = block.dataEndRow;
    }

    block.totalRow = totalRow;
  }

  return blocks.slice(0, 3);
}

function readBlockSnapshot(worksheet, block) {
  const rows = [];

  for (let rowNo = block.dataStartRow; rowNo <= block.dataEndRow; rowNo++) {
    const rowValues = [];

    for (let colNo = 1; colNo <= 20; colNo++) {
      rowValues.push(cloneValue(getCellValue(worksheet, rowNo, colNo)));
    }

    rows.push(rowValues);
  }

  return rows;
}

function writeBlockSnapshot(worksheet, block, snapshot) {
  const rowCount = block.dataEndRow - block.dataStartRow + 1;

  for (let i = 0; i < rowCount; i++) {
    const targetRowNo = block.dataStartRow + i;
    const sourceValues = snapshot[i] || [];

    for (let colNo = 1; colNo <= 20; colNo++) {
      setCellValue(worksheet, targetRowNo, colNo, cloneValue(sourceValues[colNo - 1] ?? ""));
    }
  }
}

function clearBlockDataArea(worksheet, block) {
  for (let rowNo = block.dataStartRow; rowNo <= block.dataEndRow; rowNo++) {
    for (let colNo = 1; colNo <= 20; colNo++) {
      setCellValue(worksheet, rowNo, colNo, "");
    }
  }
}

function writeCalculatedRowsToBlock(worksheet, block, detailObjects, priorBlock) {
  clearBlockDataArea(worksheet, block);

  const dataStart = block.dataStartRow;
  const totalRow = block.totalRow;
  const maxDataRows = Math.max(0, totalRow - dataStart);

  const rowsToWrite = detailObjects.slice(0, maxDataRows);

  const sums = {
    c: 0,
    d: 0,
    e: 0,
    f: 0,
    g: 0,
    h: 0,
    i: 0,
    o: 0,
    p: 0,
    q: 0,
    r: 0,
    s: 0
  };

  rowsToWrite.forEach((item, index) => {
    const rowNo = dataStart + index;

    const c = toNumber(item["총액매출"]);
    const d = toNumber(item["당해매출반품"]);
    const e = toNumber(item["1년이상반품"]);
    const f = toNumber(item["1년"]);
    const g = toNumber(item["2년"]);
    const h = toNumber(item["순매출액"]);
    const i = toNumber(item["원가금액"]);
    const j = toNumber(item["원가율"]);

    const k = toNumber(item["당해1년반품율"]);
    const l = toNumber(item["당해2년반품율"]);
    const m = toNumber(item["적용1년반품율"]);
    const n = toNumber(item["적용2년반품율"]);

    const o = toNumber(item["당해매출기준_반품추정액"]);
    const p = toNumber(item["전기매출기준_반품추정액"]);
    const q = toNumber(item["당해매출기준_원가추정액"]);
    const r = toNumber(item["전기매출기준_원가추정액"]);
    const s = toNumber(item["순충당부채"]);

    setCellValue(worksheet, rowNo, 1, item["대분류"] || "");
    setCellValue(worksheet, rowNo, 2, item["구분"] || "");
    setCellValue(worksheet, rowNo, 3, c);
    setCellValue(worksheet, rowNo, 4, d);
    setCellValue(worksheet, rowNo, 5, e);
    setCellValue(worksheet, rowNo, 6, f);
    setCellValue(worksheet, rowNo, 7, g);
    setCellValue(worksheet, rowNo, 8, h);
    setCellValue(worksheet, rowNo, 9, i);
    setCellValue(worksheet, rowNo, 10, j);
    setCellValue(worksheet, rowNo, 11, k);
    setCellValue(worksheet, rowNo, 12, l);
    setCellValue(worksheet, rowNo, 13, m);
    setCellValue(worksheet, rowNo, 14, n);
    setCellValue(worksheet, rowNo, 15, o);
    setCellValue(worksheet, rowNo, 16, p);
    setCellValue(worksheet, rowNo, 17, q);
    setCellValue(worksheet, rowNo, 18, r);
    setCellValue(worksheet, rowNo, 19, s);
    setCellValue(worksheet, rowNo, 20, 0);

    sums.c += c;
    sums.d += d;
    sums.e += e;
    sums.f += f;
    sums.g += g;
    sums.h += h;
    sums.i += i;
    sums.o += o;
    sums.p += p;
    sums.q += q;
    sums.r += r;
    sums.s += s;
  });

  const priorTotalSales = priorBlock
    ? toNumber(getCellValue(worksheet, priorBlock.totalRow, 3))
    : 0;

  const totalCostRate = sums.h !== 0 ? sums.i / sums.h : 0;
  const currentRate1 = sums.c !== 0 ? -sums.f / sums.c : 0;
  const currentRate2 = sums.c !== 0 ? -sums.g / sums.c : 0;

  const selectedRate2 = priorTotalSales !== 0 ? sums.p / priorTotalSales : 0;
  const selectedRateTotal = sums.c !== 0 ? sums.o / sums.c : 0;
  const selectedRate1 = selectedRateTotal - selectedRate2;

  setCellValue(worksheet, totalRow, 1, "");
  setCellValue(worksheet, totalRow, 2, "");
  setCellValue(worksheet, totalRow, 3, sums.c);
  setCellValue(worksheet, totalRow, 4, sums.d);
  setCellValue(worksheet, totalRow, 5, sums.e);
  setCellValue(worksheet, totalRow, 6, sums.f);
  setCellValue(worksheet, totalRow, 7, sums.g);
  setCellValue(worksheet, totalRow, 8, sums.h);
  setCellValue(worksheet, totalRow, 9, sums.i);
  setCellValue(worksheet, totalRow, 10, totalCostRate);
  setCellValue(worksheet, totalRow, 11, currentRate1);
  setCellValue(worksheet, totalRow, 12, currentRate2);
  setCellValue(worksheet, totalRow, 13, selectedRate1);
  setCellValue(worksheet, totalRow, 14, selectedRate2);
  setCellValue(worksheet, totalRow, 15, sums.o);
  setCellValue(worksheet, totalRow, 16, sums.p);
  setCellValue(worksheet, totalRow, 17, sums.q);
  setCellValue(worksheet, totalRow, 18, sums.r);
  setCellValue(worksheet, totalRow, 19, sums.s);
  setCellValue(worksheet, totalRow, 20, -1);
}

function replaceAgeSheet(workbook, ageCsv) {
  if (!ageCsv) {
    return;
  }

  let worksheet = workbook.getWorksheet("반품연령집계");

  if (!worksheet) {
    worksheet = workbook.addWorksheet("반품연령집계");
  }

  const rows = csvToRows(ageCsv);

  if (rows.length === 0) {
    return;
  }

  const maxOldRows = worksheet.rowCount;
  const maxOldCols = worksheet.columnCount || 20;

  for (let rowNo = 1; rowNo <= maxOldRows; rowNo++) {
    for (let colNo = 1; colNo <= maxOldCols; colNo++) {
      worksheet.getRow(rowNo).getCell(colNo).value = "";
    }
  }

  rows.forEach((rowValues, rowIndex) => {
    const rowNo = rowIndex + 1;

    rowValues.forEach((value, colIndex) => {
      worksheet.getRow(rowNo).getCell(colIndex + 1).value = toCellValue(value);
    });
  });
}

app.post("/api/return-liability/export", async (req, res) => {
  try {
    const {
      mode,
      close_date,
      template_xlsx_base64,
      detail_csv,
      age_csv
    } = req.body;

    if (!template_xlsx_base64) {
      return res.status(400).json({
        status: "error",
        message: "template_xlsx_base64가 없습니다."
      });
    }

    if (!detail_csv) {
      return res.status(400).json({
        status: "error",
        message: "detail_csv가 없습니다."
      });
    }

    if (!["update", "rollover"].includes(mode)) {
      return res.status(400).json({
        status: "error",
        message: "mode는 update 또는 rollover여야 합니다."
      });
    }

    const workbook = new ExcelJS.Workbook();
    const templateBuffer = Buffer.from(cleanBase64(template_xlsx_base64), "base64");

    await workbook.xlsx.load(templateBuffer);

    const liabilitySheet = workbook.getWorksheet("반품충당부채");

    if (!liabilitySheet) {
      throw new Error("반품충당부채 시트를 찾지 못했습니다.");
    }

    const blocks = findLiabilityBlocks(liabilitySheet);

    const block1Snapshot = readBlockSnapshot(liabilitySheet, blocks[0]);
    const block2Snapshot = readBlockSnapshot(liabilitySheet, blocks[1]);
    const block3Snapshot = readBlockSnapshot(liabilitySheet, blocks[2]);

    if (mode === "update") {
      writeBlockSnapshot(liabilitySheet, blocks[0], block1Snapshot);
      writeBlockSnapshot(liabilitySheet, blocks[1], block2Snapshot);
    }

    if (mode === "rollover") {
      writeBlockSnapshot(liabilitySheet, blocks[0], block2Snapshot);
      writeBlockSnapshot(liabilitySheet, blocks[1], block3Snapshot);
    }

    const detailObjects = csvToObjects(detail_csv);

    const priorBlockForNewCalculation = blocks[1];

    writeCalculatedRowsToBlock(
      liabilitySheet,
      blocks[2],
      detailObjects,
      priorBlockForNewCalculation
    );

    replaceAgeSheet(workbook, age_csv);

    workbook.calcProperties.fullCalcOnLoad = true;

    const outputBuffer = await workbook.xlsx.writeBuffer();

    const id =
      Date.now().toString() +
      "_" +
      Math.random().toString(36).slice(2);

    const safeDate = String(close_date || "")
      .replace(/\./g, "")
      .replace(/-/g, "")
      .replace(/\s/g, "");

    const filename = `반품충당부채_${safeDate || "result"}.xlsx`;

    fileStore.set(id, {
      buffer: Buffer.from(outputBuffer),
      filename,
      createdAt: Date.now()
    });

    const baseUrl = `https://${req.get("host")}`;
    const downloadUrl = `${baseUrl}/downloads/${id}`;

    return res.json({
      status: "success",
      message: "재투입용 XLSX 생성 완료",
      filename,
      download_url: downloadUrl
    });

  } catch (error) {
    return res.status(500).json({
      status: "error",
      message: "XLSX 생성 중 오류가 발생했습니다.",
      detail: error.message
    });
  }
});

app.get("/downloads/:id", (req, res) => {
  const item = fileStore.get(req.params.id);

  if (!item) {
    return res.status(404).send("파일을 찾을 수 없습니다. 다시 생성해 주세요.");
  }

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  res.setHeader(
    "Content-Disposition",
    `attachment; filename*=UTF-8''${encodeURIComponent(item.filename)}`
  );

  res.send(item.buffer);
});

app.listen(PORT, () => {
  console.log(`Return Liability XLSX API running on port ${PORT}`);
});
