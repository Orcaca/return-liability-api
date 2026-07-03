const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "20mb" }));

const fileStore = new Map();

app.get("/", (req, res) => {
  res.send("반품충당부채 재투입용 XLSX 생성 API 서버가 실행 중입니다.");
});

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

  return Number.isFinite(num) ? num : 0;
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

  if (
    Number.isFinite(num) &&
    /^-?\d+(\.\d+)?(e-?\d+)?$/i.test(text.replace(/,/g, ""))
  ) {
    return num;
  }

  return text;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function normalizeRawRows(rows) {
  if (!Array.isArray(rows)) {
    return [];
  }

  const output = [];

  rows.forEach((row) => {
    const out = [];

    for (let i = 0; i < 20; i++) {
      out.push(toCellValue(Array.isArray(row) ? row[i] : ""));
    }

    const hasAnyValue = out.some((v) => !isBlank(v));

    if (!hasAnyValue) {
      return;
    }

    const isTotalRow =
      isBlank(out[0]) &&
      isBlank(out[1]) &&
      out.slice(2).some((v) => !isBlank(v));

    if (isTotalRow) {
      return;
    }

    output.push(out);
  });

  return output;
}

function detailObjectToRow(item) {
  return [
    item["대분류"] || "",
    item["구분"] || "",
    toNumber(item["총액매출"]),
    toNumber(item["당해매출반품"]),
    toNumber(item["1년이상반품"]),
    toNumber(item["1년"]),
    toNumber(item["2년"]),
    toNumber(item["순매출액"]),
    toNumber(item["원가금액"]),
    toNumber(item["원가율"]),

    // K/L = 적용 반품율
    toNumber(item["적용1년반품율"]),
    toNumber(item["적용2년반품율"]),

    // M/N = 당해 반품율
    toNumber(item["당해1년반품율"]),
    toNumber(item["당해2년반품율"]),

    toNumber(item["당해매출기준_반품추정액"]),
    toNumber(item["전기매출기준_반품추정액"]),
    toNumber(item["당해매출기준_원가추정액"]),
    toNumber(item["전기매출기준_원가추정액"]),
    toNumber(item["순충당부채"]),
    ""
  ];
}

function sumColumn(rows, colIndex) {
  return rows.reduce((sum, row) => {
    const value = row[colIndex];
    return sum + (typeof value === "number" ? value : toNumber(value));
  }, 0);
}

function makeTotalRow(rows, block1Rows, block2Rows) {
  const total = new Array(20).fill("");

  for (let col = 2; col <= 18; col++) {
    total[col] = sumColumn(rows, col);
  }

  const currentSalesTotal = total[2];
  const netSalesTotal = total[7];
  const costTotal = total[8];

  const oneYearReturnTotal = total[5];
  const twoYearReturnTotal = total[6];

  const currentSalesReturnEstimate = total[14];
  const priorSalesReturnEstimate = total[15];

  const block1SalesTotal = sumColumn(block1Rows || [], 2);
  const block2SalesTotal = sumColumn(block2Rows || [], 2);

  // J열 원가율 = I / H
  total[9] = netSalesTotal !== 0 ? costTotal / netSalesTotal : 0;

  // K/L열 적용 반품율
  total[11] = block2SalesTotal !== 0 ? priorSalesReturnEstimate / block2SalesTotal : 0;
  total[10] =
    currentSalesTotal !== 0
      ? currentSalesReturnEstimate / currentSalesTotal - total[11]
      : 0;

  // M/N열 당해 반품율
  total[12] = block2SalesTotal !== 0 ? -oneYearReturnTotal / block2SalesTotal : 0;
  total[13] = block1SalesTotal !== 0 ? -twoYearReturnTotal / block1SalesTotal : 0;

  return total;
}

function addBlock(aoa, title, rows, block1Rows, block2Rows) {
  const header = [
    "대분류",
    "구분",
    "총액매출",
    "당해매출반품",
    "1년이상 반품",
    "1년",
    "2년",
    "순 공급가액",
    "원가금액",
    "원가율",
    "적용 1년 반품율",
    "적용 2년 반품율",
    "당해 1년 반품율",
    "당해 2년 반품율",
    "당해매출기준 반품추정액",
    "전기매출기준 반품추정액",
    "당해매출기준 원가추정액",
    "전기매출기준 원가추정액",
    "순 충당부채",
    "체크"
  ];

  aoa.push([title]);
  aoa.push(header);

  rows.forEach((row) => {
    aoa.push(row);
  });

  aoa.push(makeTotalRow(rows, block1Rows, block2Rows));
}

function buildJournalLines(liabilityAmount, recoveryAmount) {
  const lines = [];

  if (liabilityAmount >= 0) {
    lines.push([
      "매출액",
      Math.abs(liabilityAmount),
      "반품충당부채",
      Math.abs(liabilityAmount)
    ]);
  } else {
    lines.push([
      "반품충당부채",
      Math.abs(liabilityAmount),
      "매출액",
      Math.abs(liabilityAmount)
    ]);
  }

  if (recoveryAmount >= 0) {
    lines.push([
      "반환제품회수권",
      Math.abs(recoveryAmount),
      "매출원가",
      Math.abs(recoveryAmount)
    ]);
  } else {
    lines.push([
      "매출원가",
      Math.abs(recoveryAmount),
      "반환제품회수권",
      Math.abs(recoveryAmount)
    ]);
  }

  return lines;
}

function putJournalArea(aoa, journal) {
  while (aoa.length < 85) {
    aoa.push([]);
  }

  const prevLiability = toNumber(journal?.prev_liability_balance);
  const prevRecovery = toNumber(journal?.prev_recovery_balance);
  const monthlyLiability = toNumber(journal?.monthly_liability_change);
  const monthlyRecovery = toNumber(journal?.monthly_recovery_change);

  const cumulativeLines = buildJournalLines(prevLiability, prevRecovery);
  const monthlyLines = buildJournalLines(monthlyLiability, monthlyRecovery);

  for (let i = 0; i < 5; i++) {
    const row = new Array(20).fill("");

    if (i === 0) {
      row[11] = "전월말까지 누적 결산분개";
      row[16] = "당월말 입력 결산분개";
    }

    if (i === 1) {
      row[11] = "차변";
      row[12] = "금액";
      row[13] = "대변";
      row[14] = "금액";

      row[16] = "차변";
      row[17] = "금액";
      row[18] = "대변";
      row[19] = "금액";
    }

    if (i >= 2 && i <= 3) {
      const left = cumulativeLines[i - 2] || ["", "", "", ""];
      const right = monthlyLines[i - 2] || ["", "", "", ""];

      row[11] = left[0];
      row[12] = left[1];
      row[13] = left[2];
      row[14] = left[3];

      row[16] = right[0];
      row[17] = right[1];
      row[18] = right[2];
      row[19] = right[3];
    }

    aoa.push(row);
  }
}

app.post("/api/return-liability/export", async (req, res) => {
  try {
    const {
      mode,
      close_date,
      summary_text,
      block1_rows,
      block2_rows,
      detail_csv,
      age_csv,
      scenario_factor,
      return_rate_assumption,
      journal
    } = req.body;

    if (!["update", "rollover"].includes(mode)) {
      return res.status(400).json({
        status: "error",
        message: "mode는 update 또는 rollover여야 합니다."
      });
    }

    if (!detail_csv) {
      return res.status(400).json({
        status: "error",
        message: "detail_csv가 없습니다."
      });
    }

    const workbook = XLSX.utils.book_new();

    const liabilityAoa = [];

    const oldBlock1 = normalizeRawRows(block1_rows);
    const oldBlock2 = normalizeRawRows(block2_rows);

    const detailObjects = csvToObjects(detail_csv);
    const newBlock = detailObjects.map(detailObjectToRow);

    // 표1 위 빈행 6개
    for (let i = 0; i < 6; i++) {
      liabilityAoa.push([]);
    }

    addBlock(liabilityAoa, "표 1", oldBlock1, oldBlock1, oldBlock2);

    // 표1과 표2 사이 빈행 4개
    for (let i = 0; i < 4; i++) {
      liabilityAoa.push([]);
    }

    addBlock(liabilityAoa, "표 2", oldBlock2, oldBlock1, oldBlock2);

    // 표2와 표3 사이 빈행 4개
    for (let i = 0; i < 4; i++) {
      liabilityAoa.push([]);
    }

    addBlock(liabilityAoa, "표 3", newBlock, oldBlock1, oldBlock2);

    // L86:O90, Q86:T90 결산분개 영역
    putJournalArea(liabilityAoa, journal || {});

    const liabilitySheet = XLSX.utils.aoa_to_sheet(liabilityAoa);

    liabilitySheet["!cols"] = [
      { wch: 16 },
      { wch: 24 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 12 },
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
      { wch: 18 },
      { wch: 18 },
      { wch: 18 },
      { wch: 18 },
      { wch: 15 },
      { wch: 10 }
    ];

    XLSX.utils.book_append_sheet(workbook, liabilitySheet, "반품충당부채");

    const rateAoa = [
      [],
      ["", "", toNumber(scenario_factor)],
      ["", "", toNumber(return_rate_assumption)]
    ];

    const rateSheet = XLSX.utils.aoa_to_sheet(rateAoa);
    XLSX.utils.book_append_sheet(workbook, rateSheet, "반품율");

    const ageRows = csvToRows(age_csv || "");
    const ageSheet = XLSX.utils.aoa_to_sheet(ageRows);
    XLSX.utils.book_append_sheet(workbook, ageSheet, "반품연령집계");

    const summaryRows = String(summary_text || "")
      .split(/\r?\n/)
      .map((line) => [line]);

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    summarySheet["!cols"] = [{ wch: 70 }];
    XLSX.utils.book_append_sheet(workbook, summarySheet, "요약");

    const buffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx"
    });

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
      buffer,
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
