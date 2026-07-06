const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "20mb" }));

const fileStore = new Map();

app.get("/", (req, res) => {
  res.send("반품충당부채 재투입용 XLSX 생성 API 서버가 실행 중입니다. total-v1");
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

function roundWon(value) {
  return Math.round(toNumber(value));
}

function normalizeDateText(value) {
  return String(value || "")
    .trim()
    .replace(/-/g, ".");
}

function previousMonthEndText(closeDateText) {
  const text = normalizeDateText(closeDateText);
  const match = /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/.exec(text);

  if (!match) {
    return "";
  }

  const year = Number(match[1]);
  const month = Number(match[2]);

  const date = new Date(year, month - 1, 0);

  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");

  return `${yyyy}.${mm}.${dd}`;
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

  const rawNumber = text.replace(/,/g, "");
  const num = Number(rawNumber);

  if (Number.isFinite(num) && /^-?\d+(\.\d+)?(e-?\d+)?$/i.test(rawNumber)) {
    return num;
  }

  return text;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function splitRowsAndTotal(rows) {
  const output = [];
  let totalRow = null;

  if (!Array.isArray(rows)) {
    return {
      rows: [],
      totalRow: null
    };
  }

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
      totalRow = out;
      return;
    }

    output.push(out);
  });

  return {
    rows: output,
    totalRow
  };
}

function getTotalValue(totalRow, colIndex, fallbackValue) {
  if (Array.isArray(totalRow)) {
    const value = toNumber(totalRow[colIndex]);

    if (value !== 0 || !isBlank(totalRow[colIndex])) {
      return value;
    }
  }

  return fallbackValue;
}

function averageNumbers(values) {
  const numbers = values.filter((value) => {
    return typeof value === "number" && Number.isFinite(value);
  });

  if (numbers.length === 0) {
    return 0;
  }

  return numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
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
    toNumber(item["적용1년반품율"]),
    toNumber(item["적용2년반품율"]),
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

function makeTotalRow(rows, block1Rows, block2Rows, block1TotalRow, block2TotalRow) {
  const total = new Array(20).fill("");

  for (let col = 2; col <= 18; col++) {
    total[col] = sumColumn(rows, col);
  }

  const netSalesTotal = total[7];
  const costTotal = total[8];

  const block1SalesTotal = getTotalValue(
    block1TotalRow,
    2,
    sumColumn(block1Rows || [], 2)
  );

  const block2SalesTotal = getTotalValue(
    block2TotalRow,
    2,
    sumColumn(block2Rows || [], 2)
  );

  total[9] = netSalesTotal !== 0 ? costTotal / netSalesTotal : 0;

  // M총계 = -F총계 / 표2 C총계
  total[12] = block2SalesTotal !== 0 ? -total[5] / block2SalesTotal : 0;

  // N총계 = -G총계 / 표1 C총계
  total[13] = block1SalesTotal !== 0 ? -total[6] / block1SalesTotal : 0;

  // K총계 = AVERAGE(M현재, M표2, M표1)
  total[10] = averageNumbers([
    total[12],
    getTotalValue(block2TotalRow, 12, null),
    getTotalValue(block1TotalRow, 12, null)
  ]);

  // L총계 = AVERAGE(N현재, N표2, N표1)
  total[11] = averageNumbers([
    total[13],
    getTotalValue(block2TotalRow, 13, null),
    getTotalValue(block1TotalRow, 13, null)
  ]);

  // 수기파일 방식: O~S는 ROUND(SUM(...), 0)
  for (let col = 14; col <= 18; col++) {
    total[col] = Math.round(total[col]);
  }

  return total;
}
function addBlock(aoa, title, rows, options = {}) {
  const {
    block1Rows = [],
    block2Rows = [],
    block1TotalRow = null,
    block2TotalRow = null,
    totalRowOverride = null
  } = options;

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

  if (Array.isArray(totalRowOverride)) {
    aoa.push(totalRowOverride);
  } else {
    aoa.push(
      makeTotalRow(
        rows,
        block1Rows,
        block2Rows,
        block1TotalRow,
        block2TotalRow
      )
    );
  }
}

function buildJournalLines(liabilityAmount, recoveryAmount) {
  const liability = roundWon(liabilityAmount);
  const recovery = roundWon(recoveryAmount);

  return [
    ["", "", "반품충당부채", liability],
    ["", "", "제품매출", -liability],
    ["제품타계정대체", -recovery, "", ""],
    ["반환제품회수권", recovery, "", ""]
  ];
}

function putJournalArea(aoa, journal, closeDate, journalMemoText) {
  while (aoa.length < 85) {
    aoa.push([]);
  }

  const cumulativeLiability = toNumber(journal?.cumulative_liability_change);
  const cumulativeRecovery = toNumber(journal?.cumulative_recovery_change);
  const monthlyLiability = toNumber(journal?.monthly_liability_change);
  const monthlyRecovery = toNumber(journal?.monthly_recovery_change);

  const cumulativeLines = buildJournalLines(cumulativeLiability, cumulativeRecovery);
  const monthlyLines = buildJournalLines(monthlyLiability, monthlyRecovery);

  const memoText =
    journalMemoText || " /재경팀/월말 반품충당부채 설정/";

  const titleRow = new Array(20).fill("");

  titleRow[11] = previousMonthEndText(closeDate);
  titleRow[12] = "적요";
  titleRow[13] = memoText;

  titleRow[16] = normalizeDateText(closeDate);
  titleRow[17] = "적요";
  titleRow[18] = memoText;

  aoa.push(titleRow);

  for (let i = 0; i < 4; i++) {
    const row = new Array(20).fill("");

    const left = cumulativeLines[i] || ["", "", "", ""];
    const right = monthlyLines[i] || ["", "", "", ""];

    // 전월말 누적분개
    // L = 차변 계정, M = 대변 계정, N = 차변 금액, O = 대변 금액
    row[11] = left[0];
    row[12] = left[2];
    row[13] = left[1];
    row[14] = left[3];

    // 당월말 입력분개
    // Q = 차변 계정, R = 대변 계정, S = 차변 금액, T = 대변 금액
    row[16] = right[0];
    row[17] = right[2];
    row[18] = right[1];
    row[19] = right[3];

    aoa.push(row);
  }
}

function putTable3CheckFormulas(sheet) {
  for (let row = 63; row <= 83; row++) {
    sheet[`T${row}`] = {
      t: "n",
      f: `+(O${row}+P${row})-(Q${row}+R${row}+S${row})`,
      v: 0
    };
  }
}

function getSheetNumber(sheet, col, row) {
  const cell = sheet[`${col}${row}`];

  if (!cell) {
    return 0;
  }

  return toNumber(cell.v);
}

function setFormulaCell(sheet, col, row, formula, cachedValue) {
  const value = Number.isFinite(cachedValue) ? cachedValue : 0;

  sheet[`${col}${row}`] = {
    t: "n",
    f: formula,
    v: value
  };
}

function sumComputedRows(rows, key) {
  return rows.reduce((sum, row) => {
    return sum + toNumber(row[key]);
  }, 0);
}

function putTable3Formulas(sheet, returnRateAssumption) {
  const rateCell = "'반품율'!$C$3";
  const assumption = toNumber(returnRateAssumption) || 2;
  const computedRows = [];

  for (let row = 63; row <= 82; row++) {
    const table2Row = row - 27; // 63 -> 36
    const table1Row = row - 54; // 63 -> 9

    const c = getSheetNumber(sheet, "C", row);
    const f = getSheetNumber(sheet, "F", row);
    const g = getSheetNumber(sheet, "G", row);
    const h = getSheetNumber(sheet, "H", row);
    const i = getSheetNumber(sheet, "I", row);

    const table2C = getSheetNumber(sheet, "C", table2Row);
    const table1C = getSheetNumber(sheet, "C", table1Row);
    const table2J = getSheetNumber(sheet, "J", table2Row);

    const table2M = getSheetNumber(sheet, "M", table2Row);
    const table1M = getSheetNumber(sheet, "M", table1Row);
    const table2N = getSheetNumber(sheet, "N", table2Row);
    const table1N = getSheetNumber(sheet, "N", table1Row);

    const j = h !== 0 ? i / h : 0;
    const m = table2C !== 0 ? -f / table2C : 0;
    const n = table1C !== 0 ? -g / table1C : 0;

    let k = 0;
    let l = 0;

    if (assumption === 1) {
      k = m;
      l = n;
    } else if (assumption === 2) {
      k = averageNumbers([m, table2M, table1M]);
      l = averageNumbers([n, table2N, table1N]);
    } else {
      k = averageNumbers([m, table2M]);
      l = averageNumbers([n, table2N]);
    }

    const o = Math.abs(j) > 1 ? 0 : c * (assumption === 1 ? m + n : k + l);
    const p = Math.abs(j) > 1 ? 0 : table2C * (assumption === 1 ? n : l);
    const q = o * (j > 1 ? 1 : j);
    const r = p * (table2J > 1 ? 1 : table2J);
    const s = o + p - q - r;
    const t = (o + p) - (q + r + s);

    computedRows.push({ j, k, l, m, n, o, p, q, r, s, t });

    setFormulaCell(sheet, "J", row, `IFERROR(I${row}/H${row},0)`, j);

    setFormulaCell(
      sheet,
      "K",
      row,
      `CHOOSE(${rateCell},M${row},AVERAGE(M${row},M${table2Row},M${table1Row}),AVERAGE(M${row},M${table2Row}))`,
      k
    );

    setFormulaCell(
      sheet,
      "L",
      row,
      `CHOOSE(${rateCell},N${row},AVERAGE(N${row},N${table2Row},N${table1Row}),AVERAGE(N${row},N${table2Row}))`,
      l
    );

    setFormulaCell(sheet, "M", row, `IFERROR(-F${row}/C${table2Row},0)`, m);
    setFormulaCell(sheet, "N", row, `IFERROR(-G${row}/C${table1Row},0)`, n);

    setFormulaCell(
      sheet,
      "O",
      row,
      `IF(ABS(J${row})>1,0,C${row}*CHOOSE(${rateCell},SUM(M${row}:N${row}),SUM(K${row}:L${row}),SUM(K${row}:L${row})))`,
      o
    );

    setFormulaCell(
      sheet,
      "P",
      row,
      `IF(ABS(J${row})>1,0,C${table2Row}*CHOOSE(${rateCell},N${row},L${row},L${row}))`,
      p
    );

    setFormulaCell(sheet, "Q", row, `O${row}*IF(J${row}>1,1,J${row})`, q);
    setFormulaCell(sheet, "R", row, `P${row}*IF(J${table2Row}>1,1,J${table2Row})`, r);
    setFormulaCell(sheet, "S", row, `O${row}+P${row}-Q${row}-R${row}`, s);
    setFormulaCell(sheet, "T", row, `+(O${row}+P${row})-(Q${row}+R${row}+S${row})`, t);
  }

  const j83 = getSheetNumber(sheet, "H", 83) !== 0
    ? getSheetNumber(sheet, "I", 83) / getSheetNumber(sheet, "H", 83)
    : 0;

  const m83 = getSheetNumber(sheet, "C", 56) !== 0
    ? -getSheetNumber(sheet, "F", 83) / getSheetNumber(sheet, "C", 56)
    : 0;

  const n83 = getSheetNumber(sheet, "C", 29) !== 0
    ? -getSheetNumber(sheet, "G", 83) / getSheetNumber(sheet, "C", 29)
    : 0;

  let k83 = 0;
  let l83 = 0;

  if (assumption === 1) {
    k83 = m83;
    l83 = n83;
  } else if (assumption === 2) {
    k83 = averageNumbers([
      m83,
      getSheetNumber(sheet, "M", 56),
      getSheetNumber(sheet, "M", 29)
    ]);

    l83 = averageNumbers([
      n83,
      getSheetNumber(sheet, "N", 56),
      getSheetNumber(sheet, "N", 29)
    ]);
  } else {
    k83 = averageNumbers([
      m83,
      getSheetNumber(sheet, "M", 56)
    ]);

    l83 = averageNumbers([
      n83,
      getSheetNumber(sheet, "N", 56)
    ]);
  }

  const o83 = Math.round(sumComputedRows(computedRows, "o"));
  const p83 = Math.round(sumComputedRows(computedRows, "p"));
  const q83 = Math.round(sumComputedRows(computedRows, "q"));
  const r83 = Math.round(sumComputedRows(computedRows, "r"));
  const s83 = Math.round(sumComputedRows(computedRows, "s"));
  const t83 = (o83 + p83) - (q83 + r83 + s83);

  setFormulaCell(sheet, "J", 83, "IFERROR(I83/H83,0)", j83);

  setFormulaCell(
    sheet,
    "K",
    83,
    `CHOOSE(${rateCell},M83,AVERAGE(M83,M56,M29),AVERAGE(M83,M56))`,
    k83
  );

  setFormulaCell(
    sheet,
    "L",
    83,
    `CHOOSE(${rateCell},N83,AVERAGE(N83,N56,N29),AVERAGE(N83,N56))`,
    l83
  );

  setFormulaCell(sheet, "M", 83, "IFERROR(-F83/C56,0)", m83);
  setFormulaCell(sheet, "N", 83, "IFERROR(-G83/C29,0)", n83);

  setFormulaCell(sheet, "O", 83, "ROUND(SUM(O63:O82),0)", o83);
  setFormulaCell(sheet, "P", 83, "ROUND(SUM(P63:P82),0)", p83);
  setFormulaCell(sheet, "Q", 83, "ROUND(SUM(Q63:Q82),0)", q83);
  setFormulaCell(sheet, "R", 83, "ROUND(SUM(R63:R82),0)", r83);
  setFormulaCell(sheet, "S", 83, "ROUND(SUM(S63:S82),0)", s83);
  setFormulaCell(sheet, "T", 83, "+(O83+P83)-(Q83+R83+S83)", t83);
}

function setJournalCell(sheet, address, value) {
  if (value === "" || value === null || value === undefined) {
    delete sheet[address];
    return;
  }

  if (typeof value === "number") {
    sheet[address] = {
      t: "n",
      v: value
    };
    return;
  }

  sheet[address] = {
    t: "s",
    v: String(value)
  };
}

function setJournalCell(sheet, address, value) {
  if (value === "" || value === null || value === undefined) {
    delete sheet[address];
    return;
  }

  if (typeof value === "number") {
    sheet[address] = {
      t: "n",
      v: value
    };
    return;
  }

  sheet[address] = {
    t: "s",
    v: String(value)
  };
}

function overwriteRolloverJournalFromSheet(sheet, mode) {
  const modeKey = String(mode || "").trim().toLowerCase();

  if (modeKey !== "rollover") {
    return;
  }

  const table3Liability = getSheetNumber(sheet, "O", 83) + getSheetNumber(sheet, "P", 83);
  const table2Liability = getSheetNumber(sheet, "O", 56) + getSheetNumber(sheet, "P", 56);

  const table3Recovery = getSheetNumber(sheet, "Q", 83) + getSheetNumber(sheet, "R", 83);
  const table2Recovery = getSheetNumber(sheet, "Q", 56) + getSheetNumber(sheet, "R", 56);

  const monthlyLiability = roundWon(table3Liability - table2Liability);
  const monthlyRecovery = roundWon(table3Recovery - table2Recovery);

  // 전월누적분개는 rollover에서 0
  setJournalCell(sheet, "O87", 0);
  setJournalCell(sheet, "O88", 0);
  setJournalCell(sheet, "N89", 0);
  setJournalCell(sheet, "N90", 0);

  // 당월분개는 표3 - 표2 증감분
  setJournalCell(sheet, "T87", monthlyLiability);
  setJournalCell(sheet, "T88", -monthlyLiability);

  setJournalCell(sheet, "S89", -monthlyRecovery);
  setJournalCell(sheet, "S90", monthlyRecovery);
}

function formatWon(value) {
  return `${roundWon(value).toLocaleString("ko-KR")}원`;
}

function formatSignedWon(value) {
  const n = roundWon(value);

  if (n > 0) {
    return `+${n.toLocaleString("ko-KR")}원`;
  }

  if (n < 0) {
    return `${n.toLocaleString("ko-KR")}원`;
  }

  return "0원";
}

function getSummaryLine(summaryText, prefix) {
  const lines = String(summaryText || "").split(/\r?\n/);

  for (const line of lines) {
    if (line.trim().startsWith(prefix)) {
      return line.trim();
    }
  }

  return "";
}

function modeLabelText(mode) {
  const modeKey = String(mode || "").trim().toLowerCase();

  if (modeKey === "rollover") {
    return "연말 롤오버";
  }

  return "월별 업데이트";
}

function buildFinalSummaryText(sheet, mode, closeDate, originalSummaryText) {
  const liability =
    roundWon(getSheetNumber(sheet, "O", 83) + getSheetNumber(sheet, "P", 83));

  const recovery =
    roundWon(getSheetNumber(sheet, "Q", 83) + getSheetNumber(sheet, "R", 83));

  const netProvision = roundWon(getSheetNumber(sheet, "S", 83));

  // 최종 파일의 분개박스 기준
  const monthlyLiability = roundWon(getSheetNumber(sheet, "T", 87));
  const monthlyRecovery = roundWon(getSheetNumber(sheet, "S", 90));

  const debitTransfer = roundWon(getSheetNumber(sheet, "S", 89));
  const debitRecovery = roundWon(getSheetNumber(sheet, "S", 90));

  const creditLiability = roundWon(getSheetNumber(sheet, "T", 87));
  const creditSales = roundWon(getSheetNumber(sheet, "T", 88));

  const erpLine = getSummaryLine(originalSummaryText, "ERP 매칭:");
  const unmatchedLine = getSummaryLine(originalSummaryText, "미매칭 ERP 항목:");

  const lines = [];

  lines.push("반품충당부채 계산 완료");
  lines.push("");
  lines.push(`작업구분: ${modeLabelText(mode)}`);
  lines.push(`기준일자: ${normalizeDateText(closeDate)}`);
  lines.push("");
  lines.push(`반품충당부채: ${formatWon(liability)}`);
  lines.push(`전기 대비 증감: ${formatSignedWon(monthlyLiability)}`);
  lines.push("");
  lines.push(`반환제품회수권: ${formatWon(recovery)}`);
  lines.push(`전기 대비 증감: ${formatSignedWon(monthlyRecovery)}`);
  lines.push("");
  lines.push(`순 충당부채: ${formatWon(netProvision)}`);
  lines.push("");
  lines.push("결산분개 미리보기");
  lines.push("");
  lines.push("차변");
  lines.push(`제품타계정대체 ${formatWon(debitTransfer)}`);
  lines.push(`반환제품회수권 ${formatWon(debitRecovery)}`);
  lines.push("");
  lines.push("대변");
  lines.push(`반품충당부채 ${formatWon(creditLiability)}`);
  lines.push(`제품매출 ${formatWon(creditSales)}`);

  if (erpLine || unmatchedLine) {
    lines.push("");

    if (erpLine) {
      lines.push(erpLine);
    }

    if (unmatchedLine) {
      lines.push(unmatchedLine);
    }
  }

  return lines.join("\n");
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
      journal,
      journal_memo_text
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

    const oldBlock1Split = splitRowsAndTotal(block1_rows);
    const oldBlock2Split = splitRowsAndTotal(block2_rows);
    
    const oldBlock1 = oldBlock1Split.rows;
    const oldBlock2 = oldBlock2Split.rows;

    const oldBlock1TotalRow = oldBlock1Split.totalRow;
    const oldBlock2TotalRow = oldBlock2Split.totalRow;

    const detailObjects = csvToObjects(detail_csv);
    const newBlock = detailObjects.map(detailObjectToRow);

    for (let i = 0; i < 6; i++) {
      liabilityAoa.push([]);
    }

    addBlock(liabilityAoa, "표 1", oldBlock1, {
      totalRowOverride: oldBlock1TotalRow,
      block1Rows: oldBlock1,
      block2Rows: oldBlock2,
      block1TotalRow: oldBlock1TotalRow,
      block2TotalRow: oldBlock2TotalRow
    });

    for (let i = 0; i < 4; i++) {
      liabilityAoa.push([]);
    }

    addBlock(liabilityAoa, "표 2", oldBlock2, {
      totalRowOverride: oldBlock2TotalRow,
      block1Rows: oldBlock1,
      block2Rows: oldBlock2,
      block1TotalRow: oldBlock1TotalRow,
      block2TotalRow: oldBlock2TotalRow
    });

    for (let i = 0; i < 4; i++) {
      liabilityAoa.push([]);
    }

    addBlock(liabilityAoa, "표 3", newBlock, {
      block1Rows: oldBlock1,
      block2Rows: oldBlock2,
      block1TotalRow: oldBlock1TotalRow,
      block2TotalRow: oldBlock2TotalRow
    });

    putJournalArea(liabilityAoa, journal || {}, close_date, journal_memo_text);

    const liabilitySheet = XLSX.utils.aoa_to_sheet(liabilityAoa);

    putTable3Formulas(liabilitySheet, return_rate_assumption);

    overwriteRolloverJournalFromSheet(liabilitySheet, mode);

    const finalSummaryText = buildFinalSummaryText(
      liabilitySheet,
      mode,
      close_date,
      summary_text
    );

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
      ["", "반품 1년차 비율(10기준)", toNumber(scenario_factor)],
      ["", "반품율 가정", toNumber(return_rate_assumption)]
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

    workbook.Workbook = workbook.Workbook || {};
    workbook.Workbook.CalcPr = {
      calcMode: "auto"
    };

    const buffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx"
    });

    const id = Date.now().toString() + "_" + Math.random().toString(36).slice(2);
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
      download_url: downloadUrl,
      final_summary_text: finalSummaryText
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
