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

function normalizeRawRows(rows) {
  if (!Array.isArray(rows)) {
    return [];
  }

  return rows.map((row) => {
    const out = [];

    for (let i = 0; i < 20; i++) {
      out.push(toCellValue(Array.isArray(row) ? row[i] : ""));
    }

    return out;
  });
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
    toNumber(item["당해1년반품율"]),
    toNumber(item["당해2년반품율"]),
    toNumber(item["적용1년반품율"]),
    toNumber(item["적용2년반품율"]),
    toNumber(item["당해매출기준_반품추정액"]),
    toNumber(item["전기매출기준_반품추정액"]),
    toNumber(item["당해매출기준_원가추정액"]),
    toNumber(item["전기매출기준_원가추정액"]),
    toNumber(item["순충당부채"]),
    ""
  ];
}

function makeTotalRow(rows) {
  const total = new Array(20).fill("");

  for (let col = 2; col <= 18; col++) {
    total[col] = rows.reduce((sum, row) => {
      const value = row[col];
      return sum + (typeof value === "number" ? value : toNumber(value));
    }, 0);
  }

  if (total[7] !== 0) {
    total[9] = total[8] / total[7];
  }

  if (total[2] !== 0) {
    total[10] = -total[5] / total[2];
    total[11] = -total[6] / total[2];
    total[12] = total[14] / total[2] - total[13];
  }

  return total;
}

function addBlock(aoa, title, rows) {
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
    "당해 1년 반품율",
    "당해 2년 반품율",
    "적용 1년 반품율",
    "적용 2년 반품율",
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

  aoa.push(makeTotalRow(rows));
  aoa.push([]);
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
      return_rate_assumption
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

    addBlock(liabilityAoa, "표 1", oldBlock1);
    addBlock(liabilityAoa, "표 2", oldBlock2);
    addBlock(liabilityAoa, "표 3", newBlock);

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
