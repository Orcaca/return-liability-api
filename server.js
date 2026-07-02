const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "20mb" }));

const fileStore = new Map();

app.get("/", (req, res) => {
  res.send("반품충당부채 XLSX 생성 API 서버가 실행 중입니다.");
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

function csvToRows(csvText) {
  return String(csvText || "")
    .split(/\r?\n/)
    .filter(line => line.trim() !== "")
    .map(parseCsvLine);
}

app.post("/api/return-liability/export", async (req, res) => {
  try {
    const {
      close_date,
      summary_text,
      detail_csv
    } = req.body;

    if (!summary_text) {
      return res.status(400).json({
        status: "error",
        message: "summary_text가 없습니다."
      });
    }

    if (!detail_csv) {
      return res.status(400).json({
        status: "error",
        message: "detail_csv가 없습니다."
      });
    }

    const workbook = XLSX.utils.book_new();

    const summaryRows = String(summary_text)
      .split(/\r?\n/)
      .map(line => [line]);

    const detailRows = csvToRows(detail_csv);

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    const detailSheet = XLSX.utils.aoa_to_sheet(detailRows);

    summarySheet["!cols"] = [{ wch: 60 }];
    detailSheet["!cols"] = [
      { wch: 18 }, { wch: 28 }, { wch: 10 },
      { wch: 15 }, { wch: 15 }, { wch: 15 },
      { wch: 15 }, { wch: 15 }, { wch: 15 },
      { wch: 15 }, { wch: 12 }, { wch: 20 },
      { wch: 20 }, { wch: 20 }, { wch: 20 },
      { wch: 15 }
    ];

    XLSX.utils.book_append_sheet(workbook, summarySheet, "요약");
    XLSX.utils.book_append_sheet(workbook, detailSheet, "상세계산표");

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
      message: "XLSX 생성 완료",
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
