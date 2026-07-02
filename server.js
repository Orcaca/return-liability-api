const express = require("express");
const cors = require("cors");
const axios = require("axios");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

app.get("/", (req, res) => {
  res.send("반품충당부채 API 서버가 실행 중입니다.");
});

async function readExcelFromUrl(fileUrl, label) {
  if (!fileUrl) {
    throw new Error(`${label} URL이 없습니다.`);
  }

  const response = await axios.get(fileUrl, {
    responseType: "arraybuffer",
    timeout: 60000,
    maxContentLength: Infinity,
    maxBodyLength: Infinity
  });

  const buffer = Buffer.from(response.data);
  const workbook = XLSX.read(buffer, { type: "buffer" });

  return {
    label,
    sheet_names: workbook.SheetNames
  };
}

app.post("/api/return-liability/run", async (req, res) => {
  try {
    const {
      mode,
      close_date,
      prev_file_url,
      age_file_url
    } = req.body;

    if (!["update", "rollover"].includes(mode)) {
      return res.status(400).json({
        status: "error",
        message: "작업구분은 update 또는 rollover여야 합니다."
      });
    }

    if (!close_date) {
      return res.status(400).json({
        status: "error",
        message: "기준일자가 없습니다."
      });
    }

    if (!prev_file_url) {
      return res.status(400).json({
        status: "error",
        message: "직전월 결과파일 URL이 없습니다."
      });
    }

    if (!age_file_url) {
      return res.status(400).json({
        status: "error",
        message: "ERP 반품연령집계 파일 URL이 없습니다."
      });
    }

    const prevFile = await readExcelFromUrl(prev_file_url, "직전월 결과파일");
    const ageFile = await readExcelFromUrl(age_file_url, "ERP 반품연령집계 파일");

    const workName = mode === "update" ? "매월 갱신" : "연말 롤오버";

    return res.json({
      status: "success",
      message: `${workName} API 호출 및 엑셀 파일 읽기 성공`,
      close_date,
      prev_file_sheets: prevFile.sheet_names,
      age_file_sheets: ageFile.sheet_names,
      download_url: "",
      logs: [
        "Dify에서 파일 URL 수신 완료",
        "Render 서버에서 엑셀 파일 다운로드 완료",
        "엑셀 시트명 읽기 완료",
        "아직 실제 반품충당부채 계산 로직은 연결 전입니다."
      ]
    });

  } catch (error) {
    return res.status(500).json({
      status: "error",
      message: "엑셀 파일 읽기 중 오류가 발생했습니다.",
      detail: error.message
    });
  }
});

app.listen(PORT, () => {
  console.log(`Return Liability API running on http://localhost:${PORT}`);
});
