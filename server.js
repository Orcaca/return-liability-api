const express = require("express");
const cors = require("cors");
const axios = require("axios");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

app.get("/", (req, res) => {
  res.send("반품충당부채 API 서버가 실행 중입니다.");
});

async function testDownload(fileUrl, label) {
  const startedAt = Date.now();

  try {
    const response = await axios.get(fileUrl, {
      responseType: "arraybuffer",
      timeout: 5000,
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
      validateStatus: () => true
    });

    return {
      label,
      ok: response.status >= 200 && response.status < 300,
      http_status: response.status,
      content_type: response.headers["content-type"] || "",
      content_length: response.headers["content-length"] || "",
      downloaded_bytes: response.data ? response.data.byteLength : 0,
      elapsed_ms: Date.now() - startedAt
    };
  } catch (error) {
    return {
      label,
      ok: false,
      error: error.message,
      elapsed_ms: Date.now() - startedAt
    };
  }
}

app.post("/api/return-liability/run", async (req, res) => {
  try {
    const {
      mode,
      close_date,
      prev_file_url,
      age_file_url
    } = req.body;

    const prevFile = await testDownload(prev_file_url, "직전월 결과파일");
    const ageFile = await testDownload(age_file_url, "ERP 반품연령집계 파일");

    return res.json({
      status: "success",
      message: "파일 URL 다운로드 가능 여부 테스트 완료",
      mode,
      close_date,
      prev_file_download: prevFile,
      age_file_download: ageFile,
      logs: [
        "Dify에서 Render API로 파일 URL 전달 성공",
        "Render 서버에서 각 파일 다운로드를 최대 5초씩 시도했습니다.",
        "ok가 true이면 다운로드 가능, false이면 다운로드 불가입니다.",
        "아직 엑셀 파싱 및 계산은 수행하지 않았습니다."
      ]
    });

  } catch (error) {
    return res.status(500).json({
      status: "error",
      message: "파일 다운로드 테스트 중 서버 오류가 발생했습니다.",
      detail: error.message
    });
  }
});

app.listen(PORT, () => {
  console.log(`Return Liability API running on http://localhost:${PORT}`);
});
