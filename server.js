const express = require("express");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

app.get("/", (req, res) => {
  res.send("반품충당부채 API 서버가 실행 중입니다.");
});

app.post("/api/return-liability/run", async (req, res) => {
  try {
    const {
      mode,
      close_date,
      prev_file_url,
      age_file_url
    } = req.body;

    return res.json({
      status: "success",
      message: "Dify에서 Render API까지 값 전달 성공",
      received: {
        mode,
        close_date,
        prev_file_url_exists: !!prev_file_url,
        age_file_url_exists: !!age_file_url,
        prev_file_url_preview: prev_file_url ? String(prev_file_url).slice(0, 80) : "",
        age_file_url_preview: age_file_url ? String(age_file_url).slice(0, 80) : ""
      },
      logs: [
        "Dify에서 Render API 호출 성공",
        "파일 URL 수신 여부 확인 완료",
        "아직 엑셀 다운로드/읽기는 수행하지 않았습니다."
      ]
    });

  } catch (error) {
    return res.status(500).json({
      status: "error",
      message: error.message
    });
  }
});

app.listen(PORT, () => {
  console.log(`Return Liability API running on http://localhost:${PORT}`);
});
