const express = require("express");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

// 서버가 켜졌는지 확인하는 주소
app.get("/", (req, res) => {
  res.send("반품충당부채 API 서버가 실행 중입니다.");
});

// Dify가 호출할 API
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
        message: "직전월 결과파일이 없습니다."
      });
    }

    if (!age_file_url) {
      return res.status(400).json({
        status: "error",
        message: "ERP 반품연령집계 파일이 없습니다."
      });
    }

    const workName = mode === "update" ? "매월 갱신" : "연말 롤오버";

    return res.json({
      status: "success",
      message: `${workName} API 호출 성공`,
      close_date,
      download_url: "",
      logs: [
        "Dify에서 값 수신 완료",
        "아직 실제 엑셀 계산 로직은 연결 전입니다."
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