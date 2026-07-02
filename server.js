const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024
  }
});

app.use(cors());
app.use(express.json({ limit: "50mb" }));

app.get("/", (req, res) => {
  res.send("반품충당부채 API 서버가 실행 중입니다.");
});

function readWorkbookFromUploadedFile(file, label) {
  if (!file) {
    throw new Error(`${label} 파일이 업로드되지 않았습니다.`);
  }

  const workbook = XLSX.read(file.buffer, { type: "buffer" });

  return {
    label,
    original_name: file.originalname,
    size: file.size,
    sheet_names: workbook.SheetNames
  };
}

app.post(
  "/api/return-liability/run",
  upload.fields([
    { name: "prev_file", maxCount: 1 },
    { name: "age_file", maxCount: 1 }
  ]),
  async (req, res) => {
    try {
      const mode = req.body.mode;
      const close_date = req.body.close_date;

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

      const prevFile = req.files?.prev_file?.[0];
      const ageFile = req.files?.age_file?.[0];

      const prevWorkbook = readWorkbookFromUploadedFile(prevFile, "직전월 결과파일");
      const ageWorkbook = readWorkbookFromUploadedFile(ageFile, "ERP 반품연령집계 파일");

      const workName = mode === "update" ? "매월 갱신" : "연말 롤오버";

      return res.json({
        status: "success",
        message: `${workName} 파일 업로드 및 엑셀 시트 읽기 성공`,
        mode,
        close_date,
        prev_file: prevWorkbook,
        age_file: ageWorkbook,
        download_url: "",
        logs: [
          "Dify에서 Render API로 파일 직접 업로드 성공",
          "Render 서버에서 업로드된 엑셀 파일 수신 완료",
          "엑셀 시트명 읽기 완료",
          "아직 실제 반품충당부채 계산 로직은 연결 전입니다."
        ]
      });

    } catch (error) {
      return res.status(500).json({
        status: "error",
        message: "파일 업로드 또는 엑셀 읽기 중 오류가 발생했습니다.",
        detail: error.message
      });
    }
  }
);

app.listen(PORT, () => {
  console.log(`Return Liability API running on http://localhost:${PORT}`);
});
