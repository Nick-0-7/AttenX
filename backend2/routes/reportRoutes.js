
const express = require("express");
const router = express.Router();

const { generateReport } = require("../services/reportService");

router.post("/create-report", async (req, res) => {
  console.log("1. Report request received");

  try {
    console.log("2. Attainment type:", req.body.attainmentType);

    const report = await generateReport(req.body);

    console.log("3. Report generated. Size:", report.length, "bytes");

    const fileName =
      `Attainx_${req.body.attainmentType}_Report.docx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${fileName}"`
    );

    res.send(report);
  } catch (error) {
    console.error("Report generation error:", error);

    if (!res.headersSent) {
      res.status(500).json({
        message: "Failed to generate report",
        error: error.message,
      });
    }
  }
});

module.exports = router;