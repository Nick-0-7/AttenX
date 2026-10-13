const express = require("express");
const { createFrontendReport } = require("../controllers/report.controller");

const router = express.Router();

router.post("/create-report", createFrontendReport);

module.exports = router;
