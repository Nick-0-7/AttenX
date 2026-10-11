const express = require("express");
const cors = require("cors");
const { hardcodedReportData } = require("./config/reportData");
const reportRoutes = require("./routes/reportRoutes");

const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.json());

app.get("/", (req, res) => {
  res.send("Attainx backend server is running!");
});

app.use("/api/reports", reportRoutes);

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Attainx server running at http://localhost:${PORT}`);
  });
}

module.exports = { app, hardcodedReportData };