const express = require("express");
const cors = require("cors");
const reportRoutes = require("./routes/reportRoutes");

const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.json());

app.get("/", (req, res) => {
  res.send("Attainx backend server is running!");
});

app.use("/api/reports", reportRoutes);

app.listen(PORT, () => {
  console.log(`Attainx server running at http://localhost:${PORT}`);
});