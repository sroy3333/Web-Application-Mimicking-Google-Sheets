require("dotenv").config();
const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const bodyParser = require("body-parser");

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Connect to MongoDB
mongoose.connect(process.env.MONGO_URI, { useNewUrlParser: true, useUnifiedTopology: true })
  .then(() => console.log("MongoDB Connected"))
  .catch(err => console.log(err));

// Spreadsheet Schema
const SpreadsheetSchema = new mongoose.Schema({
  name: String,
  data: Array,
});
const Spreadsheet = mongoose.model("Spreadsheet", SpreadsheetSchema);

// Save Spreadsheet
app.post("/save", async (req, res) => {
  const { name, data } = req.body;
  let sheet = await Spreadsheet.findOne({ name });

  if (sheet) {
    sheet.data = data;
    await sheet.save();
  } else {
    sheet = new Spreadsheet({ name, data });
    await sheet.save();
  }
  res.json({ message: "Spreadsheet saved successfully!" });
});

// Load Spreadsheet
app.get("/load/:name", async (req, res) => {
  const sheet = await Spreadsheet.findOne({ name: req.params.name });
  res.json(sheet ? sheet.data : []);
});

// Run Server
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
