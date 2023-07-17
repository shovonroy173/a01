// import npm packages
const express = require("express");
const mongoose = require("mongoose");
const path = require("path");
const XLSX = require("xlsx");
const multer = require("multer");
const cors = require("cors");
require("dotenv").config();

// import local mudule
const { postXl } = require("./controller/postXlFile");

//connect to db
mongoose
  .connect(process.env.MONGO_DB)
  .then(() => {
    console.log("connected to db");
  })
  .catch((error) => {
    console.log("error", error);
  });

//init app
const app = express();

//fetch data from the request
app.use(express.json());

// cors
app.use(cors({ origin: "*" }));

const upload = multer({ dest: "uploads/" });
//import from controller
app.post("/api/submit", upload.single("file"), postXl);

//assign port
const port = process.env.PORT || 5000;
app.listen(port, () => console.log("server run at " + port));
