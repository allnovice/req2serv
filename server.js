// server.js
import 'dotenv/config'; // or: import dotenv from 'dotenv'; dotenv.config();
import express from "express";
import cors from "cors";
import fileUpload from "express-fileupload";
import fs from "fs";
import path from "path";
import { exec } from "child_process";
import XLSX from "xlsx";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import ExcelJS from "exceljs";
import axios from "axios";

const app = express();
const UPLOAD_DIR = process.env.UPLOAD_DIR || path.join(process.cwd(), "uploads");
const DB_FILE = process.env.DB_FILE || path.join(process.cwd(), "filled_forms.db");

app.use(cors());
app.use(express.json());
app.use(fileUpload());
app.use("/uploads", express.static(UPLOAD_DIR));

if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);

// Ensure SQLite DB exists and table created
exec(`sqlite3 ${DB_FILE} "CREATE TABLE IF NOT EXISTS filled_forms (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  filename TEXT,
  version TEXT,
  data TEXT,
  created_at TEXT
);"`);

// Helper to parse version from filename, e.g., sample_v1_form.xlsx -> v1
function parseVersion(filename) {
  const match = filename.match(/_v(\d+)_/i);
  return match ? `v${match[1]}` : "v1";
}

// GET all available templates (ignore -filled files)
app.get("/forms", (req, res) => {
  const files = fs.readdirSync(UPLOAD_DIR)
    .filter(f => (f.endsWith(".docx") || f.endsWith(".xlsx")) && !f.includes("-filled"));
  res.json({ forms: files });
});

// POST upload new template
app.post("/upload", (req, res) => {
  if (!req.files?.file) return res.status(400).json({ error: "No file uploaded" });
  const file = req.files.file;
  const savePath = path.join(UPLOAD_DIR, file.name);
  file.mv(savePath, err => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ message: "Uploaded", filename: file.name, url: `/uploads/${file.name}` });
  });
});

// POST fill form
app.post("/fill", async (req, res) => {
  try {
    const { filename, data } = req.body;
    if (!filename || typeof data !== "object") return res.status(400).json({ error: "Invalid input" });

    const source = path.join(UPLOAD_DIR, filename);
    if (!fs.existsSync(source)) return res.status(404).json({ error: "file not found" });

    const version = parseVersion(filename);
    let filledFilename;

    if (filename.endsWith(".xlsx")) {
      const ExcelJS = (await import("exceljs")).default;
      const axios = (await import("axios")).default;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(source);

      for (const sheet of workbook.worksheets) {
  for (const row of sheet._rows) {
    if (!row) continue;
    for (const cell of row._cells) {
      if (!cell?.value || typeof cell.value !== "string") continue;
      let value = cell.value.trim();

      // check if this cell itself is a signature placeholder
      const sigMatch = value.match(/^{{(signature\d+)}}$/i);
      if (sigMatch) {
        const key = sigMatch[1];
        if (data[key]?.startsWith("http")) {
          try {
            const response = await axios.get(data[key], { responseType: "arraybuffer" });
            const imageId = workbook.addImage({ buffer: response.data, extension: "png" });
            sheet.addImage(imageId, {
              tl: { col: cell.col - 1, row: cell.row - 1 },
              ext: { width: 120, height: 40 },
            });
            cell.value = ""; // clear placeholder text
          } catch (e) {
            console.log("image fetch error:", e.message);
          }
        }
        continue; // skip normal text replacement
      }

      // normal text replacement for non-signature placeholders
      for (const key of Object.keys(data)) {
        if (!key.toLowerCase().startsWith("signature")) {
          value = value.replaceAll(`{{${key}}}`, data[key]);
        }
      }
      cell.value = value;
    }
  }
}

      filledFilename = filename.replace(".xlsx", `-filled-${Date.now()}.xlsx`);
      await workbook.xlsx.writeFile(path.join(UPLOAD_DIR, filledFilename));
    } else {
      return res.status(400).json({ error: "only .xlsx supported for now" });
    }

    const jsonData = JSON.stringify(data).replace(/"/g, '""');
    exec(`sqlite3 ${DB_FILE} "INSERT INTO filled_forms (filename, version, data, created_at) VALUES ('${filename}', '${version}', '${jsonData}', datetime('now'));"`);

    res.json({ message: "Form filled", url: `/uploads/${filledFilename}` });
  } catch (err) {
    console.error("fill error:", err);
    res.status(500).json({ error: err.message });
  }
});

// GET filled forms metadata
app.get("/filled", (req, res) => {
  exec(`sqlite3 -json ${DB_FILE} "SELECT * FROM filled_forms ORDER BY created_at DESC;"`, (err, stdout) => {
    if (err) {
      console.error(err);
      return res.status(500).json({ error: "Failed to fetch filled forms" });
    }
    const data = stdout ? JSON.parse(stdout) : [];
    res.json({ filled: data });
  });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
