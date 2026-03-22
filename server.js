import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import puppeteer from "puppeteer";
import archiver from "archiver";
import fs from "fs";
import path from "path";

const app = express();
const upload = multer({ dest: "uploads/" });

app.get("/", (req, res) => {
  res.send(`
    <h2>Rocketry Box POD Generator</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="excel" accept=".xlsx" required />
      <button type="submit">Upload & Generate POD ZIP</button>
    </form>
    <p>For bulk PDFs this may take time.</p>
  `);
});

app.post("/upload", upload.single("excel"), async (req, res) => {
  try {
    const filePath = req.file.path;

    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const outputDir = path.join(process.cwd(), "output");
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const browser = await puppeteer.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
      executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined
    });

    const page = await browser.newPage();

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const awb = String(row.awb || "NO_AWB").replace(/[^A-Za-z0-9]/g, "");

      const html = `
      <html><body style="font-family:Arial;padding:40px;">
      <h2 style="text-align:center;">Proof Of Delivery (E-POD)</h2>
      <p><b>Courier:</b> ${row.courier || ""}</p>
      <p><b>AWB:</b> ${row.awb || ""}</p>
      <p><b>Order Id:</b> ${row.orderId || ""}</p>
      <p><b>Product:</b> ${row.product || ""}</p>
      <p><b>Qty:</b> ${row.qty || ""}</p>
      <p><b>Date:</b> ${row.date || ""}</p>
      <hr>
      <p>This e-POD serves as confirmation of successful delivery.</p>
      <div style="position:fixed;bottom:60px;right:60px;font-weight:bold;text-align:center;">
        Signature not required<br>--------<br>Delivered with OTP
      </div>
      </body></html>`;

      await page.setContent(html, { waitUntil: "networkidle0" });
      await page.pdf({
        path: path.join(outputDir, `${awb}.pdf`),
        format: "A4",
        printBackground: true
      });

      if (i % 100 === 0) console.log(`Done ${i}/${rows.length}`);
    }

    await browser.close();

    const zipPath = path.join(process.cwd(), "PODs.zip");
    const output = fs.createWriteStream(zipPath);
    const archive = archiver("zip");

    archive.pipe(output);
    archive.directory(outputDir, false);
    await archive.finalize();

    res.download(zipPath, "PODs.zip");

  } catch (error) {
    console.error(error);
    res.status(500).send("Error generating PODs");
  }
});

const PORT = process.env.PORT || 8080;

// ✅ IMPORTANT FIX (Railway access)
app.listen(PORT, "0.0.0.0", () => console.log("Running on", PORT));
