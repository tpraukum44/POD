/**
 * POD Generator – Proof of Delivery PDFs from Excel.
 * Deploy on Railway (1000+ orders). Env: PUPPETEER_EXECUTABLE_PATH (Dockerfile), CONCURRENT_PDFS, BROWSER_RESTART_EVERY, PDF_TIMEOUT_MS.
 */
import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import puppeteer from "puppeteer";
import archiver from "archiver";
import fs from "fs";
import path from "path";

const app = express();
const upload = multer({ dest: path.join(process.cwd(), "uploads") });

const MAX_PER_ZIP = 3000;
const IMAGES_DIR = path.join(process.cwd(), "images");
// Concurrency: Local 5 | Railway 1 (set CONCURRENT_PDFS=2 or 3 if more RAM)
const isRailwayEnv = !!process.env.PUPPETEER_EXECUTABLE_PATH;
const CONCURRENT_PDFS = Math.max(1, Number(process.env.CONCURRENT_PDFS) || (isRailwayEnv ? 1 : 5));
const PDF_TIMEOUT_MS = Number(process.env.PDF_TIMEOUT_MS) || 45000;
const MAX_PDF_RETRIES = 2;
// Restart browser every N PDFs to avoid Chromium degradation
const BROWSER_RESTART_EVERY = Math.max(500, Number(process.env.BROWSER_RESTART_EVERY) || 600);

// ================= UI =================
app.get("/", (req, res) => {
  res.send(`
<!DOCTYPE html>
<html>
<head>
<title>POD Generator</title>
<style>
body { font-family: Arial; background:#f4f6fb; padding:40px; }
.box { max-width:600px; margin:auto; background:#fff; padding:30px; border-radius:10px; box-shadow:0 10px 20px rgba(0,0,0,.1); }
h2 { text-align:center; }
input, button { width:100%; padding:12px; margin-top:10px; }
button { background:#111827; color:#fff; border:none; cursor:pointer; }
.note { color:#666; margin-top:10px; font-size:14px; }
</style>
</head>
<body>
  <div class="box">
    <h2>📦 POD Generator</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="excel" accept=".xlsx" required />
      <button type="submit">Generate POD ZIP</button>
    </form>
    <p class="note">⚡ Deploy on Railway for 1000+ orders. Auto splits into ZIPs of 3000 PDFs each.</p>
  </div>
</body>
</html>
`);
});

// ================= DATE FORMAT (FORCE DD-MM-YYYY) =================
function formatDate(value) {
  if (!value) return "";

  let d;

  if (typeof value === "number") {
    const utc_days = Math.floor(value - 25569);
    const utc_value = utc_days * 86400;
    d = new Date(utc_value * 1000);
  } else if (value instanceof Date) {
    d = value;
  } else {
    d = new Date(value);
  }

  if (isNaN(d.getTime())) return "";

  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();

  return `${dd}-${mm}-${yyyy}`;
}

// ================= UPLOAD =================
app.post("/upload", upload.single("excel"), async (req, res) => {
  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const rowsWithAwb = rows.filter((r) => r.awb != null && String(r.awb).trim() !== "");
    const totalRows = rowsWithAwb.length;
    if (totalRows === 0) {
      return res.status(400).send("No rows with AWB found in the Excel file.");
    }

    const workDir = path.join(process.cwd(), "work");
    if (fs.existsSync(workDir)) fs.rmSync(workDir, { recursive: true, force: true });
    fs.mkdirSync(workDir);

    // AWB validation: use full string (no trim) for duplicate check; assign unique filenames
    const awbCount = new Map();
    const sanitizeForFilename = (s) => String(s).replace(/[<>:"/\\|?*\x00-\x1f]/g, "_").trim() || "awb";

    const tasks = [];
    let duplicateAwbCount = 0;
    for (let index = 0; index < rowsWithAwb.length; index++) {
      const row = rowsWithAwb[index];
      const rawAwb = String(row.awb);
      const count = (awbCount.get(rawAwb) || 0) + 1;
      awbCount.set(rawAwb, count);
      if (count > 1) duplicateAwbCount++;

      const baseName = count === 1 ? sanitizeForFilename(rawAwb) : `${sanitizeForFilename(rawAwb)}_${count}`;
      const zipIndex = Math.floor(index / MAX_PER_ZIP) + 1;
      const folder = path.join(workDir, `part${zipIndex}`);
      tasks.push({
        row,
        rawAwb,
        filename: path.join(folder, `${baseName}.pdf`),
        folder,
        index
      });
    }

    if (duplicateAwbCount > 0) {
      console.log(`AWB validation: ${duplicateAwbCount} duplicate AWB(s) found; using unique filenames (e.g. AWB_2.pdf).`);
    }

    console.log(`Processing ${totalRows} orders with concurrency ${CONCURRENT_PDFS}, browser restart every ${BROWSER_RESTART_EVERY} PDFs...`);

    // Railway/Docker: PUPPETEER_EXECUTABLE_PATH set by Dockerfile
    const launchBrowser = () =>
      puppeteer.launch({
        headless: true,
        args: [
          "--no-sandbox",
          "--disable-setuid-sandbox",
          "--disable-dev-shm-usage",
          "--disable-gpu",
          "--disable-software-rasterizer",
          ...(isRailwayEnv ? ["--single-process", "--no-zygote"] : []),
          "--font-render-hinting=none"
        ],
        executablePath: process.env.PUPPETEER_EXECUTABLE_PATH
      });

    const uniqueFolders = [...new Set(tasks.map((t) => t.folder))];
    uniqueFolders.forEach((folder) => {
      if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });
    });

    const results = [];

    const generateOnePdf = async (task, browserInstance) => {
      let page;
      try {
        page = await browserInstance.newPage();
        await page.setJavaScriptEnabled(false);
        await page.setContent(buildHTML(task.row), {
          waitUntil: "domcontentloaded",
          timeout: PDF_TIMEOUT_MS
        });
        await new Promise((r) => setTimeout(r, 150));
        await page.pdf({
          path: task.filename,
          format: "A4",
          printBackground: true,
          timeout: PDF_TIMEOUT_MS
        });
        return { success: true, rawAwb: task.rawAwb };
      } catch (err) {
        return { success: false, rawAwb: task.rawAwb, error: err.message };
      } finally {
        if (page && !page.isClosed()) await page.close().catch(() => {});
      }
    };

    const processWithRetry = async (task, browserInstance) => {
      let lastResult = { success: false, rawAwb: task.rawAwb, error: "unknown" };
      for (let attempt = 1; attempt <= MAX_PDF_RETRIES; attempt++) {
        lastResult = await generateOnePdf(task, browserInstance);
        if (lastResult.success) return lastResult;
        if (attempt < MAX_PDF_RETRIES) console.warn(`Retry ${attempt}/${MAX_PDF_RETRIES} for AWB ${task.rawAwb}`);
      }
      return lastResult;
    };

    // Process in chunks; restart browser every BROWSER_RESTART_EVERY PDFs to avoid Chromium degradation
    for (let chunkStart = 0; chunkStart < tasks.length; chunkStart += BROWSER_RESTART_EVERY) {
      const chunkEnd = Math.min(chunkStart + BROWSER_RESTART_EVERY, tasks.length);
      const chunkTasks = tasks.slice(chunkStart, chunkEnd);
      const browser = await launchBrowser();

      for (let i = 0; i < chunkTasks.length; i += CONCURRENT_PDFS) {
        const batch = chunkTasks.slice(i, i + CONCURRENT_PDFS);
        const batchResults = await Promise.all(batch.map((t) => processWithRetry(t, browser)));
        results.push(...batchResults);

        const completed = results.length;
        const failed = results.filter((r) => !r.success).length;
        if (completed % 100 === 0 || completed === tasks.length) {
          console.log(`Progress: ${completed}/${tasks.length} PDFs (${failed} failed)`);
        }
      }

      await browser.close();
      if (chunkEnd < tasks.length) {
        console.log(`Browser restarted; continuing from PDF ${chunkEnd + 1}...`);
      }
    }

    results.filter((r) => !r.success).forEach((r) => console.error(`PDF failed for AWB ${r.rawAwb}: ${r.error}`));

    const successCount = results.filter(r => r.success).length;
    const failCount = results.filter(r => !r.success).length;
    console.log(`Completed: ${successCount} successful, ${failCount} failed`);

    // ===== ZIP ALL =====
    const finalZipPath = path.join(process.cwd(), "POD_ZIPS.zip");

    await new Promise((resolve, reject) => {
      const output = fs.createWriteStream(finalZipPath);
      const archive = archiver("zip", { zlib: { level: 9 } });

      output.on("close", resolve);
      archive.on("error", reject);

      archive.pipe(output);
      archive.directory(workDir, false);
      archive.finalize();
    });

    res.download(finalZipPath, "POD_ZIPS.zip");

  } catch (err) {
    console.error(err);
    res.status(500).send("Error: " + err.message);
  }
});

// ================= LOCAL IMAGES (no network wait) =================
// Put images in ./images/ — they are loaded once as base64 and embedded in PDFs.
function getLocalImageDataUrl(filename) {
  const base = path.basename(filename, path.extname(filename));
  const exts = [".png", ".jpg", ".jpeg"];
  for (const ext of exts) {
    const filePath = path.join(IMAGES_DIR, base + ext);
    if (!fs.existsSync(filePath)) continue;
    try {
      const buf = fs.readFileSync(filePath);
      const mime = ext === ".png" ? "image/png" : "image/jpeg";
      return `data:${mime};base64,${buf.toString("base64")}`;
    } catch {
      return null;
    }
  }
  return null;
}

// Courier logo: prefer local file in ./images/ (e.g. logo-shadowfax.png), else remote URL, else empty
const COURIER_LOCAL_FILES = {
  shadowfax: "logo-shadowfax.png",
  dtdc: "logo-dtdc.png",
  xpressbees: "logo-xpressbees.png",
  ekart: "logo-ekart.png",
  delhivery: "logo-delhivery.png"
};
const COURIER_REMOTE_URLS = {
  shadowfax: "https://www.rocketrybox.com/images/company6.png",
  dtdc: "https://www.rocketrybox.com/images/company4.png",
  xpressbees: "https://www.rocketrybox.com/images/company5.png",
  ekart: "https://www.rocketrybox.com/images/company7.png",
  delhivery: "https://www.rocketrybox.com/images/company2.png"
};

function getCourierLogo(courier) {
  if (!courier) return "";
  const c = courier.toLowerCase();
  for (const [key, localFile] of Object.entries(COURIER_LOCAL_FILES)) {
    if (c.includes(key)) {
      const dataUrl = getLocalImageDataUrl(localFile);
      if (dataUrl) return dataUrl;
      return COURIER_REMOTE_URLS[key] || "";
    }
  }
  return "";
}

// ================= PDF HTML =================
function buildHTML(row) {
  const courierLogo = getCourierLogo(row.courier);
  const watermark =
    getLocalImageDataUrl("watermark.png") || "https://www.rocketrybox.com/images/track-order.png";

  return `
<html>
<head>
<style>
@page { margin: 10mm; }
body { margin:0; padding:0; font-family: Arial; }

.page {
  border: 2px solid #000;
  padding: 12px;
  height: 100%;
  box-sizing: border-box;
  position: relative;
}

.watermark {
  position: absolute;
  top: 50%;
  left: 50%;
  width: 420px;
  opacity: 0.12;
  transform: translate(-50%, -50%) rotate(-15deg);
  z-index: 0;
}

.content { position: relative; z-index: 1; }

.header {
  display:flex;
  justify-content:space-between;
  align-items:center;
  border-bottom:1px solid #000;
  padding-bottom:8px;
  margin-bottom:8px;
}

.header img { height:40px; }

.title { font-size:20px; font-weight:bold; }

hr { border:none; border-top:1px solid #000; }

.footer {
  position: absolute;
  bottom: 20px;
  left: 12px;
  right: 12px;
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  gap: 20px;
}

.footer-left {
  font-size: 13px;
  line-height: 1.4;
  max-width: 70%;
}

.footer-right {
  width: 220px;
  text-align: center;
}

.sign-box {
  border-top: 1px solid #000;
  padding-top: 6px;
  font-weight: bold;
}

.sign-box .line {
  border-top: 1px dashed #000;
  margin: 6px 0;
}
</style>
</head>
<body>
  <div class="page">

    <img src="${watermark}" class="watermark">

    <div class="content">

      <div class="header">
        <img src="${courierLogo}">
        <div class="title">Proof Of Delivery (E-POD)</div>
      </div>

      <p><b>Courier:</b> ${row.courier || ""}</p>
      <p><b>AWB:</b> ${row.awb || ""}</p>
      <p><b>Order Id:</b> ${row.orderId || ""}</p>
      <p><b>Product:</b> ${row.product || ""}</p>
      <p><b>Qty:</b> ${row.qty || ""}</p>
      <p><b>Delivery Date:</b> ${formatDate(row.date)}</p>

      <hr>

      <h3>Seller Details</h3>
      <p><b>Company:</b> ${row.sellerCompany || ""}</p>
      <p><b>Seller Id:</b> ${row.sellerId || ""}</p>

      <h3>Consignee Details</h3>
      <p><b>Name:</b> ${row.name || ""}</p>
      <p><b>Address:</b> ${row.address || ""}</p>
      <p><b>State:</b> ${row.state || ""}</p>
      <p><b>Pincode:</b> ${row.pincode || ""}</p>
      <p><b>Phone:</b> ${row.phone || ""}</p>

    </div>

    <div class="footer">
      <div class="footer-left">
        This e-POD serves as confirmation that the shipment has been successfully delivered and verified by OTP.
        Any dispute must be reported to <b>customer support team</b> within 7 days from delivery date.
      </div>

      <div class="footer-right">
        <div class="sign-box">
          Signature not required
          <div class="line"></div>
          Delivered with OTP
        </div>
      </div>
    </div>

  </div>
</body>
</html>
`;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
