"use strict";

const express = require("express");
const PDFDocument = require("pdfkit");
const nodemailer = require("nodemailer");
const cors = require("cors");
const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const sharp = require("sharp");
const XLSX = require("xlsx");
const OpenAI = require("openai"); // npm i openai
const { v4: uuidv4 } = require("uuid"); // npm i uuid

const app = express();
const { URL } = require("url");
/* ========================= CONFIG FIXE ========================= */
const SMTP_HOST = "avocarbon-com.mail.protection.outlook.com";
const SMTP_PORT = 25;
const EMAIL_FROM_NAME = "Administration STS";
const EMAIL_FROM = "administration.STS@avocarbon.com";
const SUPPORT_EMAIL = "chaima.benyahia@avocarbon.com";

const TEMP_UPLOAD_FOLDER = path.join(process.cwd(), "temp_uploads");
const MAX_AGE_MINUTES = 30;

if (!fs.existsSync(TEMP_UPLOAD_FOLDER)) {
  fs.mkdirSync(TEMP_UPLOAD_FOLDER, { recursive: true });
}

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY }); // Azure App Settings

/* ========================= MIDDLEWARES ========================= */
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));
app.use("/static", express.static(path.join(process.cwd(), "assets")));

app.use(
  cors({
    origin: true,
    credentials: true,
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: [
      "Content-Type",
      "Authorization",
      "openai-conversation-id",
      "openai-ephemeral-user-id",
    ],
  })
);
app.options("*", cors());

app.use((req, _res, next) => {
  console.log(`${new Date().toISOString()} - ${req.method} ${req.path}`);
  next();
});

/* ====================== TRANSPORTEUR SMTP ====================== */
const transporter = nodemailer.createTransport({
  host: SMTP_HOST,
  port: SMTP_PORT,
  secure: false,
  tls: { minVersion: "TLSv1.2" },
});

transporter
  .verify()
  .then(() => console.log("✅ SMTP EOP prêt"))
  .catch((err) =>
    console.error("❌ SMTP erreur:", err && err.message ? err.message : String(err))
  );

/* ============================ UTILS ============================ */
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function fetchUrlToBuffer(url) {
  return new Promise((resolve, reject) => {
    const client = url.startsWith("https") ? https : http;
    const req = client.get(url, (res) => {
      // redirects
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        const nextUrl = res.headers.location.startsWith("http")
          ? res.headers.location
          : new URL(res.headers.location, url).toString();
        res.resume();
        return resolve(fetchUrlToBuffer(nextUrl));
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error(`HTTP ${res.statusCode} pour ${url}`));
      }
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => resolve(Buffer.concat(chunks)));
    });
    req.on("error", reject);
    req.setTimeout(15000, () => req.destroy(new Error("Timeout téléchargement")));
  });
}

async function loadImageToBuffer({ imageUrl, imagePath }) {
  if (imageUrl) return await fetchUrlToBuffer(imageUrl);

  if (imagePath) {
    const full = path.resolve(imagePath);
    if (!fs.existsSync(full)) {
      throw new Error(`Fichier image introuvable: ${full} (cwd=${process.cwd()})`);
    }
    return fs.promises.readFile(full);
  }

  throw new Error("Aucune source d'image fournie (imageUrl | imagePath).");
}

function validateImageBuffer(buffer) {
  if (!buffer || !Buffer.isBuffer(buffer)) throw new Error("Image non Buffer");
  if (buffer.length < 10) throw new Error(`Image trop petite (${buffer.length} octets)`);

  const isPNG = buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47;
  const isJPEG = buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff;
  const isGIF = buffer[0] === 0x47 && buffer[1] === 0x49 && buffer[2] === 0x46;

  console.log(
    `Image validation - PNG:${isPNG}, JPEG:${isJPEG}, GIF:${isGIF} | Magic: ${buffer[0]
      .toString(16)
      .padStart(2, "0")} ${buffer[1].toString(16).padStart(2, "0")} ${buffer[2]
      .toString(16)
      .padStart(2, "0")} ${buffer[3].toString(16).padStart(2, "0")}`
  );

  if (!isPNG && !isJPEG && !isGIF) {
    console.warn("⚠️ Format non reconnu — tentative avec sharp/PDFKit tout de même.");
  }
  return true;
}

async function normalizeImageBuffer(buffer, { format = "png" } = {}) {
  const img = sharp(buffer, { failOn: "none" }).rotate();
  if (format === "png") return await img.png({ compressionLevel: 9 }).toBuffer();
  return await img.jpeg({ quality: 90, chromaSubsampling: "4:4:4" }).toBuffer();
}

/* ===================== TEMP FILES: SAVE + CLEANUP ===================== */
function cleanupOldFiles(folder, maxAgeMinutes = 30) {
  const cutoff = Date.now() - maxAgeMinutes * 60 * 1000;
  try {
    for (const file of fs.readdirSync(folder)) {
      const full = path.join(folder, file);
      try {
        const st = fs.statSync(full);
        if (st.isFile() && st.mtimeMs < cutoff) fs.unlinkSync(full);
      } catch (_) {}
    }
  } catch (_) {}
}

// Nettoyage automatique toutes les 10 minutes
setInterval(() => cleanupOldFiles(TEMP_UPLOAD_FOLDER, MAX_AGE_MINUTES), 10 * 60 * 1000);

function sanitizeFilename(name) {
  const raw = String(name || "uploaded_file").trim();
  const base = raw.replace(/[/\\?%*:|"<>]/g, "_").replace(/\s+/g, "_");
  return base.replace(/[^a-z0-9._-]/gi, "_") || "uploaded_file";
}

function getExtLower(filename) {
  return path.extname(String(filename || "")).toLowerCase();
}

const ALLOWED_IMAGE_EXT = new Set([".png", ".jpg", ".jpeg", ".gif", ".webp"]);

function isAllowedImageName(filename) {
  const ext = getExtLower(filename);
  return ALLOWED_IMAGE_EXT.has(ext);
}

function saveBytesToTemp(fileBytes, originalName = "uploaded_file.bin") {
  const safe = sanitizeFilename(originalName);
  const finalName = `${uuidv4().slice(0, 8)}_${Date.now()}_${safe}`;
  const localPath = path.join(TEMP_UPLOAD_FOLDER, finalName);
  fs.writeFileSync(localPath, fileBytes);
  return { filename: finalName, localPath, expiresInMinutes: MAX_AGE_MINUTES };
}

/* ===================== OPENAI FILE REFS HANDLING ===================== */
function normalizeOpenAIFileRef(ref) {
  if (!ref) return null;

  if (typeof ref === "string") {
    return { id: ref, name: "uploaded_file.bin", download_link: null };
  }

  const id = ref.id || ref.file_id || ref.fileId;
  const download_link = ref.download_link || ref.downloadUrl || ref.url || null;
  const name = ref.name || ref.filename || ref.original_name || "uploaded_file.bin";

  if (!id && !download_link) return null;

  return { id, name, download_link };
}

async function downloadBytesFromOpenAIRef(nref) {
  if (!nref) throw new Error("Ref OpenAI vide");

  const dl = nref.download_link;

  // ✅ Use download_link ONLY if it's a real http(s) URL
  if (dl && /^https?:\/\//i.test(String(dl))) {
    return await fetchUrlToBuffer(dl);
  }

  // ✅ Sometimes download_link is like "file_..." → treat it as an id
  const possibleId =
    nref.id ||
    (dl && /^file_/i.test(String(dl)) ? String(dl) : null);

  if (!possibleId) {
    throw new Error("Ref OpenAI sans id et sans download_link http(s)");
  }

  // OpenAI Files API -> bytes
  const resp = await openai.files.content(possibleId);
  const ab = await resp.arrayBuffer();
  return Buffer.from(ab);
}


function pickAppendixRef(openaiFileIdRefs, appendixOpenAIFile) {
  const prefer = normalizeOpenAIFileRef(appendixOpenAIFile);
  if (prefer) return prefer;

  const refs = Array.isArray(openaiFileIdRefs) ? openaiFileIdRefs : [];
  const normalized = refs.map(normalizeOpenAIFileRef).filter(Boolean);

  // choose first image by name ext if possible
  const img = normalized.find((r) => isAllowedImageName(r.name));
  return img || normalized[0] || null;
}

/* ============================ PDF (REPORT) ============================ */
function generatePDF(content) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({
        margin: 50,
        size: "A4",
        bufferPages: true,
        info: { Title: content.title, Author: "Assistant GPT", Subject: content.title },
      });

      const chunks = [];
      doc.on("data", (c) => chunks.push(c));
      doc.on("end", () => resolve(Buffer.concat(chunks)));
      doc.on("error", reject);

      doc.fontSize(26).font("Helvetica-Bold").fillColor("#1e40af").text(content.title, { align: "center" });
      doc.moveDown(0.5);
      doc.strokeColor("#3b82f6").lineWidth(2).moveTo(50, doc.y).lineTo(doc.page.width - 50, doc.y).stroke();
      doc.moveDown();

      doc
        .fontSize(10)
        .fillColor("#6b7280")
        .font("Helvetica")
        .text(
          `Date: ${new Date().toLocaleDateString("fr-FR", {
            year: "numeric",
            month: "long",
            day: "numeric",
          })}`,
          { align: "right" }
        );
      doc.moveDown(2);

      if (content.introduction) {
        doc.fontSize(16).font("Helvetica-Bold").fillColor("#1f2937").text("Introduction");
        doc.moveDown(0.5);
        doc.fontSize(11).font("Helvetica").fillColor("#374151").text(content.introduction, { align: "justify", lineGap: 3 });
        doc.moveDown(2);
      }

      if (Array.isArray(content.sections)) {
        (async () => {
          for (let index = 0; index < content.sections.length; index++) {
            const section = content.sections[index];

            if (doc.y > doc.page.height - 150) doc.addPage();

            doc.fontSize(14).font("Helvetica-Bold").fillColor("#1e40af").text(`${index + 1}. ${section.title || "Section"}`);
            doc.moveDown(0.5);

            if (section.content) {
              doc.fontSize(11).font("Helvetica").fillColor("#374151").text(section.content, { align: "justify", lineGap: 3 });
              doc.moveDown(1);
            }

            // keep imageUrl/imagePath for the report generator only (not base64)
            if (section.imageUrl || section.imagePath) {
              try {
                console.log("=== Insertion image ===");
                const buf = await loadImageToBuffer({ imageUrl: section.imageUrl, imagePath: section.imagePath });
                validateImageBuffer(buf);

                const maxWidth = doc.page.width - 100;
                const maxHeight = 300;

                if (doc.y > doc.page.height - maxHeight - 100) doc.addPage();

                try {
                  doc.image(buf, { fit: [maxWidth, maxHeight], align: "center" });
                } catch (err1) {
                  const m1 = String((err1 && (err1.message || err1.reason)) ?? err1);
                  console.warn("Image insert error:", m1);

                  console.log("🧼 Normalisation image via sharp (PNG)...");
                  const normalized = await normalizeImageBuffer(buf, { format: "png" });
                  validateImageBuffer(normalized);
                  doc.image(normalized, { fit: [maxWidth, maxHeight], align: "center" });
                }

                doc.moveDown(1);
                if (section.imageCaption) {
                  doc.fontSize(9).fillColor("#6b7280").font("Helvetica-Oblique").text(section.imageCaption, { align: "center" });
                  doc.moveDown(1);
                }
              } catch (imgError) {
                const msg = (imgError && (imgError.message || imgError.reason)) ?? String(imgError);
                console.error("❌ Erreur image:", msg);
                doc.fontSize(10).fillColor("#ef4444").text("⚠️ Erreur lors du chargement de l'image", { align: "center" });
                doc.fontSize(8).fillColor("#9ca3af").text(`(${msg})`, { align: "center" });
                doc.moveDown(1);
              }
            }

            doc.moveDown(1.5);
          }

          if (content.conclusion) {
            if (doc.y > doc.page.height - 150) doc.addPage();
            doc.fontSize(16).font("Helvetica-Bold").fillColor("#1f2937").text("Conclusion");
            doc.moveDown(0.5);
            doc.fontSize(11).font("Helvetica").fillColor("#374151").text(content.conclusion, { align: "justify", lineGap: 3 });
          }

          const range = doc.bufferedPageRange();
          for (let i = 0; i < range.count; i++) {
            doc.switchToPage(i);
            const oldY = doc.y;
            doc.fontSize(8).fillColor("#9ca3af");
            doc.text(`Page ${i + 1} sur ${range.count}`, 50, doc.page.height - 50, {
              align: "center",
              lineBreak: false,
              width: doc.page.width - 100,
            });
            if (i < range.count - 1) {
              doc.switchToPage(i);
              doc.y = oldY;
            }
          }

          doc.end();
        })().catch(reject);
      } else {
        doc.end();
      }
    } catch (err) {
      console.error("Erreur génération PDF:", err && err.stack ? err.stack : String(err));
      reject(err);
    }
  });
}

async function sendEmailWithPdf({ to, subject, messageHtml, pdfBuffer, pdfFilename }) {
  return transporter.sendMail({
    from: { name: EMAIL_FROM_NAME, address: EMAIL_FROM },
    to,
    subject,
    html: messageHtml,
    attachments: [{ filename: pdfFilename, content: pdfBuffer, contentType: "application/pdf" }],
  });
}

/* ============================ OFFER PDF (LOGO + HEADER) ============================ */
const COMPANY_ADDRESS_MAP = {
  "avocarbon france": [
    "AVOCarbon France - 9 rue des imprimeurs - Z.I. de la République n° 1 - 86000 POITIERS France",
    "au capital de 3 224 460 € - RCS Poitiers B339 348 450 – Code APE 2732 Z – N° identification TVA FR 01339348450",
    "Phone : +33 5 49 62 25 00",
  ],
  france: [
    "AVOCarbon France - 9 rue des imprimeurs - Z.I. de la République n° 1 - 86000 POITIERS France",
    "au capital de 3 224 460 € - RCS Poitiers B339 348 450 – Code APE 2732 Z – N° identification TVA FR 01339348450",
    "Phone : +33 5 49 62 25 00",
  ],

  "avocarbon germany": ["AVOCarbon Germany", "AVOCarbon Germany GmbH", "Talstrasse 112", "D-60437 Frankfurt am Main"],
  germany: ["AVOCarbon Germany", "AVOCarbon Germany GmbH", "Talstrasse 112", "D-60437 Frankfurt am Main"],

  "avocarbon india": ["AVOCarbon India", "25/A2, Dairy Plant Road SIDCO Industrial Estate (NP)", "Pattaravakka Ambattur Chennai – 600098", "Tamilnadu"],
  india: ["AVOCarbon India", "25/A2, Dairy Plant Road SIDCO Industrial Estate (NP)", "Pattaravakka Ambattur Chennai – 600098", "Tamilnadu"],

  "avocarbon korea": ["AVOCarbon Korea", "306, Nongong-ro, Nongong-eup", "Dalseong-Gun, Daegu"],
  korea: ["AVOCarbon Korea", "306, Nongong-ro, Nongong-eup", "Dalseong-Gun, Daegu"],

  "assymex monterrey": ["ASSYMEX MONTERREY", "San Sebastian 110", "Co. Los Lermas", "GUADALUPE, N.L", "Mexico 67190"],
  monterrey: ["ASSYMEX MONTERREY", "San Sebastian 110", "Co. Los Lermas", "GUADALUPE, N.L", "Mexico 67190"],

  tunisia: ["AVOCarbon", "Tunisia", "SCEET & SAME", "Zone industrielle Elfahs", "1140 Zaghouane"],
  tunis: ["AVOCarbon", "Tunisia", "SCEET & SAME", "Zone industrielle Elfahs", "1140 Zaghouane"],

  tianjin: ["AVOCarbon Tianjin", "Junling Road 17 # Beizhakou", "Jinnan District"],
  kunshan: ["AVOCarbon Kunshan", "N°9, Dongtinghu Road", "215335 Kunshan"],
};

function normalizeKey(v) {
  return String(v || "").trim().toLowerCase();
}

function getCompanyAddressLines(offer) {
  const key =
    normalizeKey(offer?.company) ||
    normalizeKey(offer?.site) ||
    normalizeKey(offer?.entity) ||
    "france";
  return COMPANY_ADDRESS_MAP[key] || COMPANY_ADDRESS_MAP.france;
}

function drawOfferHeader(doc, offer, logoBuf) {
  const pageW = doc.page.width;

  const TOP_BAR_H = 16;
  const HEADER_BLOCK_H = 90;
  const BOTTOM_BAR_H = 10;
  const HEADER_TOTAL_H = TOP_BAR_H + HEADER_BLOCK_H + BOTTOM_BAR_H;

  doc.save();
  doc.fillColor("#0b5fa5").rect(0, 0, pageW, TOP_BAR_H).fill();
  doc.restore();

  const addrLines = getCompanyAddressLines(offer) || [];
  const addrX = 50;
  const addrY = TOP_BAR_H + 12;

  doc.font("Helvetica").fontSize(8).fillColor("#111827");
  if (addrLines.length > 0) {
    doc.text(addrLines.join("\n"), addrX, addrY, {
      width: pageW - 100 - 170,
      lineGap: 1,
    });
  }

  if (logoBuf) {
    const logoW = 140;
    const x = pageW - 50 - logoW;
    const y = TOP_BAR_H + 12;
    doc.image(logoBuf, x, y, { width: logoW });
  }

  doc.save();
  doc.fillColor("#0b5fa5").rect(0, TOP_BAR_H + HEADER_BLOCK_H, pageW, BOTTOM_BAR_H).fill();
  doc.restore();

  doc.y = HEADER_TOTAL_H + 18;
}

function generateOfferPDFWithLogo(offer) {
  return new Promise((resolve, reject) => {
    (async () => {
      try {
        const doc = new PDFDocument({
          margin: 50,
          size: "A4",
          bufferPages: true,
          info: {
            Title: offer?.subject || "Commercial Offer",
            Author: "AVOCarbon",
            Subject: offer?.subject || "Offer",
          },
        });

        const chunks = [];
        doc.on("data", (c) => chunks.push(c));
        doc.on("end", () => resolve(Buffer.concat(chunks)));
        doc.on("error", reject);

        // ====== LOGO ======
        const logoPath = path.join(process.cwd(), "assets", "logo_avocarbon.jpg");
        let logoBuf = null;

        try {
          logoBuf = await loadImageToBuffer({ imagePath: logoPath });
          validateImageBuffer(logoBuf);
        } catch (e) {
          console.warn("⚠️ Logo non chargé:", e?.message ?? String(e));
          logoBuf = null;
        }

        // Header page 1 + toutes pages
        drawOfferHeader(doc, offer, logoBuf);
        doc.on("pageAdded", () => drawOfferHeader(doc, offer, logoBuf));

        // ====== TITRE ======
        doc.font("Helvetica-Bold").fontSize(18).fillColor("#111827").text(offer?.title || "COMMERCIAL OFFER", { align: "left" });

        doc.moveDown(0.4);
        doc.font("Helvetica").fontSize(10).fillColor("#374151").text(
          `Date: ${
            offer?.date ||
            new Date().toLocaleDateString("fr-FR", { year: "numeric", month: "long", day: "numeric" })
          }`
        );

        doc.moveDown(1);

        // ====== CUSTOMER ======
        const customerLines = Array.isArray(offer?.customerLines) ? offer.customerLines : [];
        const hasCustomer = customerLines.length > 0 || offer?.customerName || offer?.customerAddress || offer?.toPerson;

        if (hasCustomer) {
          doc.font("Helvetica-Bold").fontSize(11).fillColor("#111827").text("Customer");
          doc.font("Helvetica").fontSize(10).fillColor("#374151");

          if (customerLines.length > 0) {
            doc.text(customerLines.join("\n"));
          } else {
            if (offer?.customerName) doc.text(offer.customerName);
            if (offer?.customerAddress) doc.text(offer.customerAddress);
            if (offer?.toPerson) doc.text(`To: ${offer.toPerson}`);
          }
          doc.moveDown(1);
        }

        // ====== SUBJECT ======
        if (offer?.subject) {
          const y0 = doc.y;
          doc.save().fillColor("#dbeafe").rect(50, y0, doc.page.width - 100, 22).fill().restore();
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111827").text(offer.subject, 58, y0 + 6);
          doc.moveDown(2);
        }

        // ====== INTRO ======
        if (offer?.intro) {
          doc.font("Helvetica").fontSize(11).fillColor("#111827").text(offer.intro, { align: "justify", lineGap: 3 });
          doc.moveDown(1);
        }

        // ====== SECTIONS ======
        const sections = Array.isArray(offer?.sections) ? offer.sections : [];
        for (const s of sections) {
          if (doc.y > doc.page.height - 160) doc.addPage();

          doc.font("Helvetica-Bold").fontSize(12).fillColor("#1e40af").text(s.title || "Section");
          doc.moveDown(0.3);
          doc.font("Helvetica").fontSize(11).fillColor("#111827").text(s.content || "", { align: "justify", lineGap: 3 });
          doc.moveDown(0.8);
        }

        // ====== SIGNATURE ======
        doc.moveDown(1);
        doc.font("Helvetica").fontSize(11).fillColor("#111827");
        if (offer?.closing) doc.text(offer.closing);
        if (offer?.signatureName) doc.text(offer.signatureName);
        if (offer?.signatureTitle) doc.text(offer.signatureTitle);

        // ====== APPENDIX IMAGE (PATH ONLY - NO BASE64) ======
        if (offer?.appendixImagePath || offer?.appendixImageUrl) {
          doc.addPage();
          doc.moveDown(0.5);

          doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text(offer?.appendixImageTitle || "Appendix - Drawing / Photo");
          doc.moveDown(0.5);

          try {
            const imgBuf = await loadImageToBuffer({
              imagePath: offer.appendixImagePath,
              imageUrl: offer.appendixImageUrl,
            });

            validateImageBuffer(imgBuf);

            const normalized = await normalizeImageBuffer(imgBuf, { format: "png" });

            const maxW = doc.page.width - 100;
            const maxH = 420;
            doc.image(normalized, { fit: [maxW, maxH], align: "center" });
          } catch (e) {
            const msg = e?.message ?? String(e);
            doc.font("Helvetica").fontSize(10).fillColor("#ef4444").text("⚠️ Appendix image not loaded.");
            doc.font("Helvetica").fontSize(8).fillColor("#9ca3af").text(`(${msg})`);
          }
        }

        doc.end();
      } catch (err) {
        reject(err);
      }
    })();
  });
}

/* ============================ ROUTES ============================ */
// Debug simple
app.post("/api/echo", (req, res) => {
  res.json({ ok: true, got: req.body || {} });
});

/**
 * ROUTE COMPLETE:
 * - reçoit offer + openaiFileIdRefs (+ option appendixOpenAIFile)
 * - download bytes -> save temp (30 min) pour tous les fichiers
 * - pick 1 image pour annexe -> offer.appendixImagePath
 * - generate PDF + send email
 * - renvoie savedFiles + errors
 */
app.post("/api/generate-offer-and-send", async (req, res) => {
  try {
    const { email, subject, offer, cc, openaiFileIdRefs, appendixOpenAIFile } = req.body || {};

    // ---- validations ----
    if (!email || !subject || !offer) {
      return res.status(400).json({
        success: false,
        error: "Données manquantes",
        details: "Envoyez email, subject, offer",
      });
    }

    if (!isValidEmail(email)) {
      return res.status(400).json({ success: false, error: "Email invalide" });
    }

    if (cc && !isValidEmail(cc)) {
      return res.status(400).json({
        success: false,
        error: "Adresse email CC invalide",
        details: `cc = ${cc}`,
      });
    }

    offer.subject = offer.subject || subject;

    // cleanup opportuniste
    cleanupOldFiles(TEMP_UPLOAD_FOLDER, MAX_AGE_MINUTES);

    // ---- 1) Save ALL refs to temp (savedFiles + errors) ----
    const refs = Array.isArray(openaiFileIdRefs) ? openaiFileIdRefs : [];
    const normalizedRefs = refs.map(normalizeOpenAIFileRef).filter(Boolean);

    const savedFiles = [];
    const errors = [];

    for (const r of normalizedRefs) {
      try {
        const bytes = await downloadBytesFromOpenAIRef(r);
        const saved = saveBytesToTemp(bytes, r.name || "uploaded_file.bin");
        savedFiles.push({
           id: r.id || null,
           download_link: r.download_link || null,
           name: r.name || saved.filename,
           localPath: saved.localPath,
           expiresInMinutes: saved.expiresInMinutes,
        });

        });
      } catch (e) {
        errors.push({
          ref: r.id || r.download_link || "unknown_ref",
          name: r.name || null,
          error: e?.message ?? String(e),
        });
      }
    }

    // ---- 2) pick appendix ref (prefer appendixOpenAIFile else first image) ----
    const appendixRef = pickAppendixRef(openaiFileIdRefs, appendixOpenAIFile);

    if (appendixRef) {
      // ensure it exists in savedFiles, else download+save (when appendixOpenAIFile is not in refs)
      const already = savedFiles.find((f) => {
          if (!f.localPath) return false;
          if (appendixRef.id && f.id && f.id === appendixRef.id) return true;
                 // match by download_link
          if (appendixRef.download_link && f.download_link && f.download_link === appendixRef.download_link) return true;
         // fallback match by name
          if (appendixRef.name && f.name && f.name === appendixRef.name) return true;
          return false;
      });

      if (already) {
        offer.appendixImagePath = already.localPath;
        offer.appendixExpiresInMinutes = already.expiresInMinutes;
      } else {
        try {
          // validate name ext if image
          const appendixName = appendixRef.name || "appendix.png";
          if (!isAllowedImageName(appendixName)) {
            throw new Error(`Annexe: extension non autorisée (${getExtLower(appendixName) || "no-ext"})`);
          }

          const bytes = await downloadBytesFromOpenAIRef(appendixRef);
          const saved = saveBytesToTemp(bytes, appendixName);

          // optional extra validation
          try {
            const buf = await loadImageToBuffer({ imagePath: saved.localPath });
            validateImageBuffer(buf);
          } catch (_) {
            // leave it; PDF generator will attempt sharp normalize and may still work
          }

          offer.appendixImagePath = saved.localPath;
          offer.appendixExpiresInMinutes = saved.expiresInMinutes;

          // also push to savedFiles if not present
          savedFiles.push({
            id: appendixRef.id || null,
            name: appendixName,
            localPath: saved.localPath,
            expiresInMinutes: saved.expiresInMinutes,
          });
        } catch (e) {
          errors.push({
            ref: appendixRef.id || appendixRef.download_link || "appendix_ref",
            name: appendixRef.name || null,
            error: e?.message ?? String(e),
          });
        }
      }
    }

    // ---- 3) Generate PDF + send email ----
    const pdfBuffer = await generateOfferPDFWithLogo(offer);
    const pdfName = `offre_${Date.now()}.pdf`;

    const html = `
      <!DOCTYPE html>
      <html>
        <body style="font-family: Arial, sans-serif; line-height:1.6; color:#111827;">
          <h2 style="margin:0 0 8px 0;">📄 Votre offre est prête</h2>
          <div style="background:#e0e7ff;padding:12px;border-left:4px solid #667eea;border-radius:6px;margin:12px 0;">
            <strong>📧 Sujet :</strong> ${escapeHtml(subject)}<br>
            <strong>📅 Date :</strong> ${new Date().toLocaleDateString("fr-FR")}
          </div>
          <p>Vous trouverez l’offre commerciale en pièce jointe (PDF).</p>
          <p style="color:#6b7280;font-size:12px">© ${new Date().getFullYear()} ${EMAIL_FROM_NAME}</p>
        </body>
      </html>
    `;

    await transporter.sendMail({
      from: { name: EMAIL_FROM_NAME, address: EMAIL_FROM },
      to: email,
      cc: cc || undefined,
      subject,
      html,
      attachments: [{ filename: pdfName, content: pdfBuffer, contentType: "application/pdf" }],
    });

    return res.json({
      success: true,
      message: "Offre générée et envoyée avec succès",
      savedFiles,
      errors,
      details: {
        email,
        cc: cc || null,
        filename: pdfName,
        pdfSize: `${(pdfBuffer.length / 1024).toFixed(2)} KB`,
        appendixTempFile: offer.appendixImagePath || null,
        appendixExpiresInMinutes: offer.appendixExpiresInMinutes || null,
      },
    });
  } catch (err) {
    console.error("❌ Erreur génération/envoi OFFRE:", err?.stack || String(err));
    return res.status(500).json({
      success: false,
      error: "Erreur lors du traitement",
      details: err?.message ?? String(err),
    });
  }
});

// ========== ROUTE : ENVOI EMAIL SIMPLE (SANS PJ) ==========
app.post("/api/send-email", async (req, res) => {
  try {
    const { email, subject, message, messageHtml, cc } = req.body || {};

    if (!email || !subject || (!message && !messageHtml)) {
      return res.status(400).json({
        success: false,
        error: "Données manquantes",
        details: "Envoyez email, subject, et (message ou messageHtml)",
      });
    }

    if (!isValidEmail(email)) {
      return res.status(400).json({ success: false, error: "Email invalide" });
    }

    if (cc && !isValidEmail(cc)) {
      return res.status(400).json({
        success: false,
        error: "Adresse email CC invalide",
        details: `cc = ${cc}`,
      });
    }

    const html =
      messageHtml ||
      `<!DOCTYPE html>
<html>
  <body style="font-family: Arial, sans-serif; line-height:1.6; color:#111827;">
    <h2 style="margin:0 0 8px 0;">📩 ${escapeHtml(subject)}</h2>
    <div style="background:#f9fafb;padding:12px;border:1px solid #e5e7eb;border-radius:6px;">
      <p style="white-space:pre-wrap;margin:0;">${escapeHtml(message)}</p>
    </div>
    <p style="color:#6b7280;font-size:12px;margin-top:16px;">
      © ${new Date().getFullYear()} ${EMAIL_FROM_NAME}
    </p>
  </body>
</html>`;

    await transporter.sendMail({
      from: { name: EMAIL_FROM_NAME, address: EMAIL_FROM },
      to: email,
      cc: cc || undefined,
      subject,
      html,
      text: message ? String(message) : undefined,
    });

    console.log(`✅ Email simple envoyé à ${email}${cc ? " (CC: " + cc + ")" : ""}`);

    return res.json({
      success: true,
      message: "Email envoyé avec succès",
      details: {
        email,
        cc: cc || null,
        subject,
        timestamp: new Date().toISOString(),
      },
    });
  } catch (err) {
    console.error("❌ Erreur envoi email simple (SMTP):", {
      message: err?.message,
      code: err?.code,
      response: err?.response,
      responseCode: err?.responseCode,
      command: err?.command,
      stack: err?.stack,
    });

    return res.status(500).json({
      success: false,
      error: "Erreur lors de l'envoi de l'email",
      details: err?.message ?? String(err),
      smtp: {
        code: err?.code,
        responseCode: err?.responseCode,
        response: err?.response,
        command: err?.command,
      },
    });
  }
});

// ========== ROUTE : ENVOI EMAIL SUPPORT ==========
app.post("/api/support/send-email", async (req, res) => {
  try {
    const { username, comment, assistant_name } = req.body || {};

    if (!username || !comment || !assistant_name) {
      return res.status(400).json({
        success: false,
        error: "Données manquantes",
        details: "Envoyez username, comment, assistant_name",
      });
    }

    const timestamp = new Date().toLocaleString("fr-FR", {
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });

    const emailHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; line-height: 1.6; color: #111827; max-width: 600px; margin: 0 auto; padding: 20px; }
            .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: black; padding: 20px; border-radius: 8px 8px 0 0; text-align: center; }
            .header h1 { color: #ff0000; margin: 0; font-size: 48px; font-weight: 900; letter-spacing: 8px; text-shadow: 0 0 10px rgba(255,0,0,.8), 0 0 20px rgba(255,0,0,.6), 0 0 30px rgba(255,0,0,.4), 2px 2px 4px rgba(0,0,0,.3); }
            .content { background: #f9fafb; padding: 20px; border: 1px solid #e5e7eb; border-top: none; }
            .info-box { background: white; border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; border-radius: 4px; }
            .info-box strong { color: #667eea; display: inline-block; width: 150px; }
            .comment-box { background: white; border: 2px solid #fbbf24; padding: 15px; margin: 15px 0; border-radius: 4px; }
            .comment-box h3 { margin-top: 0; color: #f59e0b; }
            .footer { text-align: center; margin-top: 20px; padding: 15px; background: #f3f4f6; border-radius: 0 0 8px 8px; font-size: 12px; color: #6b7280; }
            .priority { display: inline-block; background: #ef4444; color: white; padding: 5px 10px; border-radius: 4px; font-size: 12px; font-weight: bold; }
          </style>
        </head>
        <body>
          <div class="header"><h1>🆘</h1></div>
          <div class="content">
            <div style="text-align: center; margin-bottom: 20px;"><span class="priority">NOUVEAU TICKET</span></div>
            <div class="info-box">
              <strong>👤 Utilisateur :</strong> ${escapeHtml(username)}<br>
              <strong>🤖 Assistant :</strong> ${escapeHtml(assistant_name)}<br>
              <strong>📅 Date :</strong> ${escapeHtml(timestamp)}
            </div>
            <div class="comment-box">
              <h3>💬 Commentaire / Problème :</h3>
              <p style="white-space: pre-wrap; margin: 0;">${escapeHtml(comment)}</p>
            </div>
            <div style="background: #e0e7ff; padding: 12px; border-radius: 4px; margin-top: 15px;">
              <strong>ℹ️ Action requise :</strong><br>
              Veuillez traiter cette demande de support dans les plus brefs délais.
            </div>
          </div>
          <div class="footer">
            <p>
              Email envoyé automatiquement par le système de support<br>
              <strong>${EMAIL_FROM_NAME}</strong><br>
              © ${new Date().getFullYear()} - Ne pas répondre à cet email
            </p>
          </div>
        </body>
      </html>
    `;

    await transporter.sendMail({
      from: { name: EMAIL_FROM_NAME, address: EMAIL_FROM },
      to: SUPPORT_EMAIL,
      subject: `🆘 Support - ${assistant_name} - ${username}`,
      html: emailHtml,
    });

    console.log(`✅ Email support envoyé à ${SUPPORT_EMAIL}`);

    return res.json({
      success: true,
      message: "Email de notification envoyé avec succès",
      email_sent_to: SUPPORT_EMAIL,
      timestamp: new Date().toISOString(),
    });
  } catch (err) {
    console.error("❌ Erreur envoi email support:", err && err.stack ? err.stack : String(err));
    return res.status(500).json({
      success: false,
      error: "Erreur lors de l'envoi de l'email",
      details: (err && err.message) ?? String(err),
    });
  }
});

// ========== ROUTE : GÉNÉRATION EXCEL + EMAIL ==========
app.post("/api/generate-excel-and-send", async (req, res) => {
  try {
    const { email, subject, sheets, filename, cc } = req.body || {};

    if (!email || !subject || !sheets) {
      return res.status(400).json({
        success: false,
        error: "Données manquantes",
        details: "Envoyez email, subject, sheets (array ou objet)",
      });
    }

    if (!isValidEmail(email)) {
      return res.status(400).json({ success: false, error: "Email invalide" });
    }

    if (cc && !isValidEmail(cc)) {
      return res.status(400).json({
        success: false,
        error: "Adresse email CC invalide",
        details: `cc = ${cc}`,
      });
    }

    const workbook = XLSX.utils.book_new();

    if (Array.isArray(sheets)) {
      if (sheets.length === 0) {
        return res.status(400).json({ success: false, error: "Le tableau sheets est vide" });
      }

      sheets.forEach((sheet, index) => {
        const sheetName = sheet.name || `Sheet${index + 1}`;
        const sheetData = sheet.data;
        if (!Array.isArray(sheetData)) throw new Error(`Les données du sheet "${sheetName}" doivent être un tableau`);
        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      });
    } else if (typeof sheets === "object" && sheets !== null) {
      const sheetNames = Object.keys(sheets);
      if (sheetNames.length === 0) {
        return res.status(400).json({ success: false, error: "L'objet sheets est vide" });
      }

      sheetNames.forEach((sheetName) => {
        const sheetData = sheets[sheetName];
        if (!Array.isArray(sheetData)) throw new Error(`Les données du sheet "${sheetName}" doivent être un tableau`);
        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      });
    } else {
      return res.status(400).json({ success: false, error: "Format sheets invalide", details: "sheets doit être un array ou un objet" });
    }

    const excelBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx", compression: true });

    const excelFilename = filename
      ? `${String(filename).replace(/[^a-z0-9]/gi, "_")}.xlsx`
      : `rapport_${Date.now()}.xlsx`;

    const sheetCount = workbook.SheetNames.length;
    const sheetsList = workbook.SheetNames
      .map((name, i) => `<li><strong>Sheet ${i + 1}:</strong> ${escapeHtml(name)}</li>`)
      .join("");

    const emailHtml = `
      <!DOCTYPE html>
      <html>
        <body style="font-family: Arial, sans-serif; line-height:1.6; color:#111827;">
          <h2 style="margin:0 0 8px 0;">📊 Votre fichier Excel est prêt</h2>

          <div style="background:#e0f2fe;padding:12px;border-left:4px solid #0ea5e9;border-radius:6px;margin:12px 0;">
            <strong>📧 Sujet :</strong> ${escapeHtml(subject)}<br>
            <strong>📁 Fichier :</strong> ${escapeHtml(excelFilename)}<br>
            <strong>📋 Nombre de feuilles :</strong> ${sheetCount}<br>
            <strong>📅 Date :</strong> ${new Date().toLocaleDateString("fr-FR", {
              year: "numeric",
              month: "long",
              day: "numeric",
              hour: "2-digit",
              minute: "2-digit",
            })}
          </div>

          ${
            sheetCount > 0
              ? `
          <div style="background:#f0fdf4;padding:12px;border-left:4px solid #10b981;border-radius:6px;margin:12px 0;">
            <strong>📑 Feuilles incluses :</strong>
            <ul style="margin:8px 0 0 0;padding-left:20px;">
              ${sheetsList}
            </ul>
          </div>
          `
              : ""
          }

          <p>Vous trouverez le fichier Excel complet en pièce jointe.</p>

          <div style="background:#fef3c7;padding:10px;border-radius:6px;margin:12px 0;font-size:13px;">
            <strong>💡 Astuce :</strong> Ouvrez le fichier avec Microsoft Excel, Google Sheets ou LibreOffice Calc.
          </div>

          <p style="color:#6b7280;font-size:12px;margin-top:20px;">
            © ${new Date().getFullYear()} ${EMAIL_FROM_NAME}
          </p>
        </body>
      </html>
    `;

    await transporter.sendMail({
      from: { name: EMAIL_FROM_NAME, address: EMAIL_FROM },
      to: email,
      cc: cc || undefined,
      subject,
      html: emailHtml,
      attachments: [
        {
          filename: excelFilename,
          content: excelBuffer,
          contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      ],
    });

    console.log(`✅ Excel envoyé à ${email}${cc ? " (CC: " + cc + ")" : ""} - ${excelFilename}`);

    return res.json({
      success: true,
      message: "Fichier Excel généré et envoyé avec succès",
      details: {
        email,
        cc: cc || null,
        filename: excelFilename,
        sheets: workbook.SheetNames,
        sheetCount,
        fileSize: `${(excelBuffer.length / 1024).toFixed(2)} KB`,
      },
    });
  } catch (err) {
    console.error("❌ Erreur génération/envoi Excel:", err && err.stack ? err.stack : String(err));
    return res.status(500).json({
      success: false,
      error: "Erreur lors du traitement",
      details: (err && err.message) ?? String(err),
    });
  }
});

// ========== ROUTE : GÉNÉRATION PDF + EMAIL (REPORT) ==========
app.post("/api/generate-and-send", async (req, res) => {
  try {
    const { email, subject, reportContent } = req.body || {};

    if (!email || !subject || !reportContent) {
      return res.status(400).json({
        success: false,
        error: "Données manquantes",
        details: "Envoyez email, subject, reportContent",
      });
    }
    if (!isValidEmail(email)) {
      return res.status(400).json({ success: false, error: "Email invalide" });
    }
    if (!reportContent.title || !reportContent.introduction || !Array.isArray(reportContent.sections) || !reportContent.conclusion) {
      return res.status(400).json({ success: false, error: "Structure du rapport invalide" });
    }

    const pdfBuffer = await generatePDF(reportContent);
    const pdfName = `rapport_${String(reportContent.title).replace(/[^a-z0-9]/gi, "_").toLowerCase()}_${Date.now()}.pdf`;

    const html = `
      <!DOCTYPE html>
      <html>
        <body style="font-family: Arial, sans-serif; line-height:1.6; color:#111827;">
          <h2 style="margin:0 0 8px 0;">📄 Votre rapport est prêt</h2>
          <div style="background:#e0e7ff;padding:12px;border-left:4px solid #667eea;border-radius:6px;margin:12px 0;">
            <strong>📊 Sujet :</strong> ${escapeHtml(subject)}<br>
            <strong>📌 Titre :</strong> ${escapeHtml(reportContent.title)}<br>
            <strong>📅 Date :</strong> ${new Date().toLocaleDateString("fr-FR")}
          </div>
          <p>Vous trouverez le rapport complet en pièce jointe au format PDF.</p>
          <p style="color:#6b7280;font-size:12px">© ${new Date().getFullYear()} ${EMAIL_FROM_NAME}</p>
        </body>
      </html>
    `;

    await sendEmailWithPdf({
      to: email,
      subject: `Rapport : ${reportContent.title}`,
      messageHtml: html,
      pdfBuffer,
      pdfFilename: pdfName,
    });

    return res.json({
      success: true,
      message: "Rapport généré et envoyé avec succès",
      details: {
        email,
        pdfSize: `${(pdfBuffer.length / 1024).toFixed(2)} KB`,
      },
    });
  } catch (err) {
    console.error("❌ Erreur:", err && err.stack ? err.stack : String(err));
    return res.status(500).json({
      success: false,
      error: "Erreur lors du traitement",
      details: (err && err.message) ?? String(err),
    });
  }
});

// Healthcheck
app.get("/health", (_req, res) => {
  res.json({
    status: "OK",
    timestamp: new Date().toISOString(),
    uptime: Math.floor(process.uptime()),
    service: "PDF / Excel Report & Support API",
  });
});

// Root : documentation rapide
app.get("/", (_req, res) => {
  res.json({
    name: "GPT PDF / Excel Email & Support API",
    version: "2.2.0",
    status: "running",
    endpoints: {
      health: "GET /health",
      echo: "POST /api/echo",
      generateOfferAndSend: "POST /api/generate-offer-and-send",
      generateAndSendPdf: "POST /api/generate-and-send",
      generateExcelAndSend: "POST /api/generate-excel-and-send",
      sendSupportEmail: "POST /api/support/send-email",
      static: "GET /static/<fichier>",
    },
  });
});

/* ========================= 404 & ERREUR ======================== */
app.use((req, res) => res.status(404).json({ error: "Route non trouvée", path: req.path }));

app.use((err, _req, res, _next) => {
  console.error("Erreur middleware:", err && err.stack ? err.stack : String(err));
  res.status(500).json({ error: "Erreur serveur", message: (err && err.message) ?? String(err) });
});

/* ============================ START ============================ */
const PORT = process.env.PORT || 3000;
app.listen(PORT, "0.0.0.0", () => {
  console.log(`🚀 API démarrée sur port ${PORT}`);
  console.log(`📧 Email support configuré vers: ${SUPPORT_EMAIL}`);
});

/* ========================= PROCESS HOOKS ======================= */
process.on("unhandledRejection", (r) => console.error("Unhandled Rejection:", r && r.stack ? r.stack : String(r)));
process.on("uncaughtException", (e) => {
  console.error("Uncaught Exception:", e && e.stack ? e.stack : String(e));
  process.exit(1);
});
