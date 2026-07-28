import fs from "node:fs/promises";
import { createReadStream } from "node:fs";
import crypto from "node:crypto";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const [manifestPath, outputDir = "pdfs", resultsPath = "download_results.json"] = process.argv.slice(2);
if (!manifestPath) throw new Error("Usage: node download_reports.mjs manifest.json [pdf_dir] [results.json]");

const reports = JSON.parse(await fs.readFile(manifestPath, "utf8"));
const maxBytes = Number(process.env.QC_MAX_PDF_BYTES || 100 * 1024 * 1024);
const timeoutMs = Number(process.env.QC_DOWNLOAD_TIMEOUT_MS || 90_000);
const workerCount = Number(process.env.QC_DOWNLOAD_WORKERS || 4);
const pdfinfo = process.env.QC_PDFINFO || "pdfinfo";
const execFileAsync = promisify(execFile);
await fs.mkdir(outputDir, { recursive: true });
const safe = (value) => String(value || "report.pdf").replace(/[\\/:*?"<>|]/g, "_");
const results = new Array(reports.length);
let cursor = 0;

function checkedUrl(raw) {
  const parsed = new URL(raw);
  if (!["http:", "https:"].includes(parsed.protocol)) throw new Error("only http/https URLs are allowed");
  return parsed;
}

async function sha256File(filePath) {
  const hash = crypto.createHash("sha256");
  for await (const chunk of createReadStream(filePath)) hash.update(chunk);
  return hash.digest("hex");
}

async function inspectPdf(filePath) {
  const file = await fs.open(filePath, "r");
  try {
    const bytes = Buffer.alloc(5);
    await file.read(bytes, 0, 5, 0);
    if (bytes.toString("latin1") !== "%PDF-") throw new Error("file does not start with %PDF-");
    const stat = await file.stat();
    if (stat.size <= 1024) throw new Error(`PDF too small: ${stat.size} bytes`);
    if (stat.size > maxBytes) throw new Error(`PDF exceeds size limit: ${stat.size} > ${maxBytes}`);
    await execFileAsync(pdfinfo, [filePath], { timeout: timeoutMs, maxBuffer: 1024 * 1024 });
    return { bytes: stat.size, sha256: await sha256File(filePath), pdfOpenVerified: true };
  } finally {
    await file.close();
  }
}

async function downloadPdf(report, localPath) {
  const sourceUrl = checkedUrl(report.url);
  const tempPath = `${localPath}.part-${process.pid}-${crypto.randomBytes(4).toString("hex")}`;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(new Error(`download timeout after ${timeoutMs}ms`)), timeoutMs);
  let file;
  try {
    const response = await fetch(encodeURI(sourceUrl.href), {
      redirect: "follow",
      headers: { "User-Agent": "Mozilla/5.0" },
      signal: controller.signal,
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    checkedUrl(response.url);
    const contentLength = Number(response.headers.get("content-length") || 0);
    if (contentLength > maxBytes) throw new Error(`PDF exceeds size limit: ${contentLength} > ${maxBytes}`);
    if (!response.body) throw new Error("response body is empty");

    file = await fs.open(tempPath, "wx");
    const hash = crypto.createHash("sha256");
    const reader = response.body.getReader();
    let total = 0;
    let prefix = Buffer.alloc(0);
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = Buffer.from(value);
      total += chunk.length;
      if (total > maxBytes) throw new Error(`PDF exceeds size limit: ${total} > ${maxBytes}`);
      if (prefix.length < 5) prefix = Buffer.concat([prefix, chunk]).subarray(0, 5);
      hash.update(chunk);
      await file.write(chunk);
    }
    if (prefix.toString("latin1") !== "%PDF-") throw new Error("response does not start with %PDF-");
    if (total <= 1024) throw new Error(`PDF too small: ${total} bytes`);
    await file.sync();
    await file.close();
    file = undefined;
    await execFileAsync(pdfinfo, [tempPath], { timeout: timeoutMs, maxBuffer: 1024 * 1024 });
    await fs.rename(tempPath, localPath);
    return {
      bytes: total,
      sha256: hash.digest("hex"),
      finalUrl: response.url,
      etag: response.headers.get("etag") || "",
      lastModified: response.headers.get("last-modified") || "",
      pdfOpenVerified: true,
    };
  } finally {
    clearTimeout(timer);
    if (file) await file.close().catch(() => {});
    await fs.unlink(tempPath).catch(() => {});
  }
}

async function worker() {
  while (cursor < reports.length) {
    const index = cursor++;
    const report = reports[index];
    const sourceName = report.fileName || `report_${index + 1}.pdf`;
    let localPath = "";
    try {
      const urlKey = crypto.createHash("sha256").update(checkedUrl(report.url).href).digest("hex").slice(0, 20);
      const fileName = `${urlKey}_${safe(sourceName)}`;
      localPath = path.join(outputDir, fileName);
      let metadata;
      let cached = false;
      try {
        metadata = await inspectPdf(localPath);
        cached = true;
      } catch {
        metadata = await downloadPdf(report, localPath);
      }
      results[index] = {
        ...report,
        index: index + 1,
        localPath,
        status: "ok",
        cached,
        downloadVersion: "generic-download-v2",
        ...metadata,
      };
    } catch (error) {
      results[index] = { ...report, index: index + 1, localPath, status: "error", error: String(error) };
    }
  }
}

await Promise.all(Array.from({ length: workerCount }, () => worker()));
const tempResults = `${resultsPath}.tmp`;
await fs.writeFile(tempResults, JSON.stringify(results, null, 2), "utf8");
await fs.rename(tempResults, resultsPath);
const failed = results.filter((row) => row.status !== "ok");
console.log(JSON.stringify({ total: results.length, ok: results.length - failed.length, failed: failed.length }, null, 2));
if (failed.length) process.exitCode = 2;
