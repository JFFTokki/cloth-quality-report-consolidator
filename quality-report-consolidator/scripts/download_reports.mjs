import fs from "node:fs/promises";
import path from "node:path";

const [manifestPath, outputDir = "pdfs", resultsPath = "download_results.json"] = process.argv.slice(2);
if (!manifestPath) throw new Error("Usage: node download_reports.mjs manifest.json [pdf_dir] [results.json]");

const reports = JSON.parse(await fs.readFile(manifestPath, "utf8"));
await fs.mkdir(outputDir, { recursive: true });
const safe = (value) => String(value || "report.pdf").replace(/[\\/:*?\"<>|]/g, "_");
const results = new Array(reports.length);
let cursor = 0;

async function isPdf(filePath) {
  try {
    const file = await fs.open(filePath, "r");
    const bytes = Buffer.alloc(5);
    await file.read(bytes, 0, 5, 0);
    await file.close();
    return bytes.toString("latin1") === "%PDF-";
  } catch {
    return false;
  }
}

async function worker() {
  while (cursor < reports.length) {
    const index = cursor++;
    const report = reports[index];
    const sourceName = report.fileName || `report_${index + 1}.pdf`;
    const fileName = `${String(index + 1).padStart(5, "0")}_${safe(sourceName)}`;
    const localPath = path.join(outputDir, fileName);
    try {
      if (await isPdf(localPath)) {
        const stat = await fs.stat(localPath);
        results[index] = { ...report, index: index + 1, localPath, bytes: stat.size, status: "ok", cached: true };
        continue;
      }
      const response = await fetch(encodeURI(report.url), {
        redirect: "follow",
        headers: { "User-Agent": "Mozilla/5.0" },
      });
      const bytes = new Uint8Array(await response.arrayBuffer());
      const signature = new TextDecoder("latin1").decode(bytes.slice(0, 5));
      if (!response.ok || signature !== "%PDF-") throw new Error(`HTTP ${response.status}; signature=${signature}`);
      await fs.writeFile(localPath, bytes);
      results[index] = { ...report, index: index + 1, localPath, bytes: bytes.length, status: "ok", cached: false };
    } catch (error) {
      results[index] = { ...report, index: index + 1, localPath, status: "error", error: String(error) };
    }
  }
}

await Promise.all(Array.from({ length: 4 }, () => worker()));
await fs.writeFile(resultsPath, JSON.stringify(results, null, 2), "utf8");
const failed = results.filter((row) => row.status !== "ok");
console.log(JSON.stringify({ total: results.length, ok: results.length - failed.length, failed: failed.length }, null, 2));
if (failed.length) process.exitCode = 2;
