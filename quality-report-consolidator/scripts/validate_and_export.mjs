import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const args = process.argv.slice(2);
const allowPartialIndex = args.indexOf("--allow-partial");
const allowPartial = allowPartialIndex >= 0;
if (allowPartial) args.splice(allowPartialIndex, 1);
const [recordsPath, downloadsPath, outputPath] = args;
if (!recordsPath || !downloadsPath || !outputPath) {
  throw new Error("Usage: node validate_and_export.mjs report_data.json download_manifest.json output.xlsx [--allow-partial]");
}

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const validationPath = outputPath.replace(/\.xlsx$/i, ".validation.json");
const exportDiagnosticsPath = outputPath.replace(/\.xlsx$/i, ".export.json");
const reviewOutputPath = outputPath.replace(/\.xlsx$/i, "_external-review-required.xlsx");
const reviewValidationPath = reviewOutputPath.replace(/\.xlsx$/i, ".validation.json");
const reviewDiagnosticsPath = reviewOutputPath.replace(/\.xlsx$/i, ".export.json");
const overviewPreviewPath = outputPath.replace(/\.xlsx$/i, "_overview_preview.png");
const detailPreviewPath = outputPath.replace(/\.xlsx$/i, "_detail_preview.png");
const reviewOverviewPreviewPath = reviewOutputPath.replace(/\.xlsx$/i, "_overview_preview.png");
const reviewDetailPreviewPath = reviewOutputPath.replace(/\.xlsx$/i, "_detail_preview.png");
for (const reservedPath of [
  outputPath, validationPath, exportDiagnosticsPath,
  reviewOutputPath, reviewValidationPath, reviewDiagnosticsPath,
  overviewPreviewPath, detailPreviewPath, reviewOverviewPreviewPath, reviewDetailPreviewPath,
]) {
  if (fs.existsSync(reservedPath)) throw new Error(`Refusing to overwrite existing output: ${reservedPath}`);
}
const python = process.env.QC_PYTHON || "python3";
const validationArgs = [
  path.join(scriptDir, "validate_records.py"),
  recordsPath,
  "--downloads",
  downloadsPath,
  "--output",
  validationPath,
];
if (allowPartial) validationArgs.push("--allow-partial");
const validation = spawnSync(python, validationArgs, { encoding: "utf8" });
if (validation.stdout) process.stdout.write(validation.stdout);
if (validation.stderr) process.stderr.write(validation.stderr);
if (validation.status !== 0) process.exit(validation.status || 2);

const result = JSON.parse(fs.readFileSync(validationPath, "utf8"));
if (result.validationStatus !== "完整通过" && !(allowPartial && result.validationStatus === "部分完成")) {
  throw new Error(`Export blocked by validation status: ${result.validationStatus}`);
}
if (result.validationStatus === "部分完成" && !/(部分|partial|preview|预览|验证)/i.test(path.basename(outputPath))) {
  throw new Error("Partial export filename must visibly include 部分, partial, preview, 预览, or 验证");
}

const exporter = spawnSync(
  process.execPath,
  [path.join(scriptDir, "export_leader_workbook.mjs"), recordsPath, outputPath],
  { encoding: "utf8" },
);
if (exporter.stdout) process.stdout.write(exporter.stdout);
if (exporter.stderr) process.stderr.write(exporter.stderr);
if (exporter.status !== 0) process.exit(exporter.status || 2);
const exportDiagnostics = JSON.parse(fs.readFileSync(exportDiagnosticsPath, "utf8"));
if (exportDiagnostics.formulaErrorCount || exportDiagnostics.visualPreviewStatus !== "passed") {
  fs.renameSync(outputPath, reviewOutputPath);
  fs.renameSync(validationPath, reviewValidationPath);
  fs.renameSync(exportDiagnosticsPath, reviewDiagnosticsPath);
  if (fs.existsSync(overviewPreviewPath)) fs.renameSync(overviewPreviewPath, reviewOverviewPreviewPath);
  if (fs.existsSync(detailPreviewPath)) fs.renameSync(detailPreviewPath, reviewDetailPreviewPath);
  console.error(`Formal export blocked; review artifact retained at ${reviewOutputPath}`);
  process.exit(3);
}
