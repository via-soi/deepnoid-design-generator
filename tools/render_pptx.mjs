// pptx → 슬라이드 PNG. 사용법: node tools/render_pptx.mjs <pptx> <출력디렉터리>
import { pdf } from "pdf-to-img";
import { execFileSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

const SOFFICE = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

async function main() {
  const [src, outDir] = process.argv.slice(2);
  if (!src || !outDir) { console.error("사용법: node render_pptx.mjs <pptx> <outDir>"); process.exit(1); }
  await fs.mkdir(outDir, { recursive: true });
  const pdfDir = path.resolve(outDir);
  execFileSync(SOFFICE, ["--headless", "--convert-to", "pdf", "--outdir", pdfDir, path.resolve(src)]);
  const pdfPath = path.join(pdfDir, path.basename(src).replace(/\.pptx$/i, ".pdf"));
  const doc = await pdf(pdfPath, { scale: 2 });
  const pad = String(doc.length).length;
  let i = 1;
  for await (const img of doc) {
    await fs.writeFile(path.join(outDir, `slide-${String(i).padStart(pad, "0")}.png`), img);
    i++;
  }
  await fs.unlink(pdfPath);
  console.log(`렌더 완료: ${i - 1}장 → ${outDir}`);
}
main();
