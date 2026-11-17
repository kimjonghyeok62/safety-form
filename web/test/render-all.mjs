// 예시 데이터로 모든 버튼을 눌러 본 것처럼 hwpx 를 만든다.
//   node test/render-all.mjs <출력폴더>
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { Exporter } from "../src/export.js";
import { refreshAuto } from "../src/domain/dictionary.js";
import { toDay } from "../src/domain/schedule.js";

const here = dirname(fileURLToPath(import.meta.url));
const outDir = process.argv[2] || join(here, "out");
await mkdir(outDir, { recursive: true });

const data = JSON.parse(await readFile(join(here, "fixtures/sample.json"), "utf-8"));
data.contracts.forEach(refreshAuto);

const ex = new Exporter((n) => readFile(join(here, "../public/templates", n)), toDay("2026-08-10"));
const runs = [
  await ex.forms17(data, data.contracts),
  await ex.council(data, data.contracts),
  await ex.patrol(data, data.contracts),
  await ex.joint(data, data.contracts),
  ...(await Promise.all([11, 12, 13, 14, 15].map((n) => ex.single(data, [data.contracts[1]], n)))),
];
let n = 0;
for (const r of runs) {
  for (const f of r.files) { await writeFile(join(outDir, f.name), f.bytes); n++; }
  for (const note of r.notes) console.log("note:", note);
}
for (const c of data.contracts)
  console.log(c.계약명, "|", c.계약유형.v, "|", Object.entries(c.work).filter(([, v]) => v.v).map(([k]) => k).join(","));
console.log(`files: ${n} → ${outDir}`);
