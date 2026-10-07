// Bundles src/app.js (+ d3-selection, d3-drag, d3-sankey) and inlines it into template.html -> ../SankeyChart.html
import { createRequire } from "module";
import { readFileSync, writeFileSync } from "fs";
import { fileURLToPath } from "url";
import path from "path";

const here = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(path.join(here, "../SankeyFlow/package.json"));
const esbuild = require("esbuild");

const result = await esbuild.build({
  entryPoints: [path.join(here, "src/app.js")],
  bundle: true, minify: true, format: "iife", target: "es2022", write: false,
  nodePaths: [path.join(here, "../SankeyFlow/node_modules")]
});
const js = result.outputFiles[0].text.replace(/<\/script/gi, "<\\/script");
const html = readFileSync(path.join(here, "template.html"), "utf8").replace("/*__BUNDLE__*/", () => js);
const out = path.join(here, "../SankeyChart.html");
writeFileSync(out, html);
console.log("written", out, (html.length / 1024).toFixed(0) + " KB");
