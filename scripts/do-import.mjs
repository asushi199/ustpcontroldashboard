import fs from "fs";
const p = "src/App.jsx";
let s = fs.readFileSync(p, "utf8");
const a = "} from \"./lib/embedUrls.js\";";
const b = a + "\nimport { fetchAnalisisCsvText } from \"./lib/analisisSheetFetch.js\";\nimport {\n  parseDelimaAnalisisCsv,\n  parseDcsAnalisisCsv,\n  parseAinsAnalisisCsv,\n  parsePensijilanAnalisisCsv,\n  parseOptikAnalisisCsv,\n} from \"./lib/analisisSheetParse.js\";";
if (s.includes("analisisSheetFetch")) process.exit(0);
if (!s.includes(a)) throw new Error("no anchor");
fs.writeFileSync(p, s.replace(a, b));
console.log("import done");
