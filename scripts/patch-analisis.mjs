import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const appPath = path.join(__dirname, "..", "src", "App.jsx");
const frag = (n) => fs.readFileSync(path.join(__dirname, "fragments", n), "utf8").replace(/\r\n/g, "\n");
let s = fs.readFileSync(appPath, "utf8").replace(/\r\n/g, "\n");
function rep(a,b){ if(!s.includes(a)) throw new Error("miss"); s=s.replace(a,b); }
rep(`} from "./lib/embedUrls.js";`,`} from "./lib/embedUrls.js";` + String.fromCharCode(10) + `import { fetchAnalisisCsvText } from "./lib/analisisSheetFetch.js";` + String.fromCharCode(10) + `import { parseDelimaAnalisisCsv, parseDcsAnalisisCsv, parseAinsAnalisisCsv, parsePensijilanAnalisisCsv, parseOptikAnalisisCsv, } from "./lib/analisisSheetParse.js";`);
fs.writeFileSync(appPath,s,"utf8");
console.log("step1");
