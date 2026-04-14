# USTP Dashboard（PPD Manjung）

面向 **USTP Daerah Manjung** 的聚合看板：在同一页面中嵌入 Google Calendar、Looker Studio、Canva、Drive PDF 等，并支持从 **Google 表格（CSV 导出）** 与本地 **CSV / Excel / JSON** 驱动部分区块内容。

## 技术栈

- [React 19](https://react.dev/) + [Vite 8](https://vite.dev/)
- [Tailwind CSS 3](https://tailwindcss.com/)
- [SheetJS (xlsx)](https://sheetjs.com/) — 读取 `public` 下的 Excel（如学校网站列表）

## 环境要求

- **Node.js** 建议使用当前 LTS（与 Vite 8 兼容即可）

## 快速开始

```bash
npm install
cp .env.example .env   # 按需填写；可不填先跑本地 demo
npm run dev
```

本地开发地址一般为 `http://localhost:5173`。

### 常用脚本

| 命令 | 说明 |
|------|------|
| `npm run dev` | 开发服务器 |
| `npm run build` | 生产构建，输出到 `dist/` |
| `npm run preview` | 本地预览构建结果 |
| `npm run lint` | 运行 ESLint |

### 部署

构建产物在 `dist/`。若使用 **Netlify**，仓库中的 `netlify.toml` 与 `netlify/functions/sheet-csv.mjs` 会提供 `/api/sheet-csv` 代理，避免浏览器直接请求 Google 表格时出现 CORS 问题。部署前在 Netlify 环境变量中配置与本地相同的 `VITE_*` 变量，然后执行构建即可。

**请勿**将含真实密钥的 `.env` 提交到 Git；仅保留 `.env.example` 中的占位说明。

---

## 环境变量说明（`.env`）

复制 `.env.example` 为 `.env`，按模块填写。所有前端可见变量必须以 `VITE_` 开头。

### Google 表格总表 ID

| 变量 | 说明 |
|------|------|
| `VITE_GOOGLE_SHEET_ID` | 电子表格 ID（URL 中 `/d/` 与 `/edit` 之间的字符串）。留空时，部分功能使用 `public/data/` 下的 demo 文件。 |
| `VITE_GOOGLE_SHEET_GID` | **主 OSC 表**的标签页 ID（URL 中 `gid=`）。用于带 `section` 列的卡片网格（如 Bahan Sokongan 静态区里的「Contoh Bahan Delima」等）。 |

### OSC 大主题：一主题一标签页（subtopik + 卡片）

五个主题各自对应**同一张** `VITE_GOOGLE_SHEET_ID` 下的**不同 tab**。在表格中新建标签页，从 URL 复制 `gid`，填入下表对应变量。**不需要**在这些 tab 里写 `section` 列。

| 变量 | 界面区块 |
|------|-----------|
| `VITE_OSC_GID_INTEGRASI` | Integrasi Teknologi Pendidikan |
| `VITE_OSC_GID_HEBAHAN` | Hebahan Pendidikan Digital（COE） |
| `VITE_OSC_GID_ITM` | Inisiatif Teknologi Maklumat |
| `VITE_OSC_GID_PEMBUDAYAAN_MEMBACA` | Pembudayaan Amalan Membaca |
| `VITE_OSC_GID_PEMERKASAAN` | Program Pemerkasaan Bacaan Murid |

| `VITE_OSC_GID_LAMAN_WEB_SEKOLAH` | （可选）**ITM「Laman web sekolah」** 学校名单/网址补丁表，与 `public/data/laman-web-sekolah-bengkel-responses.xlsx` **按学校代码合并**：Sheet 中有 `url` 时覆盖 Excel 中的链接；代码在 `App.jsx` 的 `LAMAN_WEB_FEATURED` 仍提供 8 所默认展示顺序与后备网址。未配置 `VITE_GOOGLE_SHEET_ID` 时会加载 `public/data/laman-web-sekolah-sheet-demo.csv` 作为结构示例。 |

**表头建议（与 demo CSV 一致）：**

`subtopik_key, subtopik_title, subtopik_sort, subtopik_blurb, sort, key, title, url, type, blurb`

- 每一行是一张**卡片**；相同 `subtopik_key` 的行会归到同一 **subtopik** 折叠区块内。
- `type` 常用值：`pdf`、`canva`、`gdoc`、`embed`（含 Looker）、`youtube`、`image`（见 `src/lib/oscSheetCsv.js` 中的解析与别名）。

**行为摘要：**

- 已配置 `VITE_GOOGLE_SHEET_ID` **且** 该主题对应的 `VITE_OSC_GID_*` 非空：从该 tab 动态渲染 subtopik + 卡片。
- 已配置 `VITE_GOOGLE_SHEET_ID` 但某一主题的 GID 留空：该主题保留**代码内静态**界面。
- **未**配置 `VITE_GOOGLE_SHEET_ID`：五个主题会从 `public/data/osc-topik-*-demo.csv` 加载**示范结构**（便于新人预览版式）；上线前请配置真实表格 ID 与各 tab 的 `gid`。

示范文件位置：

- `public/data/osc-topik-integrasi-demo.csv`
- `public/data/osc-topik-hebahan-demo.csv`
- `public/data/osc-topik-itm-demo.csv`
- `public/data/osc-topik-pembudayaan-membaca-demo.csv`
- `public/data/osc-topik-pemerkasaan-demo.csv`

可将这些文件导入 Google 表格后，再按需改列内容与 `gid` 绑定。

### Bahan Sokongan（独立标签页）

| 变量 | 说明 |
|------|------|
| `VITE_BAHAN_SOKONGAN_GID` | 「Bahan Sokongan」整块数据来源 tab。列与上表相同；若表中出现 `page` 列且部分行有值，则只显示 `page=bahan` 的行（与主表混用时）。未配置时显示代码内静态 Bahan 区。 |

Demo：`public/data/bahan-sokongan-demo.csv`。

### Analisis Data（各模块独立 tab）

| 变量 | 说明 |
|------|------|
| `VITE_ANALISIS_GID_DELIMA` | DELIMa 分析 |
| `VITE_ANALISIS_GID_DCS` | DCS |
| `VITE_ANALISIS_GID_AINS` | Program Ains |
| `VITE_ANALISIS_GID_PENSIJILAN` | Pensijilan |
| `VITE_ANALISIS_GID_OPTIK` | Optik |

未配置时使用 `public/data/analisis-*-demo.csv` 等本地 demo（若代码中有对应逻辑）。

### Maklumat Asas

| 变量 | 说明 |
|------|------|
| `VITE_MAKLUMAT_ASAS_GID` | Carta、PKG、Takwim、pegawai 等；未配置可用 `public/data/maklumat-asas-demo.csv`。 |

---

## 本地如何拉取 Google 表格 CSV

- **开发**：若已设置 `VITE_GOOGLE_SHEET_ID`，Vite 会将 `/api/sheet-csv` 代理到 Google 的 CSV 导出地址（见 `vite.config.js`）。
- **生产（Netlify）**：由 `netlify/functions/sheet-csv.mjs` 转发请求。

表格需设为「知道链接的人可查看」或等价权限，否则导出会失败。

---

## 其他本地数据与静态资源

应用常通过 **固定 URL 路径** 读取 `public/` 下文件（更新数据时替换文件并重新 `npm run build` 部署即可）。

| 路径 | 用途 |
|------|------|
| `public/data/users_PERAK.csv` | Program Ains 等 — 按列名解析 |
| `public/data/laman-web-sekolah-bengkel-responses.xlsx` | ITM「Laman web sekolah」— Google 表单导出格式（列如 `KOD SEKOLAH`、`NAMA SEKOLAH`、`PAUTAN LAMAN WEB…`）；区块**始终显示**（即使 ITM 其他内容已由 Google Sheet 动态加载） |
| `public/data/laman-web-sekolah-sheet-demo.csv` | 上述补丁表的列表示例（`code,name,url`） |
| `public/data/opr-amalan-membaca-manjung.json` | Pembudayaan Membaca — OPR Manjung |
| `public/data/osc-sheet-template.csv` | 主 OSC 表带 `section` 的卡片示例（与 `VITE_GOOGLE_SHEET_GID` 主 tab 用法相关） |
| `public/assets/`、`public/*.png` 等 | 水印、hebahan 图、ITM 预览图等 |

**预设 8 所「主打」学校**（无搜索时优先展示）的代码与固定链接在 `src/App.jsx` 的 `LAMAN_WEB_FEATURED` 中；修改需改代码后重新构建部署。

大量嵌入的 Calendar、Looker、Canva 等 URL 也分布在 `src/App.jsx` 与 `src/lib/bahanSokonganConstants.js` 等文件中；替换资源时优先改这些常量或改为由 Sheet 驱动（若该区块已接表格）。

---

## 项目结构（简要）

```
ustp-dashboard/
├── public/                 # 静态资源与数据（CSV/XLSX/JSON/图片）
├── netlify/
│   └── functions/
│       └── sheet-csv.mjs   # 生产环境 Google Sheet CSV 代理
├── src/
│   ├── App.jsx             # 主界面与大量业务逻辑
│   ├── components/         # 可复用区块（如 Bahan Sokongan、OSC 卡片网格）
│   ├── lib/                # Sheet 拉取、CSV 解析、分组等
│   ├── main.jsx
│   └── index.css
├── .env.example            # 环境变量模板（勿提交真实 .env）
├── vite.config.js
├── tailwind.config.js
└── postcss.config.js
```

## Cursor / 协作

- 项目规则：`.cursor/rules/project.mdc`
- 索引排除：`.cursorignore`

## 许可与声明

仓库为 `private` 配置；对外分发、数据使用及嵌入内容的版权与合规须遵循贵单位与 Google、Canva 等平台条款。
