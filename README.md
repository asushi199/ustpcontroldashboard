# ustp-dashboard

面向 **USTP Daerah Manjung**（PPD Manjung）的聚合看板：在同一页面中嵌入 Google Calendar、Looker Studio 报表、Canva 等内容，并在前端解析 **CSV / Excel** 展示 Program Ains、DELIMa 等本地数据与图表。

## 技术栈

- [React 19](https://react.dev/) + [Vite 8](https://vite.dev/)
- [Tailwind CSS 3](https://tailwindcss.com/)
- [SheetJS (xlsx)](https://sheetjs.com/) — 读取 `public` 下的 Excel

## 环境要求

- **Node.js** 建议当前 LTS（与 Vite 8 兼容即可）

## 快速开始

```bash
npm install
npm run dev
```

本地开发默认由 Vite 提供地址（一般为 `http://localhost:5173`）。

### 其他脚本

| 命令 | 说明 |
|------|------|
| `npm run build` | 生产构建，输出到 `dist/` |
| `npm run preview` | 本地预览构建结果 |
| `npm run lint` | 运行 ESLint |

## 数据与静态资源

应用通过 **固定路径** 拉取本地文件（见 `src/App.jsx` 顶部常量）。部署或更新数据时，将文件放到 `public/` 下对应位置即可（无需改 import）。

| 路径 | 用途 |
|------|------|
| `public/data/users_PERAK.csv` | Program Ains — 按 CSV 列名解析，并筛选与 Manjung 相关的行 |
| `public/data/delima-ppd-manjung.xlsx` | DELIMa — 由前端用 `xlsx` 解析；更新数据时**直接替换该文件** |
| `public/assets/ustp-ppd-manjung-watermark.png` | 页面水印 / 品牌图（若缺失则需在仓库中补充该目录与文件） |

> **注意**：若仓库中未包含 `.xlsx` 或部分 `assets`（体积或版权原因），克隆后需自行放入上述路径，否则对应区块可能无数据或图片裂图。

嵌入的 **Google 表格 / Looker / Calendar** 等链接也集中在 `App.jsx` 常量中；更换报表或日历时，优先改代码内 URL，而不是改构建配置。

## 项目结构（简要）

```
ustp-dashboard/
├── public/           # 静态资源与数据（CSV/XLSX/图片等）
├── src/
│   ├── App.jsx       # 主界面与数据逻辑（体量较大）
│   ├── main.jsx
│   └── index.css     # 全局样式 + Tailwind 入口
├── index.html
├── vite.config.js
├── tailwind.config.js
└── postcss.config.js
```

## Cursor / 协作

- 项目规则：`.cursor/rules/project.mdc`（`alwaysApply`）
- 索引排除：`.cursorignore`（如忽略 `node_modules/`、构建产物、`.env` 等）

## 许可与声明

仓库为 `private` 包名配置；对外分发、数据使用与嵌入报表的版权/合规需遵循贵单位与 Google / Canva 等平台条款。
