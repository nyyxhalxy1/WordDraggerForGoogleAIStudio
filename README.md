# DOCX to Google AI Studio

Chrome / Edge 浏览器扩展，让你可以直接拖拽 Word (.docx) 文件到 Google AI Studio，自动转换为 Markdown 格式上传。

## 功能特性

- **拖放上传**：直接将 `.docx` 文件拖到 Google AI Studio 页面
- **文件选择器支持**：通过 AI Studio 的文件上传按钮选择 `.docx` 文件
- **智能转换**：保留标题、列表、表格、加粗/斜体等格式
- **表格支持**：使用 GFM 表格语法保留 Word 中的表格结构
- **图片处理**：
  - 无图片文档 → 直接转换为 `.md` 文件上传
  - 含图片文档 → 提供两种选择：
    - **仅上传文本**：丢弃图片，只上传 Markdown 文件（节省文件配额）
    - **上传文本 + 图片**：上传 Markdown 文件 + 所有图片文件

## 安装方法

### 开发者模式加载（推荐本地使用）

1. 下载或克隆本项目到本地
2. 打开 Chrome 或 Edge 浏览器，进入扩展管理页面：
   - Chrome: `chrome://extensions/`
   - Edge: `edge://extensions/`
3. 开启的「开发者模式」
4. 点击「加载已解压的扩展程序」
5. 选择本项目根目录（包含 `manifest.json` 的文件夹）
6. 扩展图标将出现在浏览器工具栏

### Chrome Web Store 安装

*即将上架...*

## 使用方法

1. 打开 [Google AI Studio](https://aistudio.google.com/)
2. 将 `.docx` 文件拖到页面上，或通过 AI Studio 的文件上传按钮选择 `.docx` 文件
3. 扩展会自动拦截并转换文件：
   - 如果文档不含图片，会直接转换为 Markdown 并上传
   - 如果文档包含图片，会弹出选择对话框，选择处理方式
4. 转换完成后，文件会自动提交给 AI Studio

## 技术架构

```
manifest.json          # 扩展配置 (Manifest V3)
content.js             # 内容脚本 (隔离世界) — 拦截拖放事件、解析 docx、UI 交互
inject.js              # 注入脚本 (主世界) — 触发 AI Studio 原生文件上传
dialog.css             # 对话框和 UI 样式
lib/
  ├── jszip.min.js            # 解压 .docx (ZIP) 提取图片
  ├── mammoth.browser.min.js  # .docx → HTML 转换
  ├── turndown.min.js         # HTML → Markdown 转换
  └── turndown-plugin-gfm.min.js  # GFM 表格/删除线支持
icons/
  ├── icon16.png
  ├── icon48.png
  └── icon128.png
```

### 工作流程

```
用户拖入 .docx 文件
        ↓
content.js 捕获阶段拦截 drop 事件
        ↓
读取文件 ArrayBuffer
        ↓
  ┌─────────────┐    ┌─────────────┐
  │ mammoth.js  │    │   JSZip     │
  │ docx → HTML │    │  提取图片    │
  └──────┬──────┘    └──────┬──────┘
         ↓                  ↓
  Turndown: HTML → MD       ↓
         ↓                  ↓
    ┌────┴──── 图片数量 ─────┘
    │
    ├─ 0 张 → 直接上传 .md
    │
    └─ N 张 → 弹出对话框
                ├─ 仅文本 → 上传 .md
                └─ 文本+图片 → 上传 .md + N 张图片
                        ↓
              通过 postMessage 发送到 inject.js
                        ↓
              inject.js 构造 DataTransfer
              触发 AI Studio 原生上传
```

## 依赖库

| 库 | 版本 | 用途 | 许可证 |
|---|---|---|---|
| [JSZip](https://stuk.github.io/jszip/) | 3.10.1 | 解压 .docx ZIP 包 | MIT |
| [mammoth.js](https://github.com/mwilliamson/mammoth.js) | 1.8.0 | .docx → HTML 转换 | BSD-2-Clause |
| [Turndown](https://github.com/mixmark-io/turndown) | 7.2.0 | HTML → Markdown | MIT |
| [turndown-plugin-gfm](https://github.com/mixmark-io/turndown-plugin-gfm) | 1.0.2 | GFM 表格/删除线支持 | MIT |

## 已知限制

- **不支持的图片格式**：`.emf` / `.wmf` 格式（Windows 矢量图）无法在浏览器中处理，会被自动跳过
- **复杂排版**：高度复杂的 Word 排版（如分栏、文本框、艺术字等）可能无法完美保留
- **页面更新**：Google AI Studio 更新页面结构后，上传触发机制可能需要适配调整
- **大文件**：超大 Word 文档（> 50MB）处理时间较长，请耐心等待

## 隐私说明

- 所有文件处理均在浏览器本地完成，不会上传到任何第三方服务器
- 扩展仅在 `aistudio.google.com` 域名下激活
- 不收集、存储或传输任何用户数据

## 许可证

MIT License
