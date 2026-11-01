# 更新日志

## v1.3.0 — 支持腾讯文档在线表格

### 新增

- 支持 `docs.qq.com/sheet/...` 腾讯文档在线表格的扫描与导出（TSV / CSV / JSON / 剪贴板）。
- 腾讯文档适配层：
  - 工作表列表：`SpreadsheetApp.workbook.worksheetManager.getSheetList()`
  - 有效范围：`sheet.cellDataGrid.usedRange`
  - 单元格：`sheet.getCellDataAtPosition(row, col).formattedValue.value / .value`
- 腾讯文档专用「表格模型诊断」：比较显示值与底层值，判断打码发生在显示层还是数据/模型层。
- `console-fallback.js` 同步支持腾讯文档（控制台免安装版）。
- 多工作表导出 JSON 时合并为一个文件（修复原有多表 JSON 文件名不匹配的问题）。
- JSON 下载不再添加 UTF-8 BOM，避免部分解析器 `JSON.parse` 失败。
- 工作表名称在弹窗中做 HTML 转义，避免特殊字符破坏 UI。

### 变更

- `manifest.json`：
  - 新增 `*://*.docs.qq.com/*` host permission / content script 匹配；
  - 腾讯文档使用 `"world": "MAIN"` 在页面世界注入 `inject.js`（Chrome / Edge 111+），避开页面 CSP 对脚本标签注入的限制；
  - 版本号升至 `1.3.0`。
- `inject.js`：重构为「平台适配层」，金山/WPS 与腾讯文档共用抽取、序列化和下载协议。
- `popup.js` / `popup.html`：平台识别、腾讯文档模型诊断 UI、错误提示和 URL 校验。

### 兼容性

- 金山文档 / WPS 365：行为保持不变，仍使用原有的 `APP.workbook` 数据模型和 `probe.js` 网络探针。
- 腾讯文档：仅支持「在线表格」`/sheet/`，暂不支持 `/smartsheet/` 智能表格、文字文档和幻灯片。
- 腾讯文档的 `world: "MAIN"` 需要 Chrome / Edge 111+；更老版本会尝试脚本标签兜底，但可能被 CSP 拦截。
