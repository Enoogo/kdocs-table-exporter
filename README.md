# KDocs 表格导出器（KDocs Table Exporter）

Chrome / Edge 扩展（Manifest V3）。一键导出**金山文档 / WPS 365**（`*.kdocs.cn`、`*.wps.cn`）和**腾讯文档在线表格**（`docs.qq.com/sheet/...`）中的任意工作表为 **TSV / CSV / JSON**，或复制到剪贴板。

专治：

- 表格用 **canvas 渲染**，普通"解除复制限制"插件（SuperCopy、Simple Allow Copy 等）无效的情况；
- 文档**锁定 / 禁止复制 / 只读**的情况；
- 想一次性拿走整张表（几百行 × 几十列）而不是手动刷屏截图的情况。

## 工作原理

### 金山文档 / WPS 365

WPS 在线表格虽然用 canvas 画格子，但数据一直保存在页面的 JS 数据模型里：

```
APP.workbook._worksheets._sheets[i]           // 工作表对象
  .sheetData.getName()                        // 表名
  .sheetData.private.gridCells.getLastRow()   // 有效行数
  .getCell(row, col).getValue()               // 单元格取值（0 基下标）
    └─ .private.str / .private.num            // 文本 / 数值
```

### 腾讯文档在线表格

腾讯文档同样用 canvas 画格子，数据在 `SpreadsheetApp` 的 workbook 模型里：

```
SpreadsheetApp.workbook.worksheetManager.getSheetList()
  -> sheet.getSheetName()                     // 表名
  -> sheet.cellDataGrid.usedRange             // 有效范围（0 基，含首尾）
  -> sheet.getCellDataAtPosition(row, col)    // 单元格对象
       .formattedValue.value                  // 显示值（日期/百分比等按表格格式）
       .value                                 // 底层值
```

扩展的 `inject.js` 注入到**页面世界**直接读取该模型（内容脚本的隔离世界看不到 `APP` / `SpreadsheetApp`），再通过 `postMessage` 桥把数据交给 `content.js` 落盘下载。全程只读，不修改文档内容。

> 腾讯文档的页面 CSP 比较严格，`inject.js` 在 `manifest.json` 里用 `"world": "MAIN"` 声明为页面世界的内容脚本（Chrome / Edge 111+）。金山/WPS 仍沿用脚本标签注入。

## 安装

1. 在本地新建一个文件夹，例如 `D:\chrome-ext\kdocs-table-exporter`。
2. 把本目录下的文件保存进去（在 VFS 页面里打开每个文件 → 复制内容 → 本地新建同名文件粘贴保存）：
   - `manifest.json`
   - `content.js`
   - `inject.js`
   - `popup.html` / `popup.js`
   - `probe.js`（可选，金山/WPS 的网络探针）
3. 打开 `chrome://extensions`（Edge 是 `edge://extensions`），右上角打开 **开发者模式**。
4. 点 **加载已解压的扩展程序**，选择刚才的文件夹。
5. 打开任意在线表格页面并**刷新**（首次安装/更新后需刷新已打开的页面），点工具栏上的扩展图标即可使用。

> 若想固定在工具栏：点浏览器右上角拼图图标 → 找到「KDocs 表格导出器」→ 点图钉。

## 使用

1. 打开一个在线表格文档：
   - 金山文档 / WPS 365：URL 形如 `https://365.kdocs.cn/l/xxxxx`；
   - 腾讯文档：URL 形如 `https://docs.qq.com/sheet/DVUpkdUl4S0FKVXVp`。
2. 点扩展图标，弹窗会自动扫描出所有工作表（金山/WPS 会自动跳过引擎保留表 `WpsReserved_*`）。
3. 勾选要导出的工作表，选格式：
   - **TSV**：Excel 可直接「粘贴」，或用「数据 → 从文本/CSV」导入；
   - **CSV**：通用交换格式，带 UTF-8 BOM，Excel 打开不乱码；
   - **JSON**：`{ 表名: [[行], [行], ...] }`，方便程序处理；多选工作表时会合并成一个 JSON 文件。
4. 点 **导出并下载**（多个工作表会分别下载多个文件，浏览器首次会询问是否允许多文件下载），或点 **复制到剪贴板** 直接粘进 Excel。

参数说明：

- **最大行 / 最大列**：扫描上限，默认 20000 × 256。表特别大时可调大（上限 200000 × 1024），扫描耗时会增加。
- 弹窗里显示的 `≈行×列` 是有效范围的估算值（WPS 按 64 行 × 32 列的块存储；腾讯文档用 `usedRange`），实际导出会**自动裁掉尾部空行空列**。
- 腾讯文档导出优先使用 `formattedValue.value`，所以日期、百分比、货币等会按表格里看到的样子导出；表格里显示为 `****` / `●●●` 的字段，导出后仍是打码值，扩展不会尝试还原被打码的原值。

## 诊断工具：脱敏层级探针

### 金山文档 / WPS 365：网络探针（probe.js）

用于判断敏感字段的 `●●●●●●●` 打码发生在哪一层（严格只统计次数，不保存任何号码原文）：

- 在 `document_start` 注入，先于页面脚本挂钩 `WebSocket` / `fetch` / `XMLHttpRequest`；
- 统计下发数据里 `●●●●●●●` 出现次数、身份证/手机号形态字符串出现次数（仅存打码预览 `前3位****后2位`）；
- 结果挂在 `window.__KDTE_PROBE__`，读取它即可判断：
  - 下发数据含圆点、无号码形态 → **服务端脱敏**（真值从未下发到浏览器）；
  - 下发数据含号码形态字符串 → **客户端安装时打码**（真值到过浏览器）。

### 腾讯文档：表格模型诊断

腾讯文档的表格数据以 **zlib + protobuf** 压缩后下发，网络探针看不到明文；弹窗改为直接检查表格模型：

- 统计显示值里含 `****` / `●●●` 的单元格数量；
- 比较 `formattedValue.value`（显示值）和 `value`（底层值）：
  - 显示值打码、底层值是非打码原值 → **显示层/客户端打码**（真值到过浏览器）；
  - 显示值和底层值都打码 → **数据下发/建模阶段已打码**（客户端模型里没有真值）；
  - 两者一致且无打码 → 未发现打码单元格。
- 最多展示 5 条 `前3位****后2位` 的打码预览，不保存完整原文。

启用：把更新后的 `manifest.json`、`content.js`、`inject.js`、`popup.js`、`popup.html` 放进插件目录 → 在 `chrome://extensions` 点「重新加载」→ 硬刷新页面（Ctrl+Shift+R）。

## 已知限制

| 情况 | 说明 |
|---|---|
| 日期单元格 | 金山/WPS 导出底层值（可能需要 Excel 再格式化）；腾讯文档导出 `formattedValue` 显示值 |
| 合并单元格 | 值只出现在左上角单元格，其余为空 |
| 公式 | 导出的是**计算结果**（显示值），不是公式本身 |
| 文字/演示/智能表格 | 只支持「表格」文档；腾讯文档的 `/smartsheet/` 智能表格暂不支持 |
| 隐藏工作表 | 会正常导出（数据仍在模型里）；`WpsReserved_*` 引擎保留表自动跳过 |
| 表格未加载完 | 若提示"未找到表格数据模型"，等几秒或刷新页面重试 |
| 腾讯文档浏览器版本 | `world: "MAIN"` 需要 Chrome / Edge 111+；更老版本会尝试脚本标签兜底，但可能被 CSP 拦截 |

## 没法装插件？用控制台脚本（备用方案）

见同目录 `console-fallback.js`：在金山/WPS 或腾讯文档的表格页按 `F12` → Console → 粘贴回车，即可把**当前第一个工作表**导出为 TSV 文件。功能不如扩展全，但零安装。

## 隐私

- 完全本地运行，不收集、不上传任何数据；
- 权限只申请了 `clipboardWrite`（复制到剪贴板）和对 `*.kdocs.cn` / `*.wps.cn` / `*.docs.qq.com` 的访问；
- 只读取数据，不写入、不修改文档。
