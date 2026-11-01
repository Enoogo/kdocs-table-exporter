# 升级到 v1.3.0（支持腾讯文档）

## 1. 备份

先把原来的插件目录复制一份，例如：

```
D:\chrome-ext\kdocs-table-exporter.bak
```

## 2. 替换文件

把下面这些文件覆盖到原插件目录（`D:\chrome-ext\kdocs-table-exporter`）：

必需：

- `manifest.json`
- `content.js`
- `inject.js`
- `popup.html`
- `popup.js`

建议一并更新：

- `probe.js`（金山/WPS 网络探针，未改动也可继续用）
- `console-fallback.js`（控制台免安装版，已支持腾讯文档）
- `README.md`
- `CHANGELOG.md`

## 3. 重新加载扩展

1. 打开 `chrome://extensions`（Edge 是 `edge://extensions`）。
2. 找到「KDocs 表格导出器」。
3. 点卡片上的 **重新加载** 按钮。
4. 回到腾讯文档页面，按 `Ctrl+Shift+R` 硬刷新（首次更新后必须刷新）。

## 4. 验证

1. 打开一个腾讯文档在线表格，URL 形如：
   `https://docs.qq.com/sheet/DVUpkdUl4S0FKVXVp`
2. 点扩展图标。
3. 弹窗应显示「腾讯文档 · 共 N 个工作表」。
4. 勾选工作表 → 选择 TSV / CSV / JSON → **导出并下载**。
5. 也可以点 **复制第一个所选表（TSV）到剪贴板**。

## 5. 腾讯文档诊断

弹窗里的「🕵️ 脱敏探针诊断」在腾讯文档下会使用**表格模型诊断**：

- 对比 `formattedValue.value`（显示值）和 `value`（底层值）；
- 只统计打码单元格数量，最多显示 5 条 `前3位****后2位` 预览；
- 不会还原、也不会导出被打码的原值。

## 6. 常见问题

| 现象 | 处理 |
|---|---|
| 弹窗提示「无法连接页面」 | 确认已在 `chrome://extensions` 重新加载，并硬刷新腾讯文档页面 |
| 提示「未找到腾讯文档表格数据模型」 | 确认打开的是 `/sheet/` 在线表格（不是智能表格/文字/幻灯片），等表格加载完再点「重新扫描工作表」 |
| 提示「表格响应超时」 | 表格较大或尚未加载完，等几秒后重试；必要时刷新页面 |
| 老版本 Chrome | 腾讯文档需要 Chrome / Edge 111+（`world: "MAIN"`）；更老版本会尝试兜底，但可能被页面 CSP 拦截 |
