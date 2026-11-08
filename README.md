# Diffchecker

零依赖、纯前端的离线对比工具集。所有处理都在浏览器本地完成，不会上传任何数据。

| 页面 | 功能 |
| --- | --- |
| `text-compare.html` | 文本对比：并排 / 统一视图，智能 / 按词 / 按字符的行内差异，语法高亮，忽略规则，导出 PDF / Excel / PNG / `.diffchecker` |
| `image-compare.html` | 图片对比：并排、淡入淡出、闪烁、滑块、高亮、差值，自动 / 手动对齐，差异区域检测，导出 PNG / PDF |
| `excel-compare.html` | 表格对比：`.xlsx` / `.csv` / `.tsv`，按表头或按行匹配，导出 Excel / CSV / PDF / `.diffchecker` |

## 使用

直接用浏览器打开 `index.html`（会跳转到文本对比），无需构建，也无需联网。
界面语言默认跟随浏览器语言，可在右上角切换中文 / English。

## 说明

- 文本与 CSV 文件会优先按 UTF-8 解码，失败时回退到 GB18030，因此常见的 GBK 中文文件不会乱码。
- `.diffchecker` 文件保存两侧内容与选项，可在对应的页面重新打开；文本页与表格页的文件互不通用。
- 页面带有严格的 CSP（`default-src 'none'`），不会加载任何外部脚本。因此图片页中可选的 OpenCV 对齐只有在你自行把 `opencv.js` 内联到页面时才会启用，默认使用内置的 FFT 相位相关对齐。
- 不支持旧版二进制 `.xls`，请先另存为 `.xlsx` 或 `.csv`。
