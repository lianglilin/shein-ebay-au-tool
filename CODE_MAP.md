# 代码结构

## 主要文件

| 文件 | 职责 |
|---|---|
| `ebay_tool.py` | 桌面界面、选品流程、图片下载、XLSX 汇总表读写、编号与去重 |
| `core_rules.py` | 标题规则、风险词、价格规则、XLSX 读取 |
| `shein_fetch.py` | Chrome/Edge DevTools 抓取、希音列表页和商品页数据解析 |

## ebay_tool.py

| 模块 | 主要内容 |
|---|---|
| 浏览器与档案 | `browser_settings`、`prepare_edge_profile`、浏览器进程管理 |
| 图片 | `download_image`、`download_images`、`image_for_embed` |
| XLSX 输出 | `write_workbook`、图片锚点、超链接、多工作表 |
| 文件与编号 | 输出锁、临时文件、编号接续、历史工作簿定位 |
| 去重 | 标题指纹、商品 ID/SKU、历史记录 |
| 变体与标题 | 变体挑选、多选标题处理、澳洲拼写、多颜色商品 |
| 商品文档 | 商品文件夹中的 `商品信息.txt` |
| 业务流程 | `run` 按表格链接抓取；`run_category` 从分类/搜索页抓取 |
| 桌面界面 | Tkinter 主窗口、后台线程、停止/重试/日志 |

## core_rules.py

- 价格区间、加价预设和默认路径。
- 禁品、大牌、风险词和澳洲拼写规则。
- eBay 标题清理、去重和 80 字符截断。
- 标准库 XLSX 读取与价格计算工具。

## shein_fetch.py

- Chrome、Edge 可执行文件和 DevTools 连接管理。
- 分类页、商品页 DOM 和 `window.gbRawData` 解析。
- 验证/登录墙、语言切换、断线重连和人工处理提示。

## tools

- `merge_results.py`：把一份结果并入另一个输出目录并接续编号。
- `repair_images.py`：从商品文件夹补回汇总表中的缺失图片。
- `repair_links.py`：把链接列恢复为真实 URL，并保留超链接目标。
