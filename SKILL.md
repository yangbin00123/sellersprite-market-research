---
name: sellersprite-market-research
description: 给一个关键词，用 Playwright 自动在卖家精灵采集亚马逊市场大盘调研所需的 15 张表（近30天竞品表、近12个月月表、头部ASIN关键词反查表、核心关键词转化率表），并按关键词建文件夹整理。用户说"采集/下载 XXX 的大盘数据""调研 XXX 市场"时使用。
---

# 卖家精灵市场大盘数据采集

## 输入
一个关键词（例如 `nose rings`），默认美国站。

## 运行（在用户自己的电脑上）
```bash
pip install -r scripts/requirements.txt && playwright install chromium   # 仅首次
python scripts/collect.py "<关键词>"            # 输出到 ~/SellerSprite调研/<关键词>/
```
可选：`--cdp http://localhost:9222`（接管你已登录的 Chrome，无需重新登录；Chrome 需用 `--remote-debugging-port=9222` 启动）、`--recon`（侦察模式，保存页面截图和控件清单，用于适配改版）、`--out 目录`、`--site CA|US`（默认 CA）、`--step products|months|keywords|conv`（只跑某一步）。
竞品页/拓词页/转化率页均用网址参数直达（`/competitor-lookup`、`/traffic/extend`、`/keyword-conversion-rate`），不再点下拉框。导出每天有次数上限（约 48 次，一个关键词用 15 次），脚本导出前会读取剩余次数。
首次会弹出浏览器，用户手动登录卖家精灵，登录态保存在 `~/.sellersprite-profile`。

## 流程（与《亚马逊市场大盘调研》文档一致）
1. 竞品查询：只填关键词（不选类目），取消"展示所有变体"，按月销量降序。
   - 近30天 → 保持卖家精灵原文件名（如 `Competitor-US-Last-30-days-xxxx.xlsx`）
   - 近12个月逐月 → `Competitor-US-YYYY.MM-xxxx.xlsx` ×12
2. 取 01 表月销量前 20 个 ASIN → 拓展流量词/反查，按流量占比降序导出 → `ExpandKeywords-US-…xlsx`
3. 累计流量占比取前 95% 的关键词（最多1000）→ 关键词转化率（近90天）→ `KeywordConversionRate…xlsx`
4. 校验应共 15 张。

## 特性与排错
- 下载的表格不改名；关键词文件夹里只有这 15 张表。截图、状态、关键词清单放在 `<out>/_logs/<关键词>/`。
- 幂等：已下载的文件自动跳过，中断后直接重跑。
- 失败会在输出目录在 `_logs` 生成 `debug_*.png`；页面地址和常量在 `collect.py` 顶部，改版后优先改那里。
- 导出有当日次数限制，次数用完会报"导出超时"。
- 范围仅限文档第一部分（卖家精灵下载），不含后续 AI 分析、细分市场机会分析。
