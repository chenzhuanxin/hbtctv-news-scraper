# 云上通城（hbtctctv.cn）新闻全类型抓取工具

针对湖北通城县融媒体中心「云上通城」网站 `https://www.hbtctv.cn/` 的新闻抓取脚本，
收录全站 **全部 6 种内容类型**，输出带格式的 **Excel + HTML** 数据表。

> ⚠️ 仅供学习研究使用，请控制抓取频率，勿对目标站点造成压力。
> 严格遵守站点 robots 的建议参数：`--workers 1 --delay 2`。

## 仓库内容

| 文件 | 说明 | 大小 |
|---|---|---|
| `fetch_all_articles.py` | **全量抓取脚本**：从站点最新 ID 一直扫到最早有效 ID，收录全站历史新闻 | ~44 KB |
| `fetch_articles_by_date.py` | **按截止日期抓取脚本**：输入向前截止日期，抓取「当前日期 → 截止日期」时间段内发布的新闻 | ~48 KB |
| `data/news_since_2026-05-01.xlsx` | 按日期抓取数据示例（4,707 条，2026-05-01 ~ 2026-10-02） | ~3.3 MB |
| [`news_all.xlsx`（Release 附件）](https://github.com/chenzhuanxin/hbtctv-news-scraper/releases/tag/v1.0) | 全量数据示例（87,060 条，2020 ~ 2026-10） | ~98 MB |

> 全量 Excel 约 98 MB，超出 Git API 上传上限，放在 **Release 附件** 中提供：
> [点此下载 v1.0 附件](https://github.com/chenzhuanxin/hbtctv-news-scraper/releases/tag/v1.0)。

## 六种内容类型与内容提取方式

文章页源码 `shareVariable = {sid:'10132', aid:'1', cid:'...', ...}` 中的 `aid`
是内容类型编号（`sid` 是站点编号，全站固定 10132）。经全站普查（92,043 个 ID 抽样验证），
**全站只有以下 6 种取值，无其它类型**：

| 类型ID (aid) | 类型 | 全站条数 | 「内容」列的提取方式 |
|:---:|---|---:|---|
| 1 | 图文 | 63,893 | 正文文字 + `[图片]` 图片链接；正文内嵌视频自动解析为 mp4 直链 |
| 2 | 组图 | 124 | 图集内**全部** `[图片]` 大图地址，按顺序逐张给出，并附每张图的图注 |
| 3 | 转外链 | 11,207 | `[外链]` + 跳转目标地址（页面本身无正文，`window.onload` 直接跳外站） |
| 4 | 视频 | 11,813 | `[视频]` + **mp4 直链**（调站点取流接口 `app.cjyun.org.cn` 获取），附 `[视频封面]` |
| 11 | 直播 | 18 | `[直播页]` 链接 + 封面图 + 摘要 + 经站点接口抓取的直播图文流（含文字/图片/视频帖） |
| 12 | 报名/表单 | 5 | `[报名页]` + 表单页面链接 |

> **关于直播（aid=11）**：直播结束后页面没有主播放器，静态手段无法取得 ts/m3u8 流；
> 脚本给出直播页链接、封面、摘要及直播图文接口返回的全部内容。
> **关于转外链（aid=3）**：页面无标题无发布时间，时间列由相邻文章 ID 线性插值推算，带 `(推)` 标记。

## Excel 输出格式

- **工作表「新闻数据」**，表头：

  | 序号 | 发布时间 | 标题 | 内容 | 类型 | 类型ID | 原文链接 |
  |---|---|---|---|---|---|---|

  冻结首行、开启自动筛选、按发布时间倒序；内容列内图片/视频链接均为可点击的完整 URL。
- **工作表「类型普查」**：各类型 ID 的出现次数与示例链接。
- **工作表「其他类型明细」**：非 1/3/4 的页面逐条列出（含可点击链接）。
- 同时生成同名 `.html` 版表格（内容一致，可直接浏览器打开）。

## 环境要求

- Python 3.10+（开发验证环境：Python 3.14，Windows / Git Bash）
- 依赖：`openpyxl`（仅此一个第三方库）

```bash
pip install openpyxl
```

- 无需浏览器、无需模拟 JS，纯 HTTP 静态抓取（仅视频直链需多调一次站点取流接口）。

## 使用说明

### 1. 按截止日期抓取（fetch_articles_by_date.py）

抓取「当前日期 → 指定截止日期」时间段内发布的全部新闻（默认收录 6 种类型）：

```bash
# 抓取 2026-05-01 至今的全部新闻，输出到 output/
python fetch_articles_by_date.py --since 2026-05-01

# 指定时间段：2026-08-01 ~ 2026-09-01
python fetch_articles_by_date.py --since 2026-08-01 --until 2026-09-01

# 严格遵守 robots 的温和模式
python fetch_articles_by_date.py --since 2026-05-01 --workers 1 --delay 2
```

**工作原理**：文章 ID 与发布时间基本单调递增。脚本从站点最新 ID 倒序按批扫描，
当连续 `--stop-batches` 批的文章发布时间全部早于截止日期时自动停止，
因此只需扫过时间边界以上的 ID，通常 3 分钟内完成，无需扫完全部历史。

**无发布时间的页面**（转外链/报名页）用相邻有日期文章的 ID 做线性插值推算发布时间，
发布时间列以 `(推)` 结尾标注，日期过滤因此精确。

### 2. 全量抓取（fetch_all_articles.py）

从最新 ID 一直扫到站点最早有效文章（当前为 ID 24850），收录全站全部历史：

```bash
# 全量抓取，输出到 output_full/（默认并发 8，约需 40~60 分钟）
python fetch_all_articles.py --out output_full

# 高并发（实测 24 线程约 25 分钟）
python fetch_all_articles.py --workers 24 --out output_full

# 先小范围试跑（只扫 5000 个 ID）
python fetch_all_articles.py --max-scan 5000 --out test_run
```

### 3. 断点续跑

两个脚本都会在输出目录写 `*.records.jsonl`（逐条记录）与 `*.state.json`（断点/失败清单）。
**中途中断后，重跑同一条命令即可续传**，已抓的记录不会重复抓取；
扫描结束后还会自动补抓失败清单中的 ID。

### 4. 只重新生成表格

已有 `*.records.jsonl` 时，可跳过抓取只重新生成 Excel/HTML（改样式或清洗规则后常用）：

```bash
python fetch_articles_by_date.py --since 2026-05-01 --out output --report-only
python fetch_all_articles.py --out output_full --report-only
```

### 5. 忽略断点从头重扫

```bash
python fetch_all_articles.py --rescan
```

## 全部参数一览（两个脚本通用，除非另注）

| 参数 | 默认值 | 说明 |
|---|---|---|
| `--since`（仅日期版） | 2026-05-01 | 向前截止日期（含当天），`YYYY-MM-DD` |
| `--until`（仅日期版） | 今天 | 结束日期，可选 |
| `--out` | `output` | 输出目录 |
| `--start-id` | 自动探测 | 起始文章 ID（一般不用指定） |
| `--min-id`（仅全量版） | 24850 | 最早扫描 ID（站点最早有效文章） |
| `--workers` | 8 | 并发线程数 |
| `--batch` | 300 | 每批处理的 ID 个数 |
| `--stop-batches`（仅日期版） | 2 | 连续 N 批全部早于截止日期即停止 |
| `--max-scan` | 0（不限） | 最多扫描 ID 数，试跑用 |
| `--delay` | 0.02 | 请求间隔秒数（温和抓取建议 `--workers 1 --delay 2`） |
| `--no-video-api` | 关 | 不调取流接口，视频只给播放页地址 |
| `--drop-img-keyword KW` | 空 | 图片 URL 含关键字则丢弃，可重复传 |
| `--sort` | `desc` | 排序，`desc` 最新在前 / `asc` 最旧在前 |
| `--report-only` | 关 | 只用已有 jsonl 重新生成表格 |
| `--rescan` | 关 | 忽略断点状态，从头扫描 |

## 输出文件说明

| 文件 | 说明 |
|---|---|
| `news_all.xlsx` / `news_since_YYYY-MM-DD.xlsx` | Excel 数据表（主交付物） |
| `news_all.html` / 同名 `.html` | 同内容 HTML 表格 |
| `*.records.jsonl` | 逐条 JSON 记录（断点续跑依据，亦便于程序化二次处理） |
| `*.state.json` | 断点位置与失败 ID 清单 |
| `*.aid_census.json` | 抓取过程中实时统计的类型普查结果 |

## 实现要点（逆向结论）

1. **详情页**：`/p/{id}.html`，类型由页内 `shareVariable.aid` 区分。
2. **类型判别**：`1`=图文(`div.article-content`)、`2`=组图(`div.picture-content`,
   大图在 `img[data-img]`、图注在 `data-caption`)、`3`=转外链(极小页+`location.href`跳转)、
   `4`=视频(`div.video-content`)、`11`=直播(`div.body-live`)、`12`=报名(`signup-main`)。
3. **视频直链**：详情页 iframe 指向 `app.cjyun.org.cn/video/player/index?vid=...`，
   调 `https://app.cjyun.org.cn/video/player/video?sid=10132&vid={vid}&type=video`
   得 JSON，`file.sd/hd/ed` 为不同清晰度 mp4 直链。
4. **直播图文**：`https://www.hbtctv.cn/index/live/query?liveid={liveid}&contentid={cid}...`
   返回 JSON 图文流（接口固定返回最新一批，翻页参数无效，按帖 id 去重）。
5. **ID 与时间单调**：ID 递增≈时间递增，是按日期抓取「提前停止」与无时间页
   「线性插值推算」的基础。
6. **站点异常**：约 0.2% 的 ID 返回 HTTP 200 + 空响应体（站点自身问题），记入失败清单；
   个别视频取流接口无文件地址，回退为播放页链接。

## 数据规模参考（2026-10-02 实测）

- 站点有效 ID 范围：24850 ~ 116892，扫描 92,043 个 ID，命中 87,060 条。
- 2026-05-01 ~ 10-02：4,707 条（图文 2,399 / 转外链 1,582 / 视频 724 / 组图 2）。

## License

MIT
