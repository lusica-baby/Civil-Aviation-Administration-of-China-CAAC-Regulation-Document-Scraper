---
name: caac-regulations-crawler
description: 抓取中国民用航空局（CAAC）官网政府信息公开栏目下的规章与技术文档，覆盖「法律法规 / 民航规章 / 规范性文件 / 标准规范」四类约 2600 余条，按服务端「仅现行有效」筛选，落盘为扁平目录 + Excel 索引。当需要获取 CCAR 部号规章、AC/AP/MD 咨询通告、MH 行业标准、CTSO 技术标准订单等民航局公开文件原文，或需要批量下载、更新民航法规资料库时使用。
agent_created: true
---

# 民航局规章文档抓取器

## 用途

抓取 `www.caac.gov.cn` 政府信息公开栏目下的公开规章与技术文档。四类目标：

| 栏目 | channelid | fl | 现行有效条数 |
|---|---|---|---|
| 法律法规 | 211383 | 12 | 47 |
| 民航规章 | 269689 | 13 | 130 |
| 规范性文件 | 238066 | 14 | 1357 |
| 标准规范 | 211383 | 15 | 1152 |

**不使用浏览器自动化（RPA）。** 站点无登录、无验证码、无 JS 渲染，列表走站群检索接口 `was5/web/search`，纯 HTTP 即可。脚本直连约 40 分钟跑完；RPA 需开着窗口点两千余次详情页，慢且脆弱。

## 前置条件

- Python 3.8+（仅用标准库；`openpyxl` 只在生成 Excel 索引时需要）
- 落盘目录预留约 5 GB（实测 2686 条 / 4.28 GB）
- 网络可达 `www.caac.gov.cn`

安装可选依赖：

```bash
pip install openpyxl
```

## 使用流程

落盘目录优先级：`--out DIR` > 环境变量 `CAAC_DIR` > `./caac_regs_data`。

**第一步——小样试跑，验证链路。** 每栏目取 1 条，确认命名、附件、正文三条路径都通：

```bash
python scripts/caac_crawl.py --out /path/to/data --probe 1
```

**第二步——全量抓取。** 已有条目自动跳过，中断后重跑即续传：

```bash
python scripts/caac_crawl.py --out /path/to/data            # 约 1.5-2 小时
python scripts/caac_crawl.py --out /path/to/data --fast     # 延时压到 35%，更快但更显眼
```

**第三步——验收。** 六项硬检查：对账闭合、记录与落盘一致、命名合规、路径长度、载体完整性（PDF 魔数）、数据纯度：

```bash
python scripts/post_audit.py --out /path/to/data
```

**第四步——整理冲突件。** 默认干跑，加 `--apply` 才动手；把官网索引滞后导致的失效/废止件移入存档子目录，并给索引追加「现行有效」sheet：

```bash
python scripts/organize.py --out /path/to/data            # 预览
python scripts/organize.py --out /path/to/data --apply    # 执行
```

## 命令速查

| 命令 | 作用 |
|---|---|
| `--out DIR` | 指定落盘目录 |
| `--probe N` | 每栏目只取前 N 条（验证链路） |
| `--section 名称` | 只跑指定栏目 |
| `--fast` | 延时压缩到 35% |
| `--retry` | 只补失败清单，不重扫列表（增量补漏） |
| `--refresh` | 重取全部详情页刷新元数据，已下载文件不重下 |
| `--report` | 只重建索引与体检表，不联网 |
| `organize.py --apply` | 实际执行归档（不加则干跑） |

## 关键设计

**服务端筛选，而非事后过滤。** 列表接口支持 `selYouxiao=有效`，废止件在源头就不进流水线。`perpage=100` 把列表请求数从 200 余次降到约 28 次。

**三方对账。** 每栏目断言 `列表条数 == 成功 + 失败`，不平即报错。抓取器把各栏目列表实际条数写入 `_state/targets.json`，`post_audit.py` 读它做目标比对。

**平坦命名。** 附件在官网是无意义流水号（`P020260717459628983882.pdf`），必须按元数据重命名：

```
[民航规章] CCAR-290-R4 通用航空经营管理规定 (交通运输部令2026年第14号, 2026-07-01).pdf
[标准规范] 空中交通无线电通话用语 (MH-T 4016-2023, 2026-08-21).pdf
```

**增量落盘。** 每 50 条 checkpoint 一次，长跑中途崩溃不丢已抓元数据。已存在的 URL 直接跳过，因此重跑天然等价于增量更新。

**防封。** 随机延时 + 每 60 次请求长休 + 遇 403/429/5xx 自适应降速（系数 ×1.6，上限 6.0）。标准库 `urllib.robotparser` 自动遵守 robots.txt，读取失败则保守放行并留痕。

## 已知陷阱

以下均经线上实测，改脚本前先读 `references/site-interface-notes.md`：

- **`selYouxiao` 服务端筛选不可信。** 全量实测 781 条（29.1%）官网列表标"有效"、详情页却标"失效/废止"。分栏目差异悬殊：民航规章 0%、法律法规 4.3%、标准规范 25%、规范性文件 **36%**。它只能当粗筛，**最终归档依据必须用详情页字段**——这正是 `organize.py` 存在的理由。
- **附件锚文本即官方真实文件名**，但同一文件名也在内联 JS 字符串里出现一次（`href="./P0....pdf"`）。正则必须锚定完整的 `<a ...>锚文本</a>`，否则会把 JS 代码当成文件名（曾静默污染 103 行索引）。
- **附件扩展名白名单要放宽**：官方实际给出 `pdf / doc / zip / rar / jpg`，且存在扩展名写错的情况（实测 17 个"PDF"实为 OLE2 文档或 ZIP 图片包），需按文件魔数纠正。
- **部分条目官网自身无内容**：正文区为空、附件链接数为 0。这是官网数据缺失，不是抓取失败。
- **列表第 1 页解析出 0 条 = 选择器失效**，脚本会显式报警并记入失败清单（参考 `sanshungit/CAAC_Spider`：其 2020 年的 `#id_tblAppendix` 选择器在现站已不存在）。

## 输出结构

```
<落盘目录>/
├── 民航法规索引.xlsx              # 索引 / 对账与体检 / 现行有效（organize 后）三个 sheet
├── [民航规章] CCAR-145R4 ….pdf    # 全部平铺，靠文件名与索引定位
├── 民航法规_已失效废止存档/        # organize.py 归档的冲突件
└── _state/                        # records.jsonl / failures.jsonl / targets.json / crawl.log
```

`records.jsonl` 是唯一真相源，每条含类型、编号、文号、发布单位、成文日期、有效性、载体、官方原名、来源 URL、抓取时间。重跑 `--report` 可随时从它重建索引。

## 相关文件

- `scripts/caac_crawl.py` — 抓取器（主程序）
- `scripts/post_audit.py` — 六项验收器
- `scripts/organize.py` — 冲突件归档 + 索引整理
- `tests/test_parse.py` — 解析逻辑回归测试（30 项断言，零依赖）
- `references/site-interface-notes.md` — 站点接口逆向笔记，官网改版时先读这份

## 边界

只抓取官网公开信息，不做任何绕过访问控制的操作。遵守 robots.txt，请求频率受节拍器约束。仅供法规查阅与研究用途；文档版权归中国民用航空局所有，转载引用请注明官方来源。
