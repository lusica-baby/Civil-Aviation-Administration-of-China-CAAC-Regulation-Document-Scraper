[README.md](https://github.com/user-attachments/files/32243633/README.md)
# 民航局规章文档抓取器

批量抓取 **中国民用航空局（CAAC）** 官网政府信息公开栏目下的公开规章与技术文档，落盘为一套可直接检索的本地资料库。

覆盖四个栏目共约 **2600 余条现行有效文档**：

| 栏目 | 现行有效 | 内容 |
|---|---|---|
| 法律法规 | 47 | 民航法、行政法规 |
| 民航规章 | 130 | CCAR-XX 部号规章 |
| 规范性文件 | 1357 | AC / AP / MD 咨询通告、管理程序 |
| 标准规范 | 1152 | MH / MH-T 行业标准、CTSO 技术标准订单 |

实测全量抓取产出 **2686 条 / 4.28 GB**，耗时约 1 小时 54 分（`--fast` 模式），**0 失败**。

---

**English** — A dependency-light Python crawler that mirrors the publicly available regulations and technical documents of the Civil Aviation Administration of China (CAAC): laws, CCAR regulations, normative documents (AC / AP / MD), and industry standards (MH / CTSO). 2686 currently-effective documents, ~4.3 GB, zero failures.

This is deliberately **not** an RPA project. The site has no login, no captcha and no JS rendering, so the lists are fetched straight from the site-group search endpoint (`was5/web/search`) over plain HTTP and then reconciled against each detail page — see [为什么不用 RPA](#为什么不用-rpa) below.

```bash
pip install openpyxl
python scripts/caac_crawl.py --out D:/caac_regs --probe 1   # smoke test
python scripts/caac_crawl.py --out D:/caac_regs --fast      # full crawl
```

## 为什么不用 RPA

这件事看起来像 RPA 的活，其实不是。民航局站点：

- 无登录、无验证码、无 JS 渲染 —— 列表走站群检索接口 `was5/web/search`，纯 HTTP 返回 HTML
- robots.txt 对 `User-agent: *` 放行（仅挡 `/CAAC/local/` 与 `/image/`）

真 RPA（录制点击那类）存在的理由是**被访问控制挡住**。这三条都不存在时，硬上 RPA 的结果是慢十倍、脆十倍——跑两千次详情页，中途任何一个弹窗都能让它卡死。同样一批数据，脚本四十分钟跑完。

所以本项目走**直连接口 + 对账式抓取**。

## 安装

```bash
git clone <this-repo>
cd caac-regulations-crawler
pip install openpyxl      # 仅生成 Excel 索引时需要，其余全部标准库
```

Python 3.8+ 即可，无其他依赖。

## 快速开始

```bash
# 1. 小样试跑：每栏目取 1 条，验证链路（推荐先做）
python scripts/caac_crawl.py --out D:/民航法规资料库 --probe 1

# 2. 全量抓取（已有条目自动跳过，中断后重跑即续传）
python scripts/caac_crawl.py --out D:/民航法规资料库 --fast

# 3. 六项验收
python scripts/post_audit.py --out D:/民航法规资料库

# 4. 归档冲突件 + 生成「现行有效」索引（先干跑看范围）
python scripts/organize.py --out D:/民航法规资料库
python scripts/organize.py --out D:/民航法规资料库 --apply
```

落盘目录优先级：`--out DIR` > 环境变量 `CAAC_DIR` > `./caac_regs_data`。

### 命令行参数

| 参数 | 说明 |
|---|---|
| `--out DIR` | 落盘根目录 |
| `--probe N` | 每栏目只取前 N 条，验证链路用 |
| `--section 名称` | 只跑指定栏目，如 `--section 民航规章` |
| `--fast` | 延时压缩到 35%（更快，但请求更显眼） |
| `--retry` | 只补失败清单，不重扫列表 |
| `--refresh` | 重取详情页刷新元数据，已下文件不重下 |
| `--report` | 只重建索引与体检表，不联网 |

## 输出结构

```
D:/民航法规资料库/
├── 民航法规索引.xlsx               # 3 个 sheet：索引 / 对账与体检 / 现行有效
├── [民航规章] CCAR-145R4 民用航空器维修单位合格审定规则 (交通运输部令2022年第8号, 2022-02-11).pdf
├── [标准规范] MH-T 4014-2026 空中交通无线电通话用语 (2026-08-21).pdf
├── [法律法规] 中华人民共和国民用航空法（2026年7月1日起施行） (中华人民共和国主席令第六十五号, 2025-12-27).html
├── 民航法规_已失效废止存档/         # organize.py 归档
└── _state/
    ├── records.jsonl               # 元数据真相源
    ├── failures.jsonl              # 失败清单，供 --retry 补漏
    ├── targets.json                # 各栏目列表实际条数（对账基线）
    └── crawl.log
```

全部文件**平铺**在一个目录，靠文件名与 Excel 索引定位，不做目录分层。

### 命名规则

```
[类型] 编号 标题 (文号, 日期).扩展名
```

官方附件名是无意义流水号（`P020260717459628983882.pdf`），直接下载会得到一堆乱码文件，因此按详情页元数据重命名。官网的原始文件名保留在索引的「官方原名」列，便于追溯。

## 设计要点

**服务端筛选优先。** 列表接口支持 `selYouxiao=有效`，废止件在源头就被排除。配合 `perpage=100`，列表请求数从 200 余次降到约 28 次。

**三方对账。** 每栏目断言 `列表条数 == 成功 + 失败`，不平即报警。抓取器把实际条数写入 `_state/targets.json`，验收器读它比对。

**增量语义。** 已存在的 URL 直接跳过，因此**重跑一次就等价于增量更新**，不需要额外记录"上次抓到哪"。失败条目单独进 `failures.jsonl`，`--retry` 只补这些。

**增量落盘。** 每 50 条 checkpoint 一次，长跑中途崩溃不会丢掉已抓到的元数据。

**防封。** 随机延时 + 每 60 次请求长休 + 遇 403/429/5xx 自适应降速。标准库 `urllib.robotparser` 自动遵守 robots.txt。

**六项验收。** 对账闭合、记录与落盘一致、命名合规、路径长度（Windows 260 字符）、载体完整性（PDF 校验 `%PDF-` 魔数）、数据纯度（冲突率 / 字段缺失 / 重名）。其中 PDF 魔数校验能揪出"文件存在但存的是错误页"这种情况。

## 已知陷阱

> 完整的接口逆向笔记见 [`references/site-interface-notes.md`](references/site-interface-notes.md)。官网改版时先读那份。

- **`selYouxiao` 服务端筛选不可信。** 全量实测 **781 条（29.1%）** 官网列表标"有效"、详情页却标"失效/废止"。分栏目：民航规章 **0%**、法律法规 4.3%、标准规范 25%、规范性文件 **36%**。根因是官网只更新详情页状态、检索索引不同步。它只能当粗筛，**最终归档依据必须用详情页字段**。

- **附件锚文本 = 官方真实文件名，但同一文件名也出现在内联 JS 字符串里。** 正则必须锚定完整的 `<a ...>锚文本</a>`。早期版本没锚定，把 JS 代码当成了文件名，**静默污染 103 行索引**。`tests/test_parse.py` 用合成陷阱把这个 bug 钉住了。

- **附件类型比想象中杂。** 官方实际给出 `pdf / doc / zip / rar / jpg`，且存在扩展名标错的情况——实测 17 个"PDF"实为 OLE2 文档（真身 `.doc`）或 ZIP 图片包（扫描件，单份最大 43 MB）。脚本按文件魔数自动纠正扩展名。

- **部分条目官网自身无内容。** 正文区为空、附件链接数为 0，页面只有一句标准适用范围。这是官网数据缺失，不是抓取失败，脚本会在索引备注中标注。

- **列表第 1 页解析出 0 条 = 选择器失效。** 脚本会显式报警而非静默返回空结果。参考前车之鉴 `sanshungit/CAAC_Spider`（2020 年）：它依赖的 `#id_tblAppendix` 选择器在现站已不存在，照抄必挂。

## 回归测试

```bash
python tests/test_parse.py
```

30 项断言，零依赖。包含合成陷阱用例（同一页同时放 JS 字符串 href 和真实 `<a>`）与 4 个真实页面快照 fixture。

## 相关项目调研

动手前检索过 GitHub / Gitee / HuggingFace 上的同类项目。结论：**没有可直接复用的成品。**

- GitHub 仓库搜索只索引仓库名 / 描述 / README，中文政务网站采集器类项目在索引里基本不存在；英文关键词命中的同类项全是 FAA、沙特 GACA、14-CFR，无一中国的。
- 唯一相关的 `sanshungit/CAAC_Spider` 是 2020 年的单文件脚本，官网改版后已失效，且无延时、无重试、页数写死。
- 通用爬虫框架（Scrapy / Crawlee / Scrapling）解决的是调度与自适应选择器，而本项目的重量在**服务端筛选参数逆向 + 对账 + 扁平命名**，框架帮不上忙还要付适配成本。

真正吸收进来的三样：

1. **`urllib.robotparser`**（标准库，零成本）—— 自动遵守 robots.txt
2. **对账式验收口径**（来自 `checkpointed-scraper`）—— 列表数 == 成功 + 失败，不平即报错
3. **接口参数集中在文件顶部并标注核实日期**（来自 `safe_crawler`）—— 官网改版时能一秒定位

替代数据源也验证过，均不可用：`aviationsafety.caac.gov.cn` 的直下接口整站 502；`flk.npc.gov.cn` 的 `/api/` 返回 SPA 空壳；`hbba.sacinfo.org.cn` 查询参数不生效。

## 免责声明

本项目只抓取官网公开信息，不绕过任何访问控制，遵守 robots.txt 并限制请求频率。文档版权归中国民用航空局所有，本工具仅供法规查阅与研究用途；转载引用请注明官方来源。网站结构变更可能导致脚本失效，使用前请先用 `--probe` 验证。

## License

[MIT](LICENSE)
