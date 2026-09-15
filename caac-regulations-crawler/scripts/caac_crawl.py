#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
民航局规章 / 技术文档抓取器
============================================================
数据源：中国民用航空局 政府信息公开  https://www.caac.gov.cn/XXGK/XXGK/
四个栏目：法律法规 / 民航规章 / 规范性文件 / 标准规范

核心机制：
  1) 服务端按「有效性=有效」过滤  —— 废止件在源头就不进流水线，比自己猜可靠
  2) 随机延时 + 周期长休 + 遇阻自适应降速       —— 防封
  3) 列表数 / 成功数 / 失败数 三方对账，
     再逐条核验详情页有效性是否与筛选一致      —— 防数据污染
  4) 失败进 failures.jsonl，--retry 只补漏不重扫 —— 提效
  5) 落盘全平铺 + 一份 Excel 索引（含体检表）

依赖：openpyxl（仅用于生成索引）
落盘目录优先级：--out DIR  >  环境变量 CAAC_DIR  >  ./caac_regs_data

用法：
  python caac_crawl.py --out D:/民航法规资料库   # 指定落盘目录
  python caac_crawl.py --probe 10                # 每栏目取前 10 条，验证链路
  python caac_crawl.py                           # 全量（条数由列表页决定，随官网增长）
  python caac_crawl.py --retry                   # 只补失败清单，不扫列表
  python caac_crawl.py --report                  # 只重建索引与体检表，不联网
  python caac_crawl.py --section 民航规章 --fast
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import urllib.robotparser
from datetime import datetime

# ---------------------------------------------------------------- 配置

def default_base() -> str:
    """落盘根目录：环境变量 CAAC_DIR > 当前目录下 caac_regs_data/（--out 优先级最高）。"""
    return os.environ.get('CAAC_DIR') or os.path.join(os.getcwd(), 'caac_regs_data')


BASE = STATE = RECORDS_F = FAILURES_F = LOG_F = INDEX_F = TARGETS_F = ''


def set_base(path: str) -> None:
    """设定落盘根目录并重算全部派生路径。"""
    global BASE, STATE, RECORDS_F, FAILURES_F, LOG_F, INDEX_F, TARGETS_F
    BASE = os.path.abspath(path)
    STATE = os.path.join(BASE, '_state')
    RECORDS_F = os.path.join(STATE, 'records.jsonl')
    FAILURES_F = os.path.join(STATE, 'failures.jsonl')
    LOG_F = os.path.join(STATE, 'crawl.log')
    INDEX_F = os.path.join(BASE, '民航法规索引.xlsx')
    TARGETS_F = os.path.join(STATE, 'targets.json')

# ⚠️ 以下接口参数与选择器均于 2026-09-15 对线上站点实测核实。
#    官网若改版，先回来核对这里（参考 CAAC_Spider 项目：其 2020 年的 #id_tblAppendix
#    选择器在现站已不存在，照抄必挂）。
# (栏目, channelid, fl)
SECTIONS = [
    ('法律法规',   '211383', '12'),
    ('民航规章',   '269689', '13'),
    ('规范性文件', '238066', '14'),
    ('标准规范',   '211383', '15'),
]
PERPAGE = 100          # 每页 100 条，列表请求数从 202 降到 ~28
VALIDITY = '有效'      # 服务端筛选值：有效 / 失效 / 废止 / 历史版本 / All

DELAY_LIST = (0.30, 0.80)
DELAY_DETAIL = (0.40, 1.10)
DELAY_FILE = (0.60, 1.80)
PAUSE_EVERY = 60
DELAY_PAUSE = (4.0, 12.0)
BACKOFF_BASE = 2.0
MAX_RETRY = 3
TIMEOUT = 30

RATE_HEALTHY = 0.05    # 缺失率 <= 5% 视为健康
RATE_UNSTABLE = 0.30   # 缺失率 >= 30% 视为字段整体失效，结果不可信

HOST = 'https://www.caac.gov.cn'
HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                   'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'),
    'Accept-Language': 'zh-CN,zh;q=0.9',
}

ITEM_RE = re.compile(
    r'<a\s+href="(' + re.escape('http://www.caac.gov.cn/XXGK/XXGK/') +
    r'[^"]+?t\d{8}_\d+\.html)"[^>]*?name="([^"]*)"[^>]*>([^<]*)</a>', re.I)
VAR_RE = re.compile(r'var\s+([A-Za-z_$][\w$]*)\s*=\s*"([^"]*)"')
# 必须锚定真实 <a ...>...</a>：详情页里文件名也会出现在内联 JS 字符串中
# （形如 href=\"./P0....pdf\"），不锚定就会把后面的 JS 代码当成锚文本。
ATT_RE = re.compile(
    r'<a\b[^>]*?href\s*=\s*["\'](?:\./)?(P\d+\.(?:pdf|doc|docx|xls|xlsx|zip|wps))["\']'
    r'[^>]*>\s*([^<]{0,220}?)\s*</a>', re.I)
ATT_LOOSE_RE = re.compile(r'(P\d{12,}\.(?:pdf|docx?|xlsx?|zip|wps))', re.I)
JS_NOISE = ('$(', 'function', 'var ', ');', '\\u', '</')
CONTENT_RE = re.compile(r'<div class="content"\s+data-role="n_content"\s*>', re.I)
BAD_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')
SPACES = re.compile(r'\s{2,}')

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')


# ---------------------------------------------------------------- 基础设施

class Pacer:
    """请求节拍器：随机延时 + 周期长休 + 遇阻自适应降速。"""

    def __init__(self, scale: float = 1.0) -> None:
        self.n = 0
        self.factor = scale

    def wait(self, rng) -> None:
        if self.n and self.n % PAUSE_EVERY == 0:
            p = random.uniform(*DELAY_PAUSE) * min(self.factor, 1.0)
            log(f'    · 长休 {p:.1f}s（累计 {self.n} 次请求）')
            time.sleep(p)
        time.sleep(random.uniform(*rng) * self.factor)
        self.n += 1

    def slow_down(self) -> None:
        old = self.factor
        self.factor = min(self.factor * 1.6, 6.0)
        if self.factor != old:
            log(f'    ! 触发降速：延时系数 {old:.2f} → {self.factor:.2f}')


PACER = Pacer()
LOG_FH = None
ROBOTS = None


def log(msg: str = '') -> None:
    print(msg, flush=True)
    if LOG_FH and not LOG_FH.closed:
        LOG_FH.write(msg + '\n')
        LOG_FH.flush()


def fetch(url: str, rng) -> bytes:
    last = None
    for attempt in range(MAX_RETRY + 1):
        try:
            PACER.wait(rng)
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            last = f'HTTP {e.code}'
            if e.code in (403, 429) or 500 <= e.code < 600:
                PACER.slow_down()
        except Exception as e:                      # noqa: BLE001
            last = f'{type(e).__name__}: {e}'
        if attempt < MAX_RETRY:
            time.sleep(BACKOFF_BASE ** attempt * random.uniform(0.8, 1.6))
    raise RuntimeError(last or 'unknown error')


def decode(raw: bytes) -> str:
    for enc in ('utf-8', 'gbk'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def sanitize(s: str, limit: int = 120) -> str:
    s = BAD_CHARS.sub('-', (s or '').strip())
    s = SPACES.sub(' ', s).strip(' .-')
    return s[:limit].strip(' .-')


def load_jsonl(path: str) -> list:
    if not os.path.exists(path):
        return []
    out = []
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    out.append(json.loads(line))
                except json.JSONDecodeError:
                    pass
    return out


def write_jsonl(path: str, rows: list) -> None:
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + '\n')
    os.replace(tmp, path)


def now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def checkpoint(records: dict, failures: list) -> None:
    """增量落盘：长跑中途崩溃也不丢已抓到的元数据（载体文件已在磁盘上）。"""
    write_jsonl(RECORDS_F, list(records.values()))
    write_jsonl(FAILURES_F, failures)


def write_targets(stats: dict) -> None:
    """记录本轮各栏目列表实际条数，供 post_audit.py 对账（仅全量扫描时写）。"""
    merged = {}
    if os.path.exists(TARGETS_F):
        try:
            with open(TARGETS_F, encoding='utf-8') as f:
                merged = json.load(f)
        except Exception:                           # noqa: BLE001
            merged = {}
    merged.update({sec: st['listed'] for sec, st in stats.items()})
    with open(TARGETS_F, 'w', encoding='utf-8') as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    log(f'对账基线已写入：{TARGETS_F}')


def check_robots(url: str) -> bool:
    """用标准库 robotparser 自动遵守 robots.txt；读取失败则保守放行并记录。"""
    global ROBOTS
    if ROBOTS is None:
        ROBOTS = urllib.robotparser.RobotFileParser()
        ROBOTS.set_url(HOST + '/robots.txt')
        try:
            ROBOTS.read()
            log(f'robots.txt 已读取；本站 User-agent: 规则 '
                f'{"允许" if ROBOTS.can_fetch(HEADERS["User-Agent"], HOST + "/") else "禁止"} 抓取')
        except Exception as e:                      # noqa: BLE001
            log(f'! robots.txt 读取失败（{e}），按保守策略继续（低频 + 周期长休）')
            ROBOTS = False
    return True if ROBOTS is False else ROBOTS.can_fetch(HEADERS['User-Agent'], url)


# ---------------------------------------------------------------- 解析

def list_url(cid: str, fl: str, page: int) -> str:
    expr = urllib.parse.quote(f" PARENTID='{fl}' or CLASSINFOID='{fl}' ")
    return (f'{HOST}/was5/web/search?page={page}&channelid={cid}'
            f'&was_custom_expr={expr}&perpage={PERPAGE}&orderby=-fabuDate'
            f'&divShow=wenhao&selST=All'
            f'&selYouxiao={urllib.parse.quote(VALIDITY)}&fl={fl}')


def parse_list(html: str) -> list:
    seen, out = set(), []
    for url, name, text in ITEM_RE.findall(html):
        url = url.replace('http://www.caac.gov.cn', HOST, 1)
        if url in seen:
            continue
        seen.add(url)
        out.append((url, (name or text).strip() or text.strip()))
    return out


def extract_content_div(html: str) -> str:
    m = CONTENT_RE.search(html)
    if not m:
        return ''
    start, end = m.end(), m.end()
    depth = 1
    for tag in re.finditer(r'<(/?)div\b[^>]*>', html[start:], re.I):
        depth += -1 if tag.group(1) else 1
        if depth == 0:
            end = start + tag.start()
            break
    return html[start:end].strip()


def parse_detail(html: str) -> dict:
    v = dict(VAR_RE.findall(html))
    # 命中优先取真实 <a> 锚文本（= 官网原始文件名）；否则退化为只认文件名
    hits = ATT_RE.findall(html) or [(f, '') for f in ATT_LOOSE_RE.findall(html)]
    atts, official = [], ''
    for fn, anchor in hits:
        if fn not in atts:
            atts.append(fn)
        cand = SPACES.sub(' ', anchor.strip())
        if not official and cand and not any(t in cand for t in JS_NOISE):
            official = cand[:200]
    return {
        'buhao': (v.get('buhao') or '').strip(),
        'wenhao': (v.get('wenhao') or '').strip(),
        'bwdw': (v.get('bwdw') or '').strip(),
        'fwrq': (v.get('fwrq') or '').strip()[:10],
        'youxiao': (v.get('youxiao') or '').strip(),
        'weizhi': (v.get('weizhi001') or '').strip(),
        'atts': atts,
        'official': official,
        'body': extract_content_div(html),
    }


def looks_like_code(s: str) -> bool:
    prefixes = ('CCAR', 'CTSO', 'AC', 'AP', 'MD', 'IB', 'WM', 'MH', 'GB', 'GJB',
                'HB', 'ATS', 'SC', 'FS', 'OPS', 'AD')
    if not s:
        return False
    head = s.split(' ')[0]
    return head.split('-')[0].split('/')[0].upper() in prefixes or \
        bool(re.match(r'^[A-Z]{2,6}(?:[/\-][A-Z0-9]+)*[\s\-]?\d', s, re.I))


def build_name(sec: str, title: str, meta: dict, is_pdf: bool) -> str:
    code = meta['buhao'] or (meta['wenhao'] if looks_like_code(meta['wenhao']) else '')
    tail = []
    if meta['wenhao'] and meta['wenhao'] != code:
        tail.append(meta['wenhao'])
    if meta['fwrq']:
        tail.append(meta['fwrq'])
    parts = [f'[{sec}]']
    if code:
        parts.append(sanitize(code, 40))
    parts.append(sanitize(title, 120))
    name = ' '.join(parts)
    if tail:
        name += ' (' + ', '.join(sanitize(t, 60) for t in tail) + ')'
    return name + ('.pdf' if is_pdf else '.html')


def wrap_html(title: str, meta: dict, sec: str, body: str) -> str:
    bits = [b for b in (sec, meta['wenhao'] or meta['buhao'], meta['bwdw'], meta['fwrq']) if b]
    return (
        '<!doctype html>\n<html lang="zh-CN">\n<head>\n<meta charset="utf-8">\n'
        f'<title>{title}</title>\n<style>\n'
        'body{max-width:820px;margin:40px auto;padding:0 20px;'
        'font-family:"Microsoft YaHei","SimSun",serif;font-size:16px;line-height:1.95;color:#1a1a1a}\n'
        'h1{font-size:20px;line-height:1.5;margin:0 0 8px}\n'
        '.meta{color:#6b6b6b;font-size:13px;border-bottom:1px solid #e0e0e0;'
        'padding-bottom:14px;margin-bottom:26px}\n'
        'p{margin:.55em 0}\n'
        '</style>\n</head>\n<body>\n'
        f'<h1>{title}</h1>\n<div class="meta">{" · ".join(bits)}</div>\n'
        f'{body or "<p>（该页无正文内容）</p>"}\n</body>\n</html>\n')


def make_record(sec: str, url: str, title: str, meta: dict, carrier: str,
                name: str, reconcile: str) -> dict:
    return {
        'section': sec, 'title': title, 'code': meta['buhao'],
        'wenhao': meta['wenhao'], 'subcat': meta['weizhi'], 'unit': meta['bwdw'],
        'date': meta['fwrq'], 'validity': meta['youxiao'], 'carrier': carrier,
        'file': name, 'url': url, 'atts': len(meta['atts']),
        'official': meta.get('official', ''),
        'bodylen': len(meta['body']), 'reconcile': reconcile, 'fetched': now(),
    }


# ---------------------------------------------------------------- 抓取

def crawl_section(sec: str, cid: str, fl: str, limit: int | None, records: dict,
                  failures: list, stats: dict) -> None:
    log(f'\n=== {sec} ===')
    if not check_robots(list_url(cid, fl, 1)):
        log(f'  !! robots.txt 不允许抓取该路径，跳过 {sec}')
        return
    items, page = [], 1
    while True:
        url = list_url(cid, fl, page)
        try:
            html = decode(fetch(url, DELAY_LIST))
        except Exception as e:                      # noqa: BLE001
            log(f'  列表第 {page} 页失败：{e}')
            failures.append({'stage': 'list', 'section': sec, 'url': url,
                             'reason': str(e), 'ts': now()})
            break
        got = parse_list(html)
        new = [it for it in got if it[0] not in {u for u, _ in items}]
        if not got:
            if page == 1:
                # 零条目不等于「没有内容」，多半是选择器或接口参数失效 —— 必须炸响
                log('  !! 列表第 1 页解析出 0 条：选择器/接口参数可能已因官网改版失效，'
                    '请核对文件顶部的 SECTIONS 与 ITEM_RE')
                failures.append({'stage': 'selector', 'section': sec, 'url': url,
                                 'reason': '列表页解析为 0 条', 'ts': now()})
            else:
                log(f'  第 {page} 页无条目，列表到底。')
            break
        if not new:
            log(f'  第 {page} 页无新条目，列表到底。')
            break
        items += new
        log(f'  第 {page} 页：{len(got)} 条（累计 {len(items)}）')
        if limit and len(items) >= limit:
            items = items[:limit]
            log(f'  已达测试上限 {limit} 条，停。')
            break
        page += 1
        if page > 60:
            log('  !! 超过 60 页，强制停')
            break

    log(f'  {sec} 列表共 {len(items)} 条，开始逐条处理')
    for i, (detail_url, title) in enumerate(items, 1):
        if detail_url in records:
            continue
        handle_detail(sec, detail_url, title, records, failures)
        if i % 25 == 0:
            log(f'    … 进度 {i}/{len(items)}')
        if i % 50 == 0:
            checkpoint(records, failures)

    # 本轮对账：列表条数 == 落地条数 + 失败条数（checkpointed-scraper 的验收口径）
    urls = [u for u, _ in items]
    ok = sum(1 for u in urls if u in records)
    stats[sec] = {'listed': len(urls), 'ok': ok, 'fail': len(urls) - ok}
    if ok + stats[sec]['fail'] != stats[sec]['listed']:
        log(f'  !! 对账不平：{sec} 列表 {stats[sec]["listed"]} '
            f'≠ 成功 {ok} + 失败 {stats[sec]["fail"]}')


def handle_detail(sec: str, url: str, title: str, records: dict, failures: list) -> None:
    try:
        meta = parse_detail(decode(fetch(url, DELAY_DETAIL)))
    except Exception as e:                          # noqa: BLE001
        log(f'    x 详情失败 {title[:28]}：{e}')
        failures.append({'stage': 'detail', 'section': sec, 'url': url, 'title': title,
                         'reason': str(e), 'ts': now()})
        return

    # —— 对账：服务端筛的是「有效」，详情页应回「有效」 ——
    yx = meta['youxiao']
    if not yx:
        reconcile = '字段缺失'
    elif yx.startswith('有效'):
        reconcile = '一致'
    else:
        reconcile = f'冲突(详情页={yx})'

    if meta['atts']:
        name = build_name(sec, title, meta, True)
        path = os.path.join(BASE, name)
        try:
            if not (os.path.exists(path) and os.path.getsize(path) > 0):
                data = fetch(url.rsplit('/', 1)[0] + '/' + meta['atts'][0], DELAY_FILE)
                if len(data) < 512:
                    raise RuntimeError(f'响应过短 {len(data)}B')
                with open(path + '.part', 'wb') as f:
                    f.write(data)
                os.replace(path + '.part', path)
            records[url] = make_record(sec, url, title, meta, 'PDF', name, reconcile)
            flag = '' if reconcile == '一致' else f'  << {reconcile}'
            log(f'    + {os.path.getsize(path) // 1024:>6}KB  {name[:66]}{flag}')
        except Exception as e:                      # noqa: BLE001
            log(f'    x 下载失败 {title[:28]}：{e}')
            failures.append({'stage': 'download', 'section': sec, 'url': url, 'title': title,
                             'target': name, 'reason': str(e), 'ts': now()})
    else:
        name = build_name(sec, title, meta, False)
        path = os.path.join(BASE, name)
        try:
            if not (os.path.exists(path) and os.path.getsize(path) > 0):
                with open(path + '.part', 'w', encoding='utf-8') as f:
                    f.write(wrap_html(title, meta, sec, meta['body']))
                os.replace(path + '.part', path)
            records[url] = make_record(sec, url, title, meta, '网页', name, reconcile)
            flag = '' if reconcile == '一致' else f'  << {reconcile}'
            log(f'    + 网页正文      {name[:66]}{flag}')
        except Exception as e:                      # noqa: BLE001
            log(f'    x 正文保存失败 {title[:28]}：{e}')
            failures.append({'stage': 'download', 'section': sec, 'url': url, 'title': title,
                             'target': name, 'reason': str(e), 'ts': now()})


# ---------------------------------------------------------------- 对账 + 体检

def health(records: dict, stats: dict) -> list:
    """本轮对账（列表/成功/失败）+ 累计库况（有效性缺失率、冲突率）。"""
    rows = []
    for sec, _, _ in SECTIONS:
        sub = [r for r in records.values() if r['section'] == sec]
        st = stats.get(sec)
        if not sub and not st:
            continue
        n = len(sub)
        missing = sum(1 for r in sub if not r['validity'])
        conflict = sum(1 for r in sub if r.get('reconcile', '').startswith('冲突'))
        rate = missing / n if n else 0.0
        if rate <= RATE_HEALTHY:
            verdict = '字段健康'
        elif rate >= RATE_UNSTABLE:
            verdict = '!! 缺失率过高，疑似官网模板变更，过滤结果不可信'
        else:
            verdict = '个别条目未填；该条保守保留，建议抽查'
        rows.append({
            'section': sec,
            'listed': st['listed'] if st else '—',
            'run_ok': st['ok'] if st else '—',
            'run_fail': st['fail'] if st else '—',
            'ok': n, 'missing': missing, 'missing_rate': rate, 'conflict': conflict,
            'pdf': sum(1 for r in sub if r['carrier'] == 'PDF'),
            'web': sum(1 for r in sub if r['carrier'] == '网页'),
            'conflict_rate': conflict / n if n else 0.0,
            'verdict': verdict,
        })
    return rows


def print_health(rows: list) -> None:
    log('\n' + '=' * 88)
    log('本轮对账 + 累计库况 + 字段体检')
    log('=' * 88)
    log(f'{"栏目":<9}{"列表":>5}{"成功":>5}{"失败":>5} | {"累计":>6}{"缺字段":>7}'
        f'{"缺失率":>8}{"PDF":>6}{"网页":>6}  判定')
    for r in rows:
        log(f'{r["section"]:<9}{str(r["listed"]):>5}{str(r["run_ok"]):>5}'
            f'{str(r["run_fail"]):>5} | {r["ok"]:>6}{r["missing"]:>7}'
            f'{r["missing_rate"]:>7.1%}{r["pdf"]:>6}{r["web"]:>6}  {r["verdict"]}')
    tot = sum(r['ok'] for r in rows)
    miss = sum(r['missing'] for r in rows)
    conf = sum(r['conflict'] for r in rows)
    ran = [r for r in rows if isinstance(r['run_fail'], int)]
    log('-' * 88)
    if ran:
        log(f'本轮：列表 {sum(r["listed"] for r in ran)} 条 = '
            f'成功 {sum(r["run_ok"] for r in ran)} + 失败 {sum(r["run_fail"] for r in ran)}')
    log(f'累计 {tot} 条；字段缺失 {miss} 条（{miss / tot if tot else 0:.1%}）；'
        f'与筛选冲突 {conf} 条（{conf / tot if tot else 0:.1%}）')
    log(f'PDF {sum(r["pdf"] for r in rows)} 份，网页正文 {sum(r["web"] for r in rows)} 份')
    if conf:
        log('!! 存在与「仅有效」筛选冲突的条目，已在索引「对账」列标红，建议人工抽查')
    log('=' * 88)


# ---------------------------------------------------------------- 索引

def write_index(records: dict, rows: list) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        log('!! 缺少 openpyxl，跳过索引生成。安装：pip install openpyxl')
        return

    wb = Workbook()
    ws = wb.active
    ws.title = '索引'
    cols = [('类型', 'section', 12), ('编号', 'code', 16), ('标题', 'title', 54),
            ('文号', 'wenhao', 26), ('子类', 'subcat', 12), ('发布单位', 'unit', 20),
            ('成文日期', 'date', 12), ('有效性', 'validity', 10), ('载体', 'carrier', 8),
            ('对账', 'reconcile', 14), ('官方原名', 'official', 56),
            ('文件名', 'file', 72), ('来源URL', 'url', 60),
            ('抓取时间', 'fetched', 19)]
    head_fill = PatternFill('solid', fgColor='3C3489')
    head_font = Font(color='FFFFFF', bold=True, size=10)
    for i, (label, _, w) in enumerate(cols, 1):
        c = ws.cell(1, i, label)
        c.fill, c.font = head_fill, head_font
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = 'A2'

    order = {s: i for i, (s, _, _) in enumerate(SECTIONS)}
    data = sorted(records.values(), key=lambda r: (order.get(r['section'], 9), r['date'], r['title']))
    warn = Font(color='A32D2D', bold=True, size=10)
    for ri, r in enumerate(data, 2):
        for ci, (_, key, _) in enumerate(cols, 1):
            ws.cell(ri, ci, r.get(key, ''))
        if not r.get('reconcile', '').startswith('一致'):
            for ci in range(1, len(cols) + 1):
                ws.cell(ri, ci).font = warn

    ws2 = wb.create_sheet('对账与体检')
    h2 = ['栏目', '本轮列表', '本轮成功', '本轮失败', '累计', '缺字段', '缺失率',
          '与筛选冲突', '冲突率', 'PDF', '网页正文', '判定']
    keys = ['section', 'listed', 'run_ok', 'run_fail', 'ok', 'missing', 'missing_rate',
            'conflict', 'conflict_rate', 'pdf', 'web', 'verdict']
    for i, label in enumerate(h2, 1):
        c = ws2.cell(1, i, label)
        c.fill, c.font = head_fill, head_font
        ws2.column_dimensions[c.column_letter].width = 14 if i < 12 else 46
    for ri, r in enumerate(rows, 2):
        for ci, k in enumerate(keys, 1):
            ws2.cell(ri, ci, round(r[k], 4) if isinstance(r[k], float) else r[k])
        for ci in (7, 9):
            ws2.cell(ri, ci).number_format = '0.0%'
    wb.save(INDEX_F)
    log(f'\n索引已写入：{INDEX_F}  （{len(data)} 行 + 对账体检表）')


# ---------------------------------------------------------------- 主流程

def main() -> int:
    global LOG_FH, PACER
    ap = argparse.ArgumentParser(description='民航局规章 / 技术文档抓取器')
    ap.add_argument('--probe', type=int, metavar='N', help='每栏目只取前 N 条（验证链路用）')
    ap.add_argument('--retry', action='store_true', help='只重试失败清单，不扫列表')
    ap.add_argument('--report', action='store_true', help='只重建索引与体检表，不联网')
    ap.add_argument('--refresh', action='store_true',
                    help='重取全部详情页刷新元数据（补新增字段），已下载文件不重下')
    ap.add_argument('--section', help='只跑指定栏目，如 --section 民航规章')
    ap.add_argument('--fast', action='store_true', help='延时压缩到 35%%（更快，但更显眼）')
    ap.add_argument('--out', metavar='DIR',
                    help='落盘根目录（默认取环境变量 CAAC_DIR，再默认 ./caac_regs_data）')
    args = ap.parse_args()

    set_base(args.out or default_base())
    if args.fast:
        PACER = Pacer(0.35)

    os.makedirs(BASE, exist_ok=True)
    os.makedirs(STATE, exist_ok=True)
    LOG_FH = open(LOG_F, 'a', encoding='utf-8')
    log(f'\n{"#" * 82}\n# 启动 {now()}  args={vars(args)}')

    records = {r['url']: r for r in load_jsonl(RECORDS_F)}
    for r in records.values():          # 兼容早期版本写入的记录
        if 'reconcile' not in r:
            r['reconcile'] = '一致' if (r.get('validity') or '').startswith('有效') else '字段缺失'
    failures = load_jsonl(FAILURES_F)
    stats = {}
    log(f'已有记录 {len(records)} 条，待补失败 {len(failures)} 条')

    if not args.report:
        if args.retry:
            if not failures:
                log('失败清单为空，无需补漏。')
            else:
                log(f'\n=== 补漏模式：仅处理 {len(failures)} 条失败项，不扫列表 ===')
                pending, failures = failures, []
                for f in pending:
                    if f['stage'] == 'detail':
                        handle_detail(f['section'], f['url'], f.get('title', ''), records, failures)
                    elif f['stage'] == 'download':
                        records.pop(f['url'], None)
                        handle_detail(f['section'], f['url'], f.get('title', ''), records, failures)
                    elif f['stage'] == 'list':
                        try:
                            for du, ti in parse_list(decode(fetch(f['url'], DELAY_LIST))):
                                if du not in records:
                                    handle_detail(f['section'], du, ti, records, failures)
                        except Exception as e:      # noqa: BLE001
                            failures.append({**f, 'reason': str(e), 'ts': now()})
        elif args.refresh:
            log(f'\n=== 刷新元数据：重取 {len(records)} 条详情页；已下载文件不重下 ===')
            for i, (u, r) in enumerate(list(records.items()), 1):
                handle_detail(r['section'], u, r['title'], records, failures)
                if i % 25 == 0:
                    log(f'    … {i}/{len(records)}')
                if i % 50 == 0:
                    checkpoint(records, failures)
        else:
            todo = [(s, c, f) for s, c, f in SECTIONS if not args.section or s == args.section]
            for sec, cid, fl in todo:
                crawl_section(sec, cid, fl, args.probe, records, failures, stats)

        write_jsonl(RECORDS_F, list(records.values()))
        write_jsonl(FAILURES_F, failures)
        if stats and args.probe is None:
            write_targets(stats)
        tail = '（用 --retry 补漏）' if failures else ''
        log(f'\n记录 {len(records)} 条；本轮失败 {len(failures)} 条{tail}')

    rows = health(records, stats)
    print_health(rows)
    if failures:
        log('\n失败明细（前 10 条）：')
        for f in failures[:10]:
            log(f'  · [{f["stage"]}] {f.get("title", f["url"])[:44]}  <- {f["reason"]}')
    write_index(records, rows)

    if LOG_FH:
        LOG_FH.close()
        LOG_FH = None
    log(f'\n完成 {now()}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
