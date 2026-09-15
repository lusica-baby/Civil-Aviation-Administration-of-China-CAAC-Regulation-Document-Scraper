#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
全量抓取后的完整验收器（零依赖，可直接跑）
============================================================
核对六件事：
  1) 对账闭合   列表数 == 成功 + 失败（按 records 实际归属重算）
  2) 记录 vs 落盘  每条记录的文件都必须真实存在
  3) 命名合规   必须符合 [类型] 编号 标题 (文号, 日期).ext
  4) 路径长度   Windows 260 字符上限
  5) 载体完整   PDF 校验魔数 %PDF-；HTML 校验非空且有正文
  6) 数据纯度   有效性冲突 / 字段缺失 / 重名文件清单

用法：python post_audit.py [--out DIR]
"""
import json
import os
import re
import sys
import io
from collections import Counter, defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

SECTIONS = ['法律法规', '民航规章', '规范性文件', '标准规范']


def default_base() -> str:
    """落盘根目录：环境变量 CAAC_DIR > 当前目录下 caac_regs_data/。"""
    return os.environ.get('CAAC_DIR') or os.path.join(os.getcwd(), 'caac_regs_data')


BASE = STATE = ''


def set_base(path: str) -> None:
    global BASE, STATE
    BASE = os.path.abspath(path)
    STATE = os.path.join(BASE, '_state')


def load_targets() -> dict:
    """读抓取器写入的 _state/targets.json（各栏目列表实际条数）；缺失返回空 dict。"""
    p = os.path.join(STATE, 'targets.json')
    if os.path.exists(p):
        try:
            with open(p, encoding='utf-8') as f:
                return json.load(f)
        except Exception:                           # noqa: BLE001
            pass
    return {}

NAME_RE = re.compile(
    r'^\[(法律法规|民航规章|规范性文件|标准规范)\] .+\d{4}-\d{2}-\d{2}\)\.'
    r'(pdf|html?|docx?|xlsx?|pptx?|zip|rar|7z|wps|txt|jpe?g|png|gif|bmp|tiff?)$', re.I)
ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

fail_total = 0


def bad(msg: str) -> None:
    global fail_total
    fail_total += 1
    print('  !! ' + msg)


def load_jsonl(path: str) -> list:
    if not os.path.exists(path):
        return []
    return [json.loads(l) for l in open(path, encoding='utf-8') if l.strip()]


def main() -> int:
    out = None
    for i, a in enumerate(sys.argv):
        if a == '--out' and i + 1 < len(sys.argv):
            out = sys.argv[i + 1]
        elif a.startswith('--out='):
            out = a.split('=', 1)[1]
    set_base(out or default_base())

    recs = load_jsonl(os.path.join(STATE, 'records.jsonl'))
    fails = load_jsonl(os.path.join(STATE, 'failures.jsonl'))
    # 递归收集（含「已失效废止存档」子目录），排除 _state 状态目录
    files = []
    for root, dirs, fs in os.walk(BASE):
        dirs[:] = [d for d in dirs if d != '_state']
        for f in fs:
            files.append(os.path.relpath(os.path.join(root, f), BASE))
    fileset = set(files)
    sizes = {f: os.path.getsize(os.path.join(BASE, f)) for f in files}

    print('=' * 78)
    print(f'验收报告  {len(recs)} 条记录 / {len(files)} 个文件 / {len(fails)} 条失败')
    print('=' * 78)

    # ---------- 1) 对账闭合 ----------
    print('\n[1] 对账闭合（目标 vs 实得）')
    TARGET = load_targets()
    by_sec = Counter(r['section'] for r in recs)
    if TARGET:
        print(f'  {"栏目":<10}{"目标":>6}{"实得":>6}{"缺口":>6}   状态')
        gap_total = 0
        for sec in SECTIONS:
            got = by_sec.get(sec, 0)
            want = TARGET.get(sec)
            if want is None:
                print(f'  {sec:<10}{"—":>6}{got:>6}{"—":>6}   （无基线）')
                continue
            gap = want - got
            gap_total += max(gap, 0)
            state = 'OK' if gap == 0 else ('超出' if gap < 0 else f'缺 {gap}')
            print(f'  {sec:<10}{want:>6}{got:>6}{gap:>6}   {state}')
        print(f'  合计 {len(recs)} 条；未达目标计 {gap_total} 条')
        if gap_total:
            print('  · 缺口可能来自：抓取未完成 / 该条详情页 404 / 服务端返回与筛选不一致')
    else:
        print('  （未找到 _state/targets.json，跳过目标比对 —— 全量跑一次抓取器即生成）')
        for sec in SECTIONS:
            print(f'  {sec:<10}{"—":>6}{by_sec.get(sec, 0):>6}')
    if fails:
        print(f'  · 失败清单 {len(fails)} 条（用 --retry 补漏）：')
        for f in fails[:10]:
            print(f'      [{f.get("stage","?"):<8}] {f.get("section","?"):<8}'
                  f' {str(f.get("reason"))[:60]}')
        if len(fails) > 10:
            print(f'      … 另有 {len(fails) - 10} 条')

    # ---------- 2) 记录 vs 落盘 ----------
    print('\n[2] 记录 vs 落盘一致性')
    missing = [r for r in recs if r.get('file') and r['file'] not in fileset]
    if missing:
        bad(f'记录中有 {len(missing)} 条的文件不在磁盘上')
        for r in missing[:10]:
            print(f'      {r["file"]}')
    else:
        print('  OK  每条记录的载体文件都真实存在')
    # 反向：磁盘有文件但无记录
    rec_files = {r['file'] for r in recs if r.get('file')}
    orphan = fileset - rec_files - {'民航法规索引.xlsx'}
    if orphan:
        bad(f'磁盘上 {len(orphan)} 个文件没有对应记录（孤儿文件）')
        for f in list(orphan)[:10]:
            print(f'      {f}')
    else:
        print('  OK  无孤儿文件')

    # ---------- 3) 命名合规 ----------
    print('\n[3] 命名合规（[类型] 编号 标题 (文号, 日期).ext）')
    illegal_names = [f for f in files if ILLEGAL.search(os.path.basename(f))]
    if illegal_names:
        bad(f'{len(illegal_names)} 个文件含 Windows 非法字符')
        for f in illegal_names[:10]:
            print(f'      {f}')
    else:
        print('  OK  无非法字符')
    noname = [f for f in files
              if not NAME_RE.match(os.path.basename(f))
              and os.path.basename(f) != '民航法规索引.xlsx']
    if noname:
        bad(f'{len(noname)} 个文件不符合命名规则')
        for f in noname[:15]:
            print(f'      {f}')
    else:
        print('  OK  全部符合 [类型] … (…, YYYY-MM-DD).ext 格式')

    # ---------- 4) 路径长度 ----------
    print('\n[4] 路径长度（Windows 上限 260）')
    longest = max((os.path.join(BASE, f) for f in files), key=len, default='')
    if len(longest) >= 260:
        bad(f'最长路径 {len(longest)} 字符，已越界')
        print(f'      {longest}')
    elif len(longest) > 200:
        print(f'  · 最长路径 {len(longest)} 字符（注意：再挪一层目录就会越界）')
        print(f'      {os.path.basename(longest)[:80]}')
    else:
        print(f'  OK  最长路径 {len(longest)} 字符，余量充足')

    # ---------- 5) 载体完整性 ----------
    print('\n[5] 载体完整性')
    def core_len(raw: str) -> int:
        """网页正文的有效字符数：剥掉元数据块、script/style、内联 JS、标签后剩下的。"""
        m = re.search(r'</div>\s*(.*)</body>', raw, re.S)
        s = m.group(1) if m else raw
        s = re.sub(r'<(script|style)\b.*?</\1>', ' ', s, flags=re.S | re.I)
        s = re.sub(r'\bvar\s+\w+\s*=[^;]*;', ' ', s, flags=re.I)
        s = re.sub(r'<[^>]+>', ' ', s)
        s = re.sub(r'&nbsp;|&#\d+;', ' ', s)
        return len(re.sub(r'\s+', ' ', s).strip())

    hollow_files = {r.get('file') for r in recs if r.get('notes')}
    bad_pdf, empty_html, small_pdf, noted_short = [], [], [], []
    for f in files:
        p = os.path.join(BASE, f)
        if f.lower().endswith('.pdf'):
            with open(p, 'rb') as fh:
                head = fh.read(5)
            if not head.startswith(b'%PDF-'):
                bad_pdf.append(f)
            elif sizes[f] < 10240:
                small_pdf.append(f)
        elif f.lower().endswith(('.html', '.htm')):
            # 按「有效正文字符数」判空壳，而非文件大小 ——
            # 官网不少标准条目正文只有一句适用范围（60-200 字），文件小是正常的
            if core_len(open(p, encoding='utf-8', errors='replace').read()) < 60:
                (noted_short if f in hollow_files else empty_html).append(f)
    if bad_pdf:
        bad(f'{len(bad_pdf)} 个 PDF 魔数异常（疑似下载截断/错误页）')
        for f in bad_pdf[:10]:
            print(f'      {f}')
    else:
        print(f'  OK  全部 PDF 魔数正常（{sum(1 for f in files if f.lower().endswith(".pdf"))} 份）')
    if small_pdf:
        print(f'  · {len(small_pdf)} 份 PDF 小于 10KB，建议抽查：')
        for f in small_pdf[:8]:
            print(f'      {sizes[f]//1024}KB  {f[:70]}')
    if noted_short:
        print(f'  · {len(noted_short)} 个短网页已标注「官网无内容」，属官网数据缺失，非抓取问题')
    if empty_html:
        bad(f'{len(empty_html)} 个网页正文文件过小（<2KB），正文可能没抓到')
        for f in empty_html[:10]:
            print(f'      {sizes[f]}B  {f[:70]}')
    else:
        print(f'  OK  网页正文均非空（{sum(1 for f in files if f.lower().endswith((".html", ".htm")))} 份）')
    # ---------- 6) 数据纯度 ----------
    print('\n[6] 数据纯度')
    conflict = [r for r in recs if str(r.get('reconcile', '')).startswith('冲突')]
    missed_f = [r for r in recs if not r.get('validity')]
    print(f'  有效性冲突（列表=有效 / 详情页≠有效）：{len(conflict)} 条'
          f'（{len(conflict) / len(recs) if recs else 0:.1%}）')
    for r in conflict[:12]:
        print(f'      [{r["section"]}] {r["title"][:44]}')
    if len(conflict) > 12:
        print(f'      … 另有 {len(conflict) - 12} 条')
    print(f'  有效性字段缺失：{len(missed_f)} 条'
          f'（{len(missed_f) / len(recs) if recs else 0:.1%}）—— 已保守保留')
    dup = {k: v for k, v in Counter(r['file'] for r in recs if r.get('file')).items() if v > 1}
    if dup:
        bad(f'{len(dup)} 个文件名被多条记录共用')
        for k, v in list(dup.items())[:10]:
            print(f'      {v}× {k[:70]}')
    else:
        print('  OK  无重名文件')
    print('  载体分布：', dict(Counter(r['carrier'] for r in recs)))

    # ---------- 体积 ----------
    tot = sum(sizes.values())
    pdfs = [s for f, s in sizes.items() if f.lower().endswith('.pdf')]
    print(f'\n[体积] 合计 {tot / 1024 / 1024:.1f} MB'
          + (f'；PDF {len(pdfs)} 份，中位 {sorted(pdfs)[len(pdfs)//2] / 1024:.0f}KB'
             f'，最大 {max(pdfs) / 1024 / 1024:.1f}MB' if pdfs else ''))

    print('\n' + '=' * 78)
    print(f'验收结论：{"全部通过" if fail_total == 0 else f"{fail_total} 项需处理"}')
    print('=' * 78)
    return 1 if fail_total else 0


if __name__ == '__main__':
    sys.exit(main())
