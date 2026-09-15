#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
抓取完成后的整理（跑一次即可）
============================================================
做两件事：
  1) 把「列表标有效、详情页却标失效/废止」的冲突件，从主目录移入
     子目录 _已失效废止_存档 ，主目录只留真正现行有效
  2) 索引追加一张独立的「现行有效」sheet（原「索引」表保留全量含冲突件）

为什么需要：规范性文件栏目下，官网检索索引的"有效性"字段明显滞后于
详情页 —— 服务端 selYouxiao=有效 筛出来的条目里，约 1/4 在详情页标着
失效/废止。用详情页字段做二次整理，才能兑现"只留现行有效"。

用法：python organize.py [--out DIR]           # 干跑，只报告不移动
      python organize.py [--out DIR] --apply   # 实际执行
"""
import io
import json
import os
import shutil
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

ARCHIVE_NAME = '民航法规_已失效废止存档'
COLS = ['类型', '编号', '标题', '文号', '子类', '发布单位', '成文日期', '有效性',
        '载体', '对账', '官方原名', '文件名', '来源URL', '抓取时间']
KEYS = ['section', 'code', 'title', 'wenhao', 'subcat', 'unit', 'date', 'validity',
        'carrier', 'reconcile', 'official', 'file', 'url', 'fetched']
COL_W = [12, 16, 54, 26, 12, 20, 12, 10, 8, 14, 56, 72, 60, 19]
C_FILE, C_URL = 12, 13

BASE = RECORDS_F = INDEX_F = ARCHIVE = ''


def default_base() -> str:
    """落盘根目录：环境变量 CAAC_DIR > 当前目录下 caac_regs_data/。"""
    return os.environ.get('CAAC_DIR') or os.path.join(os.getcwd(), 'caac_regs_data')


def set_base(path: str) -> None:
    global BASE, RECORDS_F, INDEX_F, ARCHIVE
    BASE = os.path.abspath(path)
    RECORDS_F = os.path.join(BASE, '_state', 'records.jsonl')
    INDEX_F = os.path.join(BASE, '民航法规索引.xlsx')
    ARCHIVE = os.path.join(BASE, ARCHIVE_NAME)


def load_jsonl(path: str) -> list:
    if not os.path.exists(path):
        return []
    out = []
    for line in open(path, encoding='utf-8'):
        line = line.strip()
        if line:
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                pass
    return out


def write_jsonl(path: str, rows: list) -> None:
    with open(path, 'w', encoding='utf-8') as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + '\n')


def main() -> int:
    apply_mode = '--apply' in sys.argv
    out = None
    for i, a in enumerate(sys.argv):
        if a == '--out' and i + 1 < len(sys.argv):
            out = sys.argv[i + 1]
        elif a.startswith('--out='):
            out = a.split('=', 1)[1]
    set_base(out or default_base())
    recs = load_jsonl(RECORDS_F)
    if not recs:
        print('!! records.jsonl 为空，抓取可能还没跑完。')
        return 1

    conf = [r for r in recs if str(r.get('reconcile', '')).startswith('冲突')]
    print('=' * 70)
    print(f'整理模式：{"实际执行" if apply_mode else "干跑预览（加 --apply 才动手）"}')
    print(f'记录 {len(recs)} 条，冲突件 {len(conf)} 条 '
          f'（{len(conf) / len(recs):.1%}）')
    print('=' * 70)

    # ---- 1) 移动冲突件 ----
    os.makedirs(ARCHIVE, exist_ok=True)
    moved = already = missing = 0
    for r in conf:
        name = os.path.basename(str(r.get('file') or ''))
        if not name:
            missing += 1
            continue
        src = os.path.join(BASE, name)
        dst = os.path.join(ARCHIVE, name)
        if os.path.exists(src):
            if apply_mode:
                shutil.move(src, dst)
            moved += 1
        elif os.path.exists(dst):
            already += 1
        else:
            missing += 1
            print(f'  !! 文件不存在，记录与磁盘不符：{name[:70]}')
        r['file'] = os.path.join(ARCHIVE_NAME, name)
        r['archived'] = True
    print(f'\n[1] 冲突件归档')
    print(f'  {"将移动" if not apply_mode else "已移动"} {moved} 份 → {ARCHIVE_NAME}\\')
    print(f'  已在存档 {already} 份；文件缺失 {missing} 份')

    if apply_mode:
        write_jsonl(RECORDS_F, recs)
        print(f'  records.jsonl 已更新（冲突件路径已改为存档相对路径）')

    # ---- 2) 索引：更新路径列 + 追加「现行有效」sheet ----
    print(f'\n[2] 索引整理')
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print('  !! 缺少 openpyxl，跳过。')
        return 1

    try:
        wb = load_workbook(INDEX_F)
    except Exception as e:                              # noqa: BLE001
        print(f'  !! 打开索引失败（很可能正被预览/Excel 占用）：{e}')
        print('     请关闭对「民航法规索引.xlsx」的预览后重跑。')
        return 1

    arch_urls = {r['url'] for r in conf}
    ws = wb['索引']
    fixed = 0
    for ri in range(2, ws.max_row + 1):
        if ws.cell(ri, C_URL).value in arch_urls:
            name = os.path.basename(str(ws.cell(ri, C_FILE).value or ''))
            ws.cell(ri, C_FILE, os.path.join(ARCHIVE_NAME, name))
            fixed += 1
    print(f'  索引表路径列更新 {fixed} 行')

    if '现行有效' in wb.sheetnames:
        del wb['现行有效']
    ws3 = wb.create_sheet('现行有效')
    head_fill = PatternFill('solid', fgColor='3C3489')
    head_font = Font(color='FFFFFF', bold=True, size=10)
    for ci, (label, w) in enumerate(zip(COLS, COL_W), 1):
        c = ws3.cell(1, ci, label)
        c.fill, c.font = head_fill, head_font
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = 'A2'
    ri = 2
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[C_URL - 1] in arch_urls:
            continue
        for ci, v in enumerate(row, 1):
            ws3.cell(ri, ci, v)
        ri += 1
    valid_n = ri - 2
    print(f'  「现行有效」sheet：{valid_n} 条（原表保留全量 {ws.max_row - 1} 条）')
    print(f'  冲突件（已入存档）：{len(conf)} 条，' 
          f'校验 {valid_n} + {len(conf)} = {valid_n + len(conf)} '
          f'{"==" if valid_n + len(conf) == ws.max_row - 1 else "!="} {ws.max_row - 1}')

    if not apply_mode:
        print('  （干跑：索引未写入）')
    else:
        try:
            wb.save(INDEX_F)
            print(f'  已保存：{INDEX_F}')
        except Exception as e:                          # noqa: BLE001
            print(f'  !! 保存失败（文件可能被占用）：{e}')
            print('     请关闭对「民航法规索引.xlsx」的预览后重跑。')
            return 1

    print('\n' + '=' * 70)
    print(f'整理完成：主目录 {len(recs) - len(conf)} 条现行有效 + '
          f'存档 {len(conf)} 条')
    print('=' * 70)
    return 0


if __name__ == '__main__':
    sys.exit(main())
