#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
caac_crawl 解析逻辑回归测试（零依赖，直接运行）

    python test_parse.py

背景：曾出现过一个静默 bug —— 详情页里 PDF 文件名既出现在真实 <a href> 中，
也出现在内联 JS 字符串（href=\"./P0....pdf\"）中。早期正则没锚定 <a> 标签，
于是把后面的 JS 代码当成"官方文件名"，污染了 103 行索引数据。
这里用合成陷阱 + 4 个真实页面快照把它钉住。
"""
import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'scripts'))

import caac_crawl as C   # noqa: E402

FIX = os.path.join(HERE, 'fixtures')
PASS, FAIL = 0, []


def check(name, cond, detail=''):
    global PASS
    if cond:
        PASS += 1
        print(f'  ok   {name}')
    else:
        FAIL.append(name)
        print(f'  FAIL {name}   {detail}')


def load(fn):
    return open(os.path.join(FIX, fn), encoding='utf-8').read()


# ---------------------------------------------------------------- 1. 合成陷阱
print('\n[1] 合成陷阱：JS 字符串里的 href 不得被当成锚文本')

TRAP = '''<html><body>
<div class="content" data-role="n_content" ><p>正文一段</p></div>
<script>
  $(function(){ if(n>0){ h = "href=\\"./P020260717459628983882.pdf\\""; $("#x").show(); } });
</script>
<div id="id_tblAppendix">
  <p><a href="./P020260717459628983882.pdf" target="_blank">通用航空经营管理规定（中华人民共和国交通运输部令2026年第14号）.pdf</a></p>
</div>
</body></html>'''

m = C.parse_detail(TRAP)
check('附件名提取正确', m['atts'] == ['P020260717459628983882.pdf'], m['atts'])
check('官方原名 = 锚文本', m['official'].startswith('通用航空经营管理规定'), repr(m['official']))
check('官方原名不含 JS 噪声', not any(t in m['official'] for t in C.JS_NOISE), repr(m['official']))
check('官方原名结尾是 .pdf', m['official'].endswith('.pdf'), repr(m['official']))
check('正文容器抽取', '正文一段' in m['body'], repr(m['body'][:40]))

# ---------------------------------------------------------------- 2. 真实页面
print('\n[2] 真实页面快照（2026-09-15 抓取）')

FIXTURES = [
    ('detail_民航规章.html', 'CCAR-290-R4', '有效', '通用航空经营管理规定', True),
    ('detail_标准规范.html', '', '有效', '空中交通无线电通话用语', True),
    ('detail_规范性文件.html', '', '有效', '型号合格审定程序', False),
    ('detail_法律法规.html', '', '', '民用航空法', False),
]

for fn, code, yx, title_word, need_pdf in FIXTURES:
    if not os.path.exists(os.path.join(FIX, fn)):
        check(f'{fn} 存在', False, 'fixture 缺失')
        continue
    m = C.parse_detail(load(fn))
    print(f'  -- {fn}: 部号={m["buhao"]!r} 有效性={m["youxiao"]!r} 附件={m["atts"]}')
    check(f'{fn} 标题/正文非空', bool(m['body']) or bool(m['atts']))
    check(f'{fn} 无 JS 噪声污染',
          not any(t in m['official'] for t in C.JS_NOISE), repr(m['official']))
    if code:
        check(f'{fn} 部号={code}', m['buhao'] == code, repr(m['buhao']))
    if yx:
        check(f'{fn} 有效性={yx}', m['youxiao'] == yx, repr(m['youxiao']))
    if need_pdf:
        check(f'{fn} 有 PDF 附件', len(m['atts']) >= 1, m['atts'])
        check(f'{fn} 官方原名以 .pdf 结尾', m['official'].endswith('.pdf'), repr(m['official']))
    else:
        check(f'{fn} 官方原名不含 JS 噪声', 'function' not in m['official'], repr(m['official']))

# 法律法规页：正文应当是长文（无附件时靠正文承载）
if os.path.exists(os.path.join(FIX, 'detail_法律法规.html')):
    m = C.parse_detail(load('detail_法律法规.html'))
    check('法律法规正文长度 > 10000 字', len(m['body']) > 10000, str(len(m['body'])))

# ---------------------------------------------------------------- 3. 命名规则
print('\n[3] 命名规则')
if os.path.exists(os.path.join(FIX, 'detail_民航规章.html')):
    m = C.parse_detail(load('detail_民航规章.html'))
    name = C.build_name('民航规章', '通用航空经营管理规定', m, True)
    print(f'  -> {name}')
    check('命名以 [类型] 开头', name.startswith('[民航规章]'), name)
    check('命名含部号', 'CCAR-290-R4' in name, name)
    check('命名含文号', '交通运输部令2026年第14号' in name, name)
    check('命名含日期', '2026-07-01' in name, name)
    check('命名无非法字符', not any(ch in name for ch in '\\/:*?"<>|'), name)

# 无附件的网页正文走 .html
if os.path.exists(os.path.join(FIX, 'detail_法律法规.html')):
    m = C.parse_detail(load('detail_法律法规.html'))
    name = C.build_name('法律法规', '中华人民共和国民用航空法（2026年7月1日起施行）', m, False)
    check('网页正文用 .html 后缀', name.endswith('.html'), name)

# ---------------------------------------------------------------- 结果
print('\n' + '=' * 60)
print(f'通过 {PASS} 项' + (f'，失败 {len(FAIL)} 项：{FAIL}' if FAIL else '，全部通过'))
print('=' * 60)
sys.exit(1 if FAIL else 0)
