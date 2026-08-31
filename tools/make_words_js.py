# -*- coding: utf-8 -*-
"""把词库 xls 生成安卓端单词数据 words.js。

与 Windows 版 app.py 的 load_words() 逻辑保持一致：
- 自动跳过表头行（序号/单词/注音/释义/word/meaning…等），
  兼容「序号在前」和「无序号」两种表格式；
- 序号列数字则单词在 1 列，否则单词在 0 列；
- 释义取该行剩余列里最后一个非空列。

用法:
    python make_words_js.py [输入xls] [输出js]

默认输入:  ../02.考研英语词汇乱序版.xls
默认输出:  ../android/app/src/main/assets/words.js
生成格式:  window.WORDS=[["单词","注音","释义"], ...];
"""
import json
import os
import sys

try:
    import xlrd
except ImportError:
    sys.exit('缺少依赖 xlrd，请先：pip install xlrd')

HEADER_WORDS = {'序号', '单词', 'word', '音标', '注音', '释义', 'meaning',
                'phonetic', '英文', '中文'}


def _is_header(row):
    return str(row[0]).strip().lower() in {h.lower() for h in HEADER_WORDS}


def _first_col_numeric(sh, start):
    end = min(start + 20, sh.nrows)
    for r in range(start, end):
        try:
            int(float(str(sh.cell_value(r, 0)).strip()))
        except (ValueError, TypeError):
            return False
    return end > start


def make_rows(xls_path):
    wb = xlrd.open_workbook(xls_path)
    sh = None
    for name in wb.sheet_names():                 # 取第一个有内容的表
        s = wb.sheet_by_name(name)
        if s.nrows > 0 and s.ncols > 0:
            sh = s
            break
    if sh is None:
        raise RuntimeError('词库文件为空：%s' % xls_path)

    start = 0
    while start < sh.nrows and _is_header(sh.row_values(start)):
        start += 1
    numeric_index = _first_col_numeric(sh, start)

    rows = []
    for r in range(start, sh.nrows):
        row = sh.row_values(r)
        if len(row) < 2:
            continue
        if numeric_index:
            word = str(row[1]).strip()
            phonetic = str(row[2]).strip()
            meaning = str(row[3]).strip() if len(row) > 3 else ''
        else:
            word = str(row[0]).strip()
            phonetic = str(row[1]).strip()
            meaning = ''
            for c in range(len(row) - 1, 1, -1):
                v = str(row[c]).strip()
                if v:
                    meaning = v
                    break
        if not word:
            continue
        rows.append([word, phonetic, meaning])
    return rows


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, '..', '02.考研英语词汇乱序版.xls')
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        here, '..', 'android', 'app', 'src', 'main', 'assets', 'words.js')

    rows = make_rows(src)
    js = 'window.WORDS=' + json.dumps(rows, ensure_ascii=False) + ';'
    with open(out, 'w', encoding='utf-8', newline='') as f:
        f.write(js)
    print('已生成 %d 词 -> %s' % (len(rows), out))


if __name__ == '__main__':
    main()
