# -*- coding: utf-8 -*-
"""
考研英语背单词（桌面版）
=======================
- 独立窗口运行，无需浏览器
- 词库：02.考研英语词汇乱序版.xls（5493 个词，3 列：单词/注音/释义，自动兼容带序号 4 列格式）
- 不认识(按钮或按 3)的单词自动追加到 不认识的单词.xls，不重复记录；认识(按 2)则从生词本删除
- 学习进度自动保存到 学习进度.txt
- 打包成 exe 后：词库打进 exe 内，生词本与进度保存在 exe 同目录
- 坚果云网盘同步：电脑与手机共用同一 WebDAV 文件，全手动——打开时弹窗询问是否从云端获取，
  主界面底部「☁ 上传」「☁ 从云端获取」按钮手动同步并显示上次同步时间，关闭时若有未上传改动会弹窗确认；
  其余操作不联网，每次操作时自动把进度和生词本做一份本地备份（保留最近 10 条，可恢复）

打包：pyinstaller --onefile --windowed --name kaoyan_words --add-data "02.考研英语词汇乱序版.xls;." app.py
"""
import base64
import datetime
import re
import json
import os
import sys
import threading
import time
import traceback
import urllib.error
import urllib.request

import xlrd
import xlwt
import tkinter as tk
from tkinter import filedialog, messagebox

# ------------------------------------------------------------------ 路径
# 打包成 exe（PyInstaller --onefile --windowed）后：
#   RESOURCE_DIR = sys._MEIPASS  → 只读资源（词库 xls）解压在临时目录
#   DATA_DIR     = exe 所在目录   → 可写数据（生词本、进度、日志）保存在 exe 旁边
if getattr(sys, 'frozen', False):
    RESOURCE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    DATA_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    RESOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = RESOURCE_DIR

SOURCE_FILE = os.path.join(RESOURCE_DIR, '02.考研英语词汇乱序版.xls')
UNKNOWN_FILE = os.path.join(DATA_DIR, '不认识的单词.xls')
PROGRESS_FILE = os.path.join(DATA_DIR, '学习进度.txt')
LOG_FILE = os.path.join(DATA_DIR, '背单词.log')
CLOUD_CONFIG_FILE = os.path.join(DATA_DIR, '网盘同步配置.txt')
SYNC_STATUS_FILE = os.path.join(DATA_DIR, '同步状态.json')
BACKUP_DIR = os.path.join(DATA_DIR, '本地备份')

lock = threading.RLock()   # 可重入锁：同步合并内部会嵌套持锁

# ------------------------------------------------------------------ 日志
def log(msg):
    """无控制台窗口的 exe 也看不到 print，统一写到日志文件。"""
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write('%s\n' % msg)
    except Exception:
        pass


# ---------------------------------------------------------------- 原始词库
HEADER_WORDS = {'序号', '单词', 'word', '音标', '注音', '释义', 'meaning', 'phonetic', '英文', '中文'}


def _first_data_sheet(wb):
    for name in wb.sheet_names():
        sh = wb.sheet_by_name(name)
        if sh.nrows > 0 and sh.ncols > 0:
            return sh
    raise RuntimeError('词库文件为空：%s' % SOURCE_FILE)


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


def load_words():
    wb = xlrd.open_workbook(SOURCE_FILE)
    sh = _first_data_sheet(wb)
    start = 0
    while start < sh.nrows and _is_header(sh.row_values(start)):
        start += 1
    numeric_index = _first_col_numeric(sh, start)
    words = []
    for r in range(start, sh.nrows):
        row = sh.row_values(r)
        if len(row) < 2:
            continue
        if numeric_index:
            word, phonetic = str(row[1]).strip(), str(row[2]).strip()
            meaning = str(row[3]).strip() if len(row) > 3 else ''
        else:
            word, phonetic = str(row[0]).strip(), str(row[1]).strip()
            meaning = ''
            for c in range(len(row) - 1, 1, -1):
                v = str(row[c]).strip()
                if v:
                    meaning = v
                    break
        if not word:
            continue
        words.append({
            'index': len(words),
            'word': word,
            'phonetic': phonetic,
            'meaning': meaning,
        })
    return words


WORDS = load_words()
WORD_COUNT = len(WORDS)

# 单词+注音 → 原始序号，供生词本去重 / 同步合并用
WORD_KEY_TO_INDEX = {}
for _w in WORDS:
    WORD_KEY_TO_INDEX.setdefault((_w['word'], _w['phonetic']), _w['index'])

# ---------------------------------------------------------------- 生词本
unknown_rows = []        # [{'word','phonetic','meaning'}, ...] 顺序即序号
added_indices = set()    # 已记录过的原始序号，避免同一词重复写入


def load_unknown():
    global unknown_rows, added_indices
    unknown_rows = []
    added_indices = set()
    if not os.path.exists(UNKNOWN_FILE):
        return
    try:
        wb = xlrd.open_workbook(UNKNOWN_FILE)
        sh = wb.sheet_by_index(0)
        for r in range(1, sh.nrows):
            row = sh.row_values(r)
            if len(row) < 4:
                continue
            unknown_rows.append({
                'word': str(row[1]).strip(),
                'phonetic': str(row[2]).strip(),
                'meaning': str(row[3]).strip(),
            })
        key_to_index = {}
        for w in WORDS:
            key_to_index.setdefault((w['word'], w['phonetic']), w['index'])
        for row in unknown_rows:
            idx = key_to_index.get((row['word'], row['phonetic']))
            if idx is not None:
                added_indices.add(idx)
    except Exception as e:
        log('读取生词本失败（忽略，将重建）：%s' % e)


def save_unknown():
    wb = xlwt.Workbook(encoding='utf-8')
    sh = wb.add_sheet('单词')
    for c, h in enumerate(['序号', '单词', '注音', '释义']):
        sh.write(0, c, h)
    for i, row in enumerate(unknown_rows, start=1):
        sh.write(i, 0, i)
        sh.write(i, 1, row['word'])
        sh.write(i, 2, row['phonetic'])
        sh.write(i, 3, row['meaning'])
    wb.save(UNKNOWN_FILE)
    make_local_backup()   # 每次操作后自动本地备份（不联网），保留最近 10 条


def _rebuild_added_indices():
    """根据 unknown_rows 重建 added_indices。"""
    added_indices.clear()
    for row in unknown_rows:
        idx = WORD_KEY_TO_INDEX.get((row['word'], row['phonetic']))
        if idx is not None:
            added_indices.add(idx)


def add_unknown(pos):
    """不认识：加入生词本（幂等）。返回 {'ok','already','unknown_total'}。"""
    w = WORDS[pos]
    with lock:
        saved_rows = list(unknown_rows)
        already = w['index'] in added_indices
        if not already:
            unknown_rows.append({
                'word': w['word'],
                'phonetic': w['phonetic'],
                'meaning': w['meaning'],
            })
            added_indices.add(w['index'])
        try:
            save_unknown()
        except Exception as e:
            unknown_rows[:] = saved_rows
            _rebuild_added_indices()
            return {'ok': False, 'error': '写入生词本失败（可能文件被占用）：%s' % e}
    return {'ok': True, 'already': already, 'unknown_total': len(unknown_rows)}


def mark_word_know(pos):
    """认识：从生词本删除（幂等）。返回 {'ok','removed','error'}。"""
    w = WORDS[pos]
    with lock:
        saved_rows = list(unknown_rows)
        removed = False
        for i, row in enumerate(unknown_rows):
            if row['word'] == w['word'] and row['phonetic'] == w['phonetic']:
                del unknown_rows[i]
                added_indices.discard(w['index'])
                removed = True
                break
        try:
            save_unknown()
        except Exception as e:
            unknown_rows[:] = saved_rows
            _rebuild_added_indices()
            return {'ok': False, 'error': '写入生词本失败（可能文件被占用）：%s' % e}
    return {'ok': True, 'removed': removed}


def load_progress():
    try:
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            return int(f.read().strip())
    except Exception:
        return 0


def save_progress(pos):
    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            f.write(str(pos))
    except Exception as e:
        log('保存进度失败：%s' % e)
        return
    make_local_backup()   # 每次操作后自动本地备份（不联网），保留最近 10 条


load_unknown()

def export_backup_file(path):
    """把进度+生词本导出为 JSON 备份文件。返回错误信息或 None。"""
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({'version': 1, 'progress': load_progress(), 'unknown': unknown_rows},
                      f, ensure_ascii=False, indent=1)
        return None
    except Exception as e:
        return '导出失败：%s' % e


def _apply_state(prog, rows):
    """把 (进度, 生词行) 整体写入本地数据文件（备份导入 / 云端拉取共用）。"""
    global unknown_rows, added_indices
    prog = max(0, min(int(prog), WORD_COUNT - 1))
    new_rows = []
    new_indices = set()
    for item in rows:
        w = str(item.get('word', '')).strip()
        if not w:
            continue
        ph = str(item.get('phonetic', '')).strip()
        me = str(item.get('meaning', '')).strip()
        new_rows.append({'word': w, 'phonetic': ph, 'meaning': me})
        idx = WORD_KEY_TO_INDEX.get((w, ph))
        if idx is not None:
            new_indices.add(idx)
    with lock:
        save_progress(prog)
        unknown_rows = new_rows
        added_indices = new_indices
        save_unknown()


def import_backup_file(path):
    """从 JSON 备份文件导入进度+生词本（整体覆盖）。返回错误信息或 None。"""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        prog = int(data.get('progress', 0) or 0)
        _apply_state(prog, data.get('unknown') or [])
        return None
    except Exception as e:
        return '导入失败：%s' % e


# ---------------------------------------------------------------- 本地备份
# 每次操作（翻页/认识/不认识等）及关闭程序时，自动把进度+生词本存一份到 exe 旁的
# 「本地备份」文件夹，保留最近 10 条（纯本地、不联网）。可在「本地备份」弹窗里恢复，
# 也可直接用「导入备份」选择备份文件恢复。
BACKUP_KEEP = 10


def make_local_backup():
    """把当前进度+生词本存为本地备份文件。返回 (文件名, 错误)。"""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        # Windows 系统时钟精度约 15ms，连续两次关闭可能撞同名，冲突时自动追加序号
        base = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        name = '备份_%s.json' % base
        path = os.path.join(BACKUP_DIR, name)
        n = 1
        while os.path.exists(path):
            name = '备份_%s_%d.json' % (base, n)
            path = os.path.join(BACKUP_DIR, name)
            n += 1
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({'version': 1, 'progress': load_progress(), 'unknown': unknown_rows,
                       'backup_time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')},
                      f, ensure_ascii=False, indent=1)
        prune_local_backups(BACKUP_KEEP)
        return name, None
    except Exception as e:
        return None, '本地备份失败：%s' % e


def _backup_sort_key(name):
    """备份文件排序键：先按文件修改时间，mtime 相同（Windows 时钟 ~15ms 精度，
       同一瞬间连写）时按文件名内嵌的创建时间戳+序号排，保持真实创建顺序。"""
    try:
        mt = os.path.getmtime(os.path.join(BACKUP_DIR, name))
    except Exception:
        mt = 0.0
    m = re.match(r'^备份_(\d{8}_\d{6}_\d{6})(?:_(\d+))?\.json$', name)
    if m:
        idx = int(m.group(2)) if m.group(2) else 0
        return (mt, m.group(1), idx)
    return (mt, name, 0)


def prune_local_backups(keep=BACKUP_KEEP):
    """本地备份目录只保留最近 keep 条（新→旧）。"""
    try:
        if not os.path.isdir(BACKUP_DIR):
            return
        files = [os.path.join(BACKUP_DIR, n) for n in os.listdir(BACKUP_DIR)
                 if n.startswith('备份_') and n.endswith('.json')]
        files.sort(key=_backup_sort_key, reverse=True)
        for p in files[keep:]:
            try:
                os.remove(p)
            except Exception:
                pass
    except Exception:
        pass


def list_local_backups():
    """返回最近 BACKUP_KEEP 条本地备份（新→旧）：
       [{name, path, time, progress, unknown_total}]"""
    out = []
    try:
        if not os.path.isdir(BACKUP_DIR):
            return out
        names = [n for n in os.listdir(BACKUP_DIR)
                 if n.startswith('备份_') and n.endswith('.json')]
        names.sort(key=_backup_sort_key, reverse=True)
        for n in names[:BACKUP_KEEP]:
            p = os.path.join(BACKUP_DIR, n)
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                out.append({
                    'name': n, 'path': p,
                    'time': d.get('backup_time') or time.strftime(
                        '%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(p))),
                    'progress': int(d.get('progress', 0) or 0),
                    'unknown_total': len(d.get('unknown') or []),
                })
            except Exception:
                out.append({'name': n, 'path': p, 'time': '?', 'progress': 0, 'unknown_total': 0})
    except Exception:
        pass
    return out


# ---------------------------------------------------------------- 坚果云网盘同步
# 规则（用户定）：打开电脑/手机时从云端取最新数据；关闭电脑/手机切后台或退出时传回云端；其余操作不联网。
# 云端只放一份 JSON（即备份文件格式），电脑与手机共用同一个 WebDAV 地址。
def cloud_config():
    """读取坚果云 WebDAV 配置；未配置或字段不全返回 None。"""
    try:
        with open(CLOUD_CONFIG_FILE, 'r', encoding='utf-8') as f:
            d = json.load(f)
        if d.get('url') and d.get('user') and d.get('password'):
            return {'url': str(d['url']).strip(), 'user': str(d['user']).strip(),
                    'password': str(d['password']).strip()}
    except Exception:
        pass
    return None


def save_cloud_config(url, user, password):
    with open(CLOUD_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump({'url': str(url).strip(), 'user': str(user).strip(),
                   'password': str(password).strip()}, f, ensure_ascii=False, indent=1)


def clear_cloud_config():
    try:
        if os.path.exists(CLOUD_CONFIG_FILE):
            os.remove(CLOUD_CONFIG_FILE)
    except Exception:
        pass


def _cloud_http(method, url, user, password, body=None, timeout=15):
    """坚果云 WebDAV 请求（HTTP Basic 认证）。返回 (status, content)。"""
    req = urllib.request.Request(url, data=body, method=method)
    auth = base64.b64encode(('%s:%s' % (user, password)).encode('utf-8')).decode('ascii')
    req.add_header('Authorization', 'Basic ' + auth)
    if body is not None:
        req.add_header('Content-Type', 'application/json')
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.status, r.read().decode('utf-8', 'replace')


def cloud_fetch():
    """GET 云端数据。返回 (data_or_None, err)：
       data=None 且 err=None 表示云端还没有文件（404）；data=dict 表示取到；err 非空表示失败。"""
    cfg = cloud_config()
    if not cfg:
        return None, '未配置坚果云网盘同步'
    try:
        _, content = _cloud_http('GET', cfg['url'], cfg['user'], cfg['password'])
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None, None
        return None, '获取云端数据失败：HTTP %s' % e.code
    except Exception as e:
        return None, '获取云端数据失败：%s' % e
    try:
        return json.loads(content), None
    except Exception as e:
        return None, '云端数据格式错误：%s' % e


def cloud_push():
    """PUT 本端数据到云端。返回错误信息或 None（成功）。"""
    cfg = cloud_config()
    if not cfg:
        return '未配置坚果云网盘同步'
    payload = json.dumps({'version': 1, 'progress': load_progress(), 'unknown': unknown_rows},
                         ensure_ascii=False).encode('utf-8')
    try:
        status, _ = _cloud_http('PUT', cfg['url'], cfg['user'], cfg['password'], body=payload)
        if status not in (200, 201, 204):
            return '上传云端失败：HTTP %s' % status
        return None
    except urllib.error.HTTPError as e:
        return '上传云端失败：HTTP %s' % e.code
    except Exception as e:
        return '上传云端失败：%s' % e


def apply_cloud_data(data):
    """把云端数据整体写入本地（打开时取云端）。
       云端为空（进度 0 且无生词）时视为尚未同步，保留本地防止误清空。
       返回 (是否应用, 说明文字)。"""
    if data is None:
        return False, '云端暂无数据，保持本地进度和生词'
    try:
        cprog = int(data.get('progress', 0) or 0)
    except (TypeError, ValueError):
        cprog = 0
    crows = [it for it in (data.get('unknown') or [])
             if isinstance(it, dict) and str(it.get('word', '')).strip()]
    if cprog <= 0 and not crows:
        return False, '云端为空，保持本地进度和生词'
    _apply_state(cprog, crows)
    return True, '已从云端获取：%s' % local_state_desc()


def cloud_state_desc(data):
    """把云端数据描述成「第 X / 5493 词，生词 N 个」；云端无数据返回「暂无数据」。"""
    if not data:
        return '暂无数据'
    try:
        p = int(data.get('progress', 0) or 0)
    except (TypeError, ValueError):
        p = 0
    rows = [it for it in (data.get('unknown') or [])
            if isinstance(it, dict) and str(it.get('word', '')).strip()]
    if p <= 0 and not rows:
        return '暂无数据'
    return '第 %d / %d 词，生词 %d 个' % (min(p + 1, WORD_COUNT), WORD_COUNT, len(rows))


def local_state_desc():
    """描述本机当前进度与生词数（与界面一致，第 X 词从 1 开始）。"""
    return '第 %d / %d 词，生词 %d 个' % (
        min(load_progress() + 1, WORD_COUNT), WORD_COUNT, len(unknown_rows))


# ---------------------------------------------------------------- 同步状态（手动同步）
# 全手动同步规则：
#   · 打开时弹窗询问「是否从云端获取」；
#   · 主界面底部「☁ 上传」「☁ 从云端获取」按钮手动同步，旁边标注上次同步时间；
#   · 关闭时若有未上传到云端的改动（dirty），弹窗三选：上传并退出 / 直接退出 / 继续学习。
# 同步状态存 同步状态.json：{last_sync, baseline_progress, baseline_unknown}
# baseline 是「本机数据与云端一致的基线」，每次成功获取/上传后更新。
# dirty = 当前本机数据 ≠ 基线（从未同步过也算有改动）。
def load_sync_status():
    try:
        with open(SYNC_STATUS_FILE, 'r', encoding='utf-8') as f:
            d = json.load(f)
            if isinstance(d, dict):
                return d
    except Exception:
        pass
    return {}


def save_sync_status(**kw):
    d = load_sync_status()
    d.update(kw)
    try:
        with open(SYNC_STATUS_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def _current_state_json():
    return json.dumps({'progress': load_progress(), 'unknown': unknown_rows},
                      ensure_ascii=False, sort_keys=True)


def is_dirty():
    """本机数据相对上次成功云端同步（获取或上传）是否有未上传的变化。"""
    st = load_sync_status()
    if 'baseline_progress' not in st:
        return True    # 从未成功同步过 → 有可上传的内容
    base = json.dumps({'progress': st['baseline_progress'],
                       'unknown': st.get('baseline_unknown') or []},
                      ensure_ascii=False, sort_keys=True)
    return _current_state_json() != base


def mark_synced():
    """成功从云端获取或上传后调用：更新基线并记录上次同步时间。"""
    save_sync_status(last_sync=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                     baseline_progress=load_progress(),
                     baseline_unknown=unknown_rows)


def last_sync_str():
    """上次成功同步的时间；从未同步过返回 None。"""
    return load_sync_status().get('last_sync')


# ---------------------------------------------------------------- 界面配色
FONT = 'Microsoft YaHei'
C_HEADER_1 = '#667eea'      # 头部渐变起始
C_HEADER_2 = '#764ba2'      # 头部渐变结束
C_BG = '#f5f6fb'            # 窗口背景
C_WORD = '#2d2d3a'          # 单词颜色
C_PHON = '#8a8a99'          # 注音颜色
C_MEAN_BG = '#eef0f9'       # 释义底色
C_MEAN_FG = '#4a4a58'       # 释义文字
C_TRACK = '#e3e5f0'         # 进度条轨道
C_BAR = '#667eea'           # 进度条填充
C_NAV_BG = '#e9ebf4'        # 上一个/下一个按钮
C_NAV_FG = '#4a4a5a'
C_VIEW_BG = '#4f6ef7'       # 查看
C_KNOW_BG = '#2fbf8f'       # 认识
C_UNK_BG = '#e4565b'        # 不认识
C_BADGE = '#f59e0b'         # ★ 已在生词本


def blend(c1, c2, t):
    """两个 #RRGGBB 颜色按 t∈[0,1] 线性插值。"""
    r1, g1, b1 = (int(c1[i:i + 2], 16) for i in (1, 3, 5))
    r2, g2, b2 = (int(c2[i:i + 2], 16) for i in (1, 3, 5))
    return '#%02x%02x%02x' % (round(r1 + (r2 - r1) * t),
                              round(g1 + (g2 - g1) * t),
                              round(b1 + (b2 - b1) * t))


def round_rect(cv, x1, y1, x2, y2, r, **kw):
    points = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
              x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(points, smooth=True, **kw)


# ---------------------------------------------------------------- 主界面
class App:
    def __init__(self, root):
        self.root = root
        self.current = min(max(load_progress(), 0), WORD_COUNT - 1)
        self.recorded = False
        self.meaning_shown = False
        self._toast_after = None

        root.title('考研英语背单词')
        root.withdraw()          # 隐藏主窗口：同步完成前不可见，也就无法抢先学习
        W, H = 560, 660
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry('%dx%d+%d+%d' % (W, H, (sw - W) // 2, max(20, (sh - H) // 3)))
        root.minsize(W, H)
        root.configure(bg=C_BG)

        self._build_header(W)
        self._build_body(W)
        self._build_syncbar(W)
        self._refresh_sync_mode()
        root.protocol('WM_DELETE_WINDOW', self._on_close)
        self._closing_with_sync = False
        self.root.after(300, self._auto_sync_at_open)

        # 快捷键
        root.bind('1', lambda e: self.toggle_meaning())
        root.bind('2', lambda e: self.mark_know())
        root.bind('3', lambda e: self.mark_unknown())
        root.bind('<Left>', lambda e: self.go_prev())
        root.bind('<Right>', lambda e: self.go_next())

        self.show_word(self.current)

    # -------------------------------------------------- 头部渐变
    def _build_header(self, W):
        self.header_cv = tk.Canvas(self.root, height=66, highlightthickness=0, bd=0)
        self.header_cv.pack(fill='x')
        for i in range(W):
            self.header_cv.create_line(i, 0, i, 66, fill=blend(C_HEADER_1, C_HEADER_2, i / max(1, W - 1)))
        self.header_cv.create_text(26, 33, anchor='w', text='考研英语背单词', fill='#ffffff',
                                   font=(FONT, 17, 'bold'))
        self.unknown_total_var = tk.StringVar()
        self.header_unknown = self.header_cv.create_text(W - 26, 33, anchor='e', fill='#f0eefe',
                                                         font=(FONT, 11), text='')
        self._update_header_unknown()

    def _update_header_unknown(self):
        self.unknown_total_var.set('生词本 %d 词' % len(unknown_rows))
        try:
            self.header_cv.itemconfigure(self.header_unknown, text=self.unknown_total_var.get())
        except Exception:
            pass

    # -------------------------------------------------- 正文
    def _build_body(self, W):
        body = tk.Frame(self.root, bg=C_BG)
        body.pack(fill='both', expand=True, padx=34, pady=(16, 20))

        # 进度行 + 进度条
        prog_row = tk.Frame(body, bg=C_BG)
        prog_row.pack(fill='x')
        self.pos_var = tk.StringVar()
        tk.Label(prog_row, textvariable=self.pos_var, bg=C_BG, fg='#9a9aab',
                 font=(FONT, 10)).pack(side='left')
        self.badge = tk.Label(prog_row, text='★ 已在生词本', bg=C_BADGE, fg='#fff',
                              font=(FONT, 9), padx=8, pady=2)
        self.badge.pack(side='right')
        self.badge.pack_forget()

        self.prog_cv = tk.Canvas(body, height=10, highlightthickness=0, bd=0, bg=C_BG)
        self.prog_cv.pack(fill='x', pady=(8, 4))
        self.prog_cv.bind('<Configure>', lambda e: self._redraw_progress())

        # 单词
        self.word_var = tk.StringVar()
        self.word_lbl = tk.Label(body, textvariable=self.word_var, bg=C_BG, fg=C_WORD,
                                 font=(FONT, 42, 'bold'), wraplength=W - 68)
        self.word_lbl.pack(pady=(18, 0))

        self.phon_var = tk.StringVar()
        tk.Label(body, textvariable=self.phon_var, bg=C_BG, fg=C_PHON,
                 font=(FONT, 16)).pack(pady=(4, 14))

        # 释义（初始隐藏）
        self.meaning_var = tk.StringVar()
        self.meaning_lbl = tk.Label(body, textvariable=self.meaning_var, bg=C_MEAN_BG, fg=C_MEAN_FG,
                                    font=(FONT, 14), wraplength=W - 68, justify='left',
                                    padx=18, pady=14, anchor='w')
        self.meaning_lbl.pack(fill='x')

        # 按钮
        btns = tk.Frame(body, bg=C_BG)
        btns.pack(pady=(22, 6))
        self._btn_prev = self._mk_button(btns, '← 上一个', C_NAV_BG, C_NAV_FG, self.go_prev)
        self._btn_view = self._mk_button(btns, '查看 (1)', C_VIEW_BG, '#fff', self.toggle_meaning)
        self._btn_know = self._mk_button(btns, '认识 (2)', C_KNOW_BG, '#fff', self.mark_know)
        self._btn_unk = self._mk_button(btns, '不认识 (3)', C_UNK_BG, '#fff', self.mark_unknown)
        self._btn_next = self._mk_button(btns, '下一个 →', C_NAV_BG, C_NAV_FG, self.go_next)

        # 状态提示
        self.status_var = tk.StringVar()
        tk.Label(body, textvariable=self.status_var, bg=C_BG, fg='#8a8a99',
                 font=(FONT, 10)).pack(pady=(4, 2))

        # 快捷键提示
        tk.Label(body, text='快捷键：1 查看　2 认识　3 不认识　← → 翻页', bg=C_BG, fg='#b8b8c6',
                 font=(FONT, 9)).pack(pady=(4, 0))

    def _mk_button(self, parent, text, bg, fg, cmd):
        b = tk.Button(parent, text=text, bg=bg, fg=fg, relief='flat', bd=0,
                      font=(FONT, 11, 'bold'), cursor='hand2', activebackground=bg,
                      activeforeground=fg, padx=16, pady=9, command=cmd)
        b.pack(side='left', padx=5)
        return b

    # -------------------------------------------------- 显示与交互
    def show_word(self, pos):
        pos = max(0, min(pos, WORD_COUNT - 1))
        w = WORDS[pos]
        self.current = pos
        self.word_var.set(w['word'])
        self.phon_var.set(w['phonetic'])
        self.meaning_var.set(w['meaning'])
        self.recorded = w['index'] in added_indices
        self.set_meaning_shown(False)
        self.pos_var.set('第 %d / %d 个单词' % (pos + 1, WORD_COUNT))
        if self.recorded:
            self.badge.pack(side='right')
        else:
            self.badge.pack_forget()
        self._redraw_progress()
        self._btn_prev.configure(state='normal' if pos > 0 else 'disabled')
        self._btn_next.configure(state='normal' if pos < WORD_COUNT - 1 else 'disabled')
        save_progress(pos)
        self._refresh_last_sync()   # 每次翻词/标记后刷新「未上传」红字警示

    def _redraw_progress(self):
        """按当前画布实际宽度重画进度条轨道和填充。"""
        cv = self.prog_cv
        w = cv.winfo_width()
        if w <= 1:
            w = 470
        cv.delete('all')
        round_rect(cv, 0, 1, w, 9, 4, fill=C_TRACK, outline='')
        x2 = max(w * ((self.current + 1) / float(WORD_COUNT)), 8)
        round_rect(cv, 0, 1, x2, 9, 4, fill=C_BAR, outline='')

    def set_meaning_shown(self, shown):
        self.meaning_shown = shown
        if shown:
            self.meaning_lbl.pack(fill='x')
            self._btn_view.configure(text='收起释义')
        else:
            self.meaning_lbl.pack_forget()
            self._btn_view.configure(text='查看 (1)')

    def toggle_meaning(self):
        self.set_meaning_shown(not self.meaning_shown)

    def go_next(self):
        if self.current < WORD_COUNT - 1:
            self.show_word(self.current + 1)

    def go_prev(self):
        if self.current > 0:
            self.show_word(self.current - 1)

    def set_status(self, text):
        self.status_var.set(text)
        if self._toast_after is not None:
            self.root.after_cancel(self._toast_after)
        self._toast_after = self.root.after(1800, lambda: self.status_var.set(''))

    def mark_unknown(self):
        r = add_unknown(self.current)
        if not r['ok']:
            self.set_status(r['error'])
            return
        if r['already']:
            self.set_status('该词已在生词本中')
        else:
            self.set_status('已记入生词本')
        self.recorded = True
        self._update_header_unknown()
        self.badge.pack(side='right')
        self.go_next()

    def mark_know(self):
        r = mark_word_know(self.current)
        if not r['ok']:
            self.set_status(r['error'])
            return
        if r['removed']:
            self.recorded = False
            self.badge.pack_forget()
            self.set_status('已从生词本删除')
        else:
            self.set_status('已标记认识')
        self._update_header_unknown()
        self.go_next()

    # -------------------------------------------------- 同步栏
    def _build_syncbar(self, W):
        bar = tk.Frame(self.root, bg='#ffffff')
        bar.pack(fill='x', side='bottom')
        tk.Frame(bar, bg=C_TRACK, height=1).pack(fill='x')
        inner = tk.Frame(bar, bg='#ffffff')
        inner.pack(fill='x', padx=24, pady=(10, 12))
        self.sync_addr_var = tk.StringVar()
        tk.Label(inner, textvariable=self.sync_addr_var, bg='#ffffff', fg='#4a4a58',
                 font=(FONT, 10), anchor='w', justify='left', wraplength=W - 48).pack(fill='x')
        tk.Label(inner, text='注意：手机端与电脑端请勿同时操作，等一端的同步完成再操作另一端',
                 bg='#fff7e0', fg='#b8860b', font=(FONT, 9), anchor='w', justify='left',
                 wraplength=W - 48).pack(fill='x', pady=(8, 0))
        # 显眼的同步按钮 + 上次同步时间（主页面手动同步入口）
        sync_row = tk.Frame(inner, bg='#ffffff')
        sync_row.pack(fill='x', pady=(10, 0))
        self.sync_btn = tk.Button(sync_row, text='☁ 上传', command=self.sync_push_now,
                                  bg=C_KNOW_BG, fg='#ffffff', relief='flat', bd=0,
                                  font=(FONT, 12, 'bold'), cursor='hand2',
                                  activebackground='#27a97c', activeforeground='#ffffff',
                                  padx=18, pady=8)
        self.sync_btn.pack(side='left')
        self.fetch_btn = tk.Button(sync_row, text='☁ 从云端获取', command=self.fetch_now,
                                   bg=C_VIEW_BG, fg='#ffffff', relief='flat', bd=0,
                                   font=(FONT, 12, 'bold'), cursor='hand2',
                                   activebackground='#3f58d6', activeforeground='#ffffff',
                                   padx=14, pady=8)
        self.fetch_btn.pack(side='left', padx=(10, 0))
        self.last_sync_var = tk.StringVar(value='')
        self.last_sync_lbl = tk.Label(sync_row, textvariable=self.last_sync_var, bg='#ffffff', fg='#4a4a58',
                 font=(FONT, 11, 'bold'), anchor='w')
        self.last_sync_lbl.pack(side='left', fill='x', expand=True, padx=14)
        cloud_row = tk.Frame(inner, bg='#ffffff')
        cloud_row.pack(fill='x', pady=(8, 0))
        self.cloud_status_var = tk.StringVar()
        tk.Label(cloud_row, textvariable=self.cloud_status_var, bg='#ffffff', fg='#2fbf8f',
                 font=(FONT, 9), anchor='w').pack(side='left', fill='x', expand=True)
        tk.Button(cloud_row, text='坚果云同步设置', command=self.cloud_settings, relief='flat', bd=0,
                  bg='#eef0f9', fg='#4a4a58', font=(FONT, 9), cursor='hand2',
                  padx=10, pady=4).pack(side='right', padx=4)
        self._refresh_cloud_status()
        row = tk.Frame(inner, bg='#ffffff')
        row.pack(fill='x', pady=(8, 0))
        self.sync_status_var = tk.StringVar()
        self.sync_status_lbl = tk.Label(row, textvariable=self.sync_status_var, bg='#ffffff',
                                        fg='#2fbf8f', font=(FONT, 9), anchor='w')
        self.sync_status_lbl.pack(side='left', fill='x', expand=True)
        tk.Button(row, text='本地备份', command=self.backup_settings, relief='flat', bd=0,
                  bg=C_NAV_BG, fg=C_NAV_FG, font=(FONT, 9), cursor='hand2',
                  padx=12, pady=5).pack(side='right', padx=4)
        tk.Button(row, text='导出备份', command=self.do_export, relief='flat', bd=0,
                  bg=C_NAV_BG, fg=C_NAV_FG, font=(FONT, 9), cursor='hand2',
                  padx=12, pady=5).pack(side='right', padx=4)
        tk.Button(row, text='导入备份', command=self.do_import, relief='flat', bd=0,
                  bg=C_NAV_BG, fg=C_NAV_FG, font=(FONT, 9), cursor='hand2',
                  padx=12, pady=5).pack(side='right', padx=4)

    def _refresh_sync_mode(self):
        """底部第一行：显示当前同步模式、上次同步时间与本地备份条数。"""
        n = len(list_local_backups())
        ls = last_sync_str()
        if cloud_config():
            self.sync_addr_var.set('坚果云网盘同步已开启：点「☁ 上传」上传、点「☁ 从云端获取」拉取；上次同步：%s（本地备份已保留 %d 条）'
                                   % (ls or '尚未同步', n))
        else:
            self.sync_addr_var.set('坚果云同步未配置：仅本地学习，操作时自动本地备份（已保留 %d 条）。需要两端同步请在下方配置坚果云' % n)
        self._refresh_last_sync()

    def _refresh_last_sync(self):
        """同步按钮旁：未配置灰色提示；有未上传改动红字警示；否则显示上次同步时间。"""
        ls = last_sync_str()
        if not cloud_config():
            self.last_sync_var.set('未配置坚果云同步（点击下方「坚果云同步设置」）')
            self.last_sync_lbl.configure(fg='#4a4a58')
            self.sync_btn.configure(state='normal')
        elif is_dirty():
            self.last_sync_var.set('⚠ 有改动未上传到云端%s' % ('（上次同步：%s）' % ls if ls else ''))
            self.last_sync_lbl.configure(fg='#e4565b')
        elif ls:
            self.last_sync_var.set('上次同步：%s' % ls)
            self.last_sync_lbl.configure(fg='#4a4a58')
        else:
            self.last_sync_var.set('尚未同步过云端')
            self.last_sync_lbl.configure(fg='#4a4a58')

    # -------------------------------------------------- 坚果云网盘同步设置
    def _refresh_cloud_status(self):
        if cloud_config():
            self.cloud_status_var.set('坚果云网盘同步：已开启（手动同步）')
        else:
            self.cloud_status_var.set('坚果云网盘同步：未配置')

    def cloud_settings(self):
        cfg = cloud_config() or {}
        dlg = tk.Toplevel(self.root)
        dlg.title('坚果云网盘同步设置')
        dlg.configure(bg='#ffffff')
        dlg.transient(self.root)
        dlg.grab_set()
        tk.Label(dlg, text='坚果云网盘同步', bg='#ffffff', fg=C_WORD,
                 font=(FONT, 14, 'bold')).pack(pady=(18, 6))
        tk.Label(dlg, text=('坚果云是唯一的同步方式，全手动：\n'
                            '· 打开程序时弹窗询问「是否从云端获取最新数据」；\n'
                            '· 主界面底部点「☁ 上传」上传、点「☁ 从云端获取」拉取；\n'
                            '· 关闭时若有未上传的改动，会弹窗确认是否上传再退出。\n'
                            '需要免费注册坚果云，在网页端生成「应用密码」，并先在坚果云里建一个文件夹\n'
                            '（如 kaoyan_backup），下面地址填到该文件夹里 kaoyan.json 的完整路径。\n'
                            '电脑端与手机端必须填同一个地址。'),
                 bg='#ffffff', fg='#6a6a78', font=(FONT, 9), justify='left', anchor='w',
                 padx=24, wraplength=430).pack(fill='x', pady=(0, 10))
        frm = tk.Frame(dlg, bg='#ffffff')
        frm.pack(padx=24, fill='x')
        tk.Label(frm, text='WebDAV 地址', bg='#ffffff', fg='#4a4a58', font=(FONT, 9), anchor='w').pack(fill='x')
        url_ev = tk.Entry(frm, font=(FONT, 10))
        url_ev.pack(fill='x', pady=(2, 6))
        tk.Label(frm, text='用户名（坚果云账号）', bg='#ffffff', fg='#4a4a58', font=(FONT, 9), anchor='w').pack(fill='x')
        user_ev = tk.Entry(frm, font=(FONT, 10))
        user_ev.pack(fill='x', pady=(2, 6))
        tk.Label(frm, text='应用密码（坚果云网页生成，非登录密码）', bg='#ffffff', fg='#4a4a58',
                 font=(FONT, 9), anchor='w').pack(fill='x')
        pass_ev = tk.Entry(frm, font=(FONT, 10), show='*')
        pass_ev.pack(fill='x', pady=(2, 6))
        url_ev.insert(0, cfg.get('url', ''))
        user_ev.insert(0, cfg.get('user', ''))
        pass_ev.insert(0, cfg.get('password', ''))
        status = tk.StringVar()
        tk.Label(dlg, textvariable=status, bg='#ffffff', fg='#e4565b', font=(FONT, 9)).pack(pady=(4, 6))

        def save():
            save_cloud_config(url_ev.get(), user_ev.get(), pass_ev.get())
            self._refresh_cloud_status()
            self._refresh_sync_mode()
            status.set('已保存：点「☁ 上传」上传、点「☁ 从云端获取」拉取，打开时会询问是否获取')
            dlg.after(900, dlg.destroy)

        def test_conn():
            def run():
                data, err = cloud_fetch()
                if err:
                    self.root.after(0, lambda: status.set('连接失败：%s' % err))
                elif data is None:
                    self.root.after(0, lambda: status.set('连接成功：云端还没有数据文件'))
                else:
                    self.root.after(0, lambda: status.set(
                        '连接成功：云端进度 %s，生词 %d' % (data.get('progress'), len(data.get('unknown') or []))))
            threading.Thread(target=run, daemon=True).start()

        def clear():
            clear_cloud_config()
            self._refresh_cloud_status()
            self._refresh_sync_mode()
            status.set('已清除配置')
            dlg.after(900, dlg.destroy)

        btns = tk.Frame(dlg, bg='#ffffff')
        btns.pack(pady=(6, 16))
        tk.Button(btns, text='保存', command=save, bg=C_VIEW_BG, fg='#fff', relief='flat',
                  font=(FONT, 10, 'bold'), padx=18, pady=6).pack(side='left', padx=4)
        tk.Button(btns, text='测试连接', command=test_conn, bg=C_NAV_BG, fg=C_NAV_FG, relief='flat',
                  font=(FONT, 10), padx=12, pady=6).pack(side='left', padx=4)
        tk.Button(btns, text='清除配置', command=clear, bg=C_UNK_BG, fg='#fff', relief='flat',
                  font=(FONT, 10), padx=12, pady=6).pack(side='left', padx=4)

    # -------------------------------------------------- 打开/关闭时触发同步
    def _make_sync_dialog(self, title, heading, instructions, btn_text, btn_cmd, btn_bg, btn_fg,
                          transient=True, default_status='正在从云端获取数据…'):
        """创建模态同步提示弹窗（阻塞主界面操作，直到同步完成或用户点按钮/关窗）。
           按内容自适应大小居中；状态栏保持单行，确保按钮始终可见可点；
           × 关闭按钮行为与唯一按钮一致（防内容变化时无法退出）。
           transient=False：主窗口已隐藏（打开同步阶段）时，弹窗独立显示并获焦，
           不挂到主窗口上（挂到隐藏主窗口上会连带被隐藏）。"""
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.configure(bg='#ffffff')
        if transient:
            dlg.transient(self.root)        # 主窗口可见：弹窗保持在主窗口之上
        else:
            dlg.lift()                      # 主窗口隐藏：弹窗独立显示
            dlg.focus_force()
        dlg.grab_set()                      # 模态：先完成同步才能操作主界面
        tk.Label(dlg, text=heading, bg='#ffffff', fg=C_WORD,
                 font=(FONT, 15, 'bold')).pack(pady=(24, 10))
        tk.Label(dlg, text=instructions, bg='#ffffff', fg='#4a4a58', font=(FONT, 11),
                 justify='left', anchor='w', padx=32).pack(fill='x')
        dlg._status = tk.StringVar(value=default_status)
        tk.Label(dlg, textvariable=dlg._status, bg='#ffffff', fg='#e4565b',
                 font=(FONT, 10)).pack(pady=(14, 6))
        tk.Button(dlg, text=btn_text, command=btn_cmd, bg=btn_bg, fg=btn_fg,
                  relief='flat', font=(FONT, 11, 'bold'), padx=18, pady=8).pack(pady=(4, 18))
        dlg.protocol('WM_DELETE_WINDOW', btn_cmd)   # × 行为与唯一按钮一致
        dlg.update_idletasks()
        w = max(dlg.winfo_reqwidth(), 340)
        h = dlg.winfo_reqheight()
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry('%dx%d+%d+%d' % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 2)))
        dlg.resizable(False, False)
        return dlg

    def _auto_sync_at_open(self):
        """打开电脑端：先显示主界面；已配置坚果云则弹窗询问是否从云端获取（手动，不自动）。"""
        self._open_cancel = False
        try:
            self.root.deiconify()
        except Exception:
            pass
        if not cloud_config():
            self.set_sync_status('坚果云同步未配置：仅本地学习，操作时自动本地备份', ok=True)
            return
        # 打开时提示：是否从云端获取最新数据（可跳过，稍后可点「☁ 上传」）
        try:
            if messagebox.askyesno(
                    '坚果云同步',
                    '是否从云端获取最新数据？\n\n'
                    '获取会用云端最新进度和生词覆盖本机；\n'
                    '本机尚未上传的改动会被云端覆盖。\n\n'
                    '（也可选「否」，之后点主界面底部「☁ 上传」上传本机数据）',
                    parent=self.root):
                self._start_cloud_open_sync()
            else:
                self.set_sync_status('本次未从云端获取，使用本机数据（可随时点「☁ 上传」）', ok=True)
        except Exception as e:
            log('打开询问失败：%s' % e)

    def _start_cloud_open_sync(self):
        """打开电脑端（用户选择从云端获取）：拉取云端数据，取完应用并刷新。"""
        try:
            dlg = self._make_sync_dialog(
                '同步提示', '正在从云端获取数据…',
                '正在拉取坚果云上的最新进度和生词。\n\n'
                '获取完成后自动应用云端数据；\n'
                '若云端无法连接，可点「关闭」留在本机数据。',
                '关闭', self._skip_cloud_open, C_NAV_BG, C_NAV_FG,
                default_status='正在从云端获取数据…')
            self._open_dlg = dlg
            threading.Thread(target=lambda: self._wait_cloud_open(dlg), daemon=True).start()
            self.root.wait_window(dlg)
        except Exception as e:
            log('云端获取弹窗失败：%s' % e)

    def _skip_cloud_open(self):
        """获取云端时用户关闭弹窗：留在本机数据（本次未获取）。"""
        self._open_cancel = True
        try:
            self._open_dlg.destroy()
            self.set_sync_status('本次未从云端获取，使用本机数据', ok=False)
        except Exception:
            pass

    def _wait_cloud_open(self, dlg):
        data, err = cloud_fetch()
        if self._open_cancel:      # 用户已关闭弹窗，不再覆盖窗口状态
            return
        if err:
            self.root.after(0, lambda: self._set_dlg_status(dlg, '云端获取失败：%s' % err))
            log('打开云端获取失败：%s' % err)
            return
        # 云端还没有数据：保持本地，无需二次确认
        if cloud_state_desc(data) == '暂无数据':
            self.root.after(0, lambda: self._finish_cloud_open(dlg, '云端暂无数据，保持本地进度和生词'))
            return
        # 取到云端数据 → 弹窗显示云端与本机数字，让用户再次确认是否覆盖本机
        self.root.after(0, lambda: self._ask_cloud_apply(dlg, data))

    def _ask_cloud_apply(self, dlg, data):
        """取到云端数据后二次确认：显示云端与本机数字，确认才应用。"""
        try:
            dlg.destroy()          # 关闭「正在获取」弹窗
        except Exception:
            pass
        ok = messagebox.askyesno(
            '获取确认',
            '云端当前：%s\n本机当前：%s\n\n确认用云端数据覆盖本机吗？' % (
                cloud_state_desc(data), local_state_desc()),
            parent=self.root)
        if not ok:
            self._open_cancel = True
            self.set_sync_status('已取消获取，使用本机数据', ok=False)
            return
        applied, msg = apply_cloud_data(data)
        if applied:
            mark_synced()          # 应用了云端数据 → 更新基线与上次同步时间
        self._finish_cloud_open(dlg, msg)

    def _finish_cloud_open(self, dlg, msg):
        try:
            dlg.destroy()
            self.current = min(max(load_progress(), 0), WORD_COUNT - 1)
            self.show_word(self.current)
            self._update_header_unknown()
            self._refresh_sync_mode()
            self.set_sync_status(msg, ok=True)
            log('打开云端获取：%s' % msg)
        except Exception:
            pass

    def _on_close(self):
        """关闭窗口：先做本地备份；已配置坚果云且本机有未上传改动 → 弹窗三选
           （上传并退出 / 直接退出 / 继续学习）；无改动或未配置则直接退出。"""
        if self._closing_with_sync:
            return
        self._closing_with_sync = True
        self._close_cancel = False
        # 先本地备份（秒级、同步完成），云端是否可用都不影响备份落盘
        name, err = make_local_backup()
        if err:
            log(err)
            self.set_sync_status(err, ok=False)
        else:
            log('关闭：已生成本地备份 %s（目录 %s）' % (name, BACKUP_DIR))
            self._refresh_sync_mode()
        if not cloud_config() or not is_dirty():
            self.root.destroy()
            return
        # 有未上传到云端的改动 → 三选一
        ans = messagebox.askyesnocancel(
            '退出确认',
            '本机还有改动未上传到云端（上次同步：%s）。\n\n'
            '「是」：先上传到云端，再退出；\n'
            '「否」：不上传直接退出（本次改动只保存在本地备份里）；\n'
            '「取消」：留在程序，继续学习。' % (last_sync_str() or '从未同步'),
            parent=self.root)
        if ans is None:      # 取消 → 继续学习
            self._closing_with_sync = False
            return
        if not ans:          # 否 → 直接退出
            self.root.destroy()
            return
        # 是 → 上传并退出：先读云端数据二次确认（显示云端与本机数字）
        def start_close_push():
            self._start_cloud_close_sync()

        def cancel_close_push():
            self._closing_with_sync = False      # 允许下次再点关闭重试
            self.set_sync_status('已取消上传，留在本程序', ok=False)

        self._confirm_upload(start_close_push, cancel_close_push)

    def _cancel_close(self, dlg):
        """关闭时未同步成功：唯一选择是关闭提示窗口，继续学习。"""
        self._close_cancel = True
        self._closing_with_sync = False      # 允许下次再点关闭重试
        try:
            dlg.destroy()
        except Exception:
            pass

    def _start_cloud_close_sync(self):
        """关闭电脑端（已选「先上传再退出」）：把本端数据上传到云端，传完退出程序。"""
        dlg = self._make_sync_dialog(
            '同步提示', '正在上传数据到云端…',
            '已选择在退出前上传本机进度和生词到云端（本机已生成本地备份）。\n\n'
            '上传完成前无法关闭程序，完成后程序自动关闭。\n'
            '若云端无法连接，可点「继续学习」留在本程序稍后再试。',
            '继续学习', lambda: self._cancel_close(dlg), C_VIEW_BG, '#fff',
            default_status='正在上传到云端…')
        self._close_dlg = dlg
        threading.Thread(target=lambda: self._wait_cloud_close(dlg), daemon=True).start()
        self.root.wait_window(dlg)

    def _wait_cloud_close(self, dlg):
        err = cloud_push()
        if self._close_cancel:      # 用户已点「继续学习」，不再强制退出
            return
        if err:
            self.root.after(0, lambda: self._set_dlg_status(
                dlg, '上传云端失败：%s\n点「继续学习」留在本程序（本次未上传云端）' % err))
            log('关闭上传云端失败：%s' % err)
            return
        mark_synced()               # 上传成功 → 更新基线与上次同步时间
        self.root.after(0, lambda: self._finish_cloud_close(dlg))

    def _finish_cloud_close(self, dlg):
        try:
            dlg._status.set('已上传云端：%s，正在退出…' % local_state_desc())
            self.root.after(400, self.root.destroy)
            log('关闭上传云端成功，程序退出')
        except Exception:
            pass

    def _set_dlg_status(self, dlg, text):
        """线程安全地更新同步弹窗的状态文字（弹窗可能已被销毁，一律兜底）。"""
        try:
            dlg._status.set(text)
        except Exception:
            pass

    def set_sync_status(self, text, ok=True):
        self.sync_status_var.set(text)
        self.sync_status_lbl.configure(fg='#2fbf8f' if ok else '#e4565b')

    # -------------------------------------------------- 主界面「☁ 从云端获取」
    def fetch_now(self):
        """主界面「☁ 从云端获取」：先读云端数据二次确认，确认后应用。"""
        if not cloud_config():
            messagebox.showinfo('坚果云同步', '尚未配置坚果云同步。\n点「确定」打开设置，填写地址/用户名/应用密码后即可从云端获取。')
            self.cloud_settings()
            return
        make_local_backup()   # 获取前先本地备份，覆盖错了也能用备份找回
        self._start_cloud_open_sync()

    # -------------------------------------------------- 上传前二次确认（显示云端数字）
    def _confirm_upload(self, on_confirm, on_cancel=None):
        """上传前：先读云端当前数据，弹「上传确认」显示云端与本机数字；确认后调 on_confirm()。"""
        def work():
            data, err = cloud_fetch()
            self.root.after(0, lambda: self._show_upload_confirm(data, err, on_confirm, on_cancel))
        threading.Thread(target=work, daemon=True).start()

    def _show_upload_confirm(self, data, err, on_confirm, on_cancel):
        if err:
            ok = messagebox.askyesno(
                '上传确认',
                '无法读取云端当前数据：%s\n\n仍要把本机数据上传覆盖云端吗？' % err,
                parent=self.root)
        else:
            ok = messagebox.askyesno(
                '上传确认',
                '云端当前：%s\n本机当前：%s\n\n确认把本机数据上传覆盖云端吗？' % (
                    cloud_state_desc(data), local_state_desc()),
                parent=self.root)
        if ok:
            try:
                on_confirm()
            except Exception as e:
                log('上传确认后执行失败：%s' % e)
        elif on_cancel:
            try:
                on_cancel()
            except Exception:
                pass

    # -------------------------------------------------- 手动同步（主页面按钮）
    def sync_push_now(self):
        """主页面「☁ 上传」：先读云端数据二次确认，确认后把本机数据上传到坚果云（手动，不下载）。"""
        if not cloud_config():
            messagebox.showinfo('坚果云同步', '尚未配置坚果云同步。\n点「确定」打开设置，填写地址/用户名/应用密码后即可上传。')
            self.cloud_settings()
            return
        make_local_backup()   # 上传前先本地备份（不联网），双保险

        def start_push():
            dlg = self._make_sync_dialog(
                '同步提示', '正在上传数据到云端…',
                '正在把本机学习进度和生词上传到坚果云。\n\n上传完成后弹窗自动关闭。',
                '关闭', lambda: self._cancel_sync_push(dlg), C_NAV_BG, C_NAV_FG,
                default_status='正在上传到云端…')
            threading.Thread(target=lambda: self._wait_sync_push(dlg), daemon=True).start()

        def cancel_push():
            self.set_sync_status('已取消上传（未上传云端）', ok=False)

        self._confirm_upload(start_push, cancel_push)

    def _cancel_sync_push(self, dlg):
        try:
            dlg.destroy()
        except Exception:
            pass

    def _wait_sync_push(self, dlg):
        err = cloud_push()
        if err:
            self.root.after(0, lambda: self._set_dlg_status(dlg, err + '\n本次未上传，可稍后重试'))
            log('手动上传云端失败：%s' % err)
            return
        mark_synced()
        self.root.after(0, lambda: self._finish_sync_push(dlg))

    def _finish_sync_push(self, dlg):
        try:
            dlg.destroy()
        except Exception:
            pass
        self._refresh_sync_mode()
        self.set_sync_status('已上传云端：%s（%s）' % (local_state_desc(), last_sync_str() or ''), ok=True)

    def do_export(self):
        path = filedialog.asksaveasfilename(
            title='导出备份', defaultextension='.json', initialfile='考研英语背单词备份.json',
            filetypes=[('JSON 备份', '*.json'), ('所有文件', '*.*')])
        if not path:
            return
        err = export_backup_file(path)
        if err:
            messagebox.showerror('导出备份', err)
        else:
            self.set_status('备份已导出')
            messagebox.showinfo('导出备份', '备份已导出到：\n%s' % path)

    def do_import(self):
        path = filedialog.askopenfilename(
            title='导入备份', filetypes=[('JSON 备份', '*.json'), ('所有文件', '*.*')])
        if not path:
            return
        if not messagebox.askyesno('导入备份', '导入将覆盖当前的学习进度和生词本，确定继续？'):
            return
        err = import_backup_file(path)
        if err:
            messagebox.showerror('导入备份', err)
            return
        self._refresh_after_import()
        messagebox.showinfo('导入备份', '导入完成')

    def _refresh_after_import(self):
        self.current = min(max(load_progress(), 0), WORD_COUNT - 1)
        self.show_word(self.current)
        self._update_header_unknown()

    # -------------------------------------------------- 本地备份
    def backup_settings(self):
        """本地备份管理弹窗：列出最近 10 条，可恢复选中项或打开备份文件夹。"""
        dlg = tk.Toplevel(self.root)
        dlg.title('本地备份')
        dlg.configure(bg='#ffffff')
        dlg.transient(self.root)
        dlg.grab_set()
        tk.Label(dlg, text='本地备份（保留最近 10 条）', bg='#ffffff', fg=C_WORD,
                 font=(FONT, 14, 'bold')).pack(pady=(18, 6))
        tk.Label(dlg, text=('每次关闭程序时自动把进度和生词本存到 exe 旁的「本地备份」文件夹，\n'
                            '云端出错时可用最近备份恢复（与云端互为双保险）。'),
                 bg='#ffffff', fg='#6a6a78', font=(FONT, 9), justify='left', anchor='w',
                 padx=24, wraplength=430).pack(fill='x', pady=(0, 10))
        frm = tk.Frame(dlg, bg='#ffffff')
        frm.pack(fill='both', expand=True, padx=24)
        lb = tk.Listbox(frm, font=(FONT, 9), height=10, activestyle='none',
                        selectbackground=C_BAR, selectforeground='#fff', exportselection=False)
        lb.pack(side='left', fill='both', expand=True)
        sb = tk.Scrollbar(frm, orient='vertical', command=lb.yview)
        sb.pack(side='right', fill='y')
        lb.config(yscrollcommand=sb.set)
        backups = list_local_backups()
        for b in backups:
            lb.insert('end', '%s   进度 %d   生词 %d 个' % (b['time'], b['progress'], b['unknown_total']))
        if not backups:
            lb.insert('end', '（还没有本地备份，关闭程序时自动生成）')

        def restore():
            sel = lb.curselection()
            if not sel:
                return
            b = backups[sel[0]]
            if not messagebox.askyesno('恢复本地备份', '恢复「%s」的备份将覆盖当前进度和生词本，确定继续？' % b['time']):
                return
            err = import_backup_file(b['path'])
            if err:
                messagebox.showerror('恢复本地备份', err)
                return
            self._refresh_after_import()
            self._refresh_sync_mode()
            messagebox.showinfo('恢复本地备份', '已恢复到 %s 的备份' % b['time'])
            dlg.destroy()

        def open_folder():
            try:
                if not os.path.isdir(BACKUP_DIR):
                    os.makedirs(BACKUP_DIR, exist_ok=True)
                os.startfile(BACKUP_DIR)
            except Exception as e:
                messagebox.showerror('打开文件夹', str(e))

        btns = tk.Frame(dlg, bg='#ffffff')
        btns.pack(pady=(6, 16))
        tk.Button(btns, text='恢复选中备份', command=restore, bg=C_VIEW_BG, fg='#fff',
                  relief='flat', font=(FONT, 10, 'bold'), padx=14, pady=6).pack(side='left', padx=4)
        tk.Button(btns, text='打开备份文件夹', command=open_folder, bg=C_NAV_BG, fg=C_NAV_FG,
                  relief='flat', font=(FONT, 10), padx=12, pady=6).pack(side='left', padx=4)
        tk.Button(btns, text='关闭', command=dlg.destroy, bg=C_UNK_BG, fg='#fff',
                  relief='flat', font=(FONT, 10), padx=12, pady=6).pack(side='left', padx=4)
        dlg.update_idletasks()
        w = max(dlg.winfo_reqwidth(), 420)
        h = dlg.winfo_reqheight()
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry('%dx%d+%d+%d' % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 2)))
        dlg.resizable(False, False)


def main():
    root = tk.Tk()
    App(root)
    log('启动：%s，共 %d 词，生词本 %d 个，坚果云同步=%s'
        % (SOURCE_FILE, WORD_COUNT, len(unknown_rows), '开' if cloud_config() else '关'))
    root.mainloop()


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        log('启动失败：%s\n%s' % (e, traceback.format_exc()))
        try:
            messagebox.showerror('考研英语背单词', '程序启动失败：\n%s' % e)
        except Exception:
            pass
        sys.exit(1)
