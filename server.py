# -*- coding: utf-8 -*-
"""区域工厂液体装货趋势 - 本地轻量服务

启动:  python server.py   (默认端口 8787, 可用环境变量 PORT 覆盖)
依赖:  flask / pandas / openpyxl / xlrd
"""
import json
import math
import os
import re
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
from flask import Flask, Response, request, send_from_directory

BASE = Path(__file__).resolve().parent
UPLOAD_DIR = BASE / "data" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
MAPPING_DIR = BASE / "data" / "mappings"
MAPPING_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, static_folder="static", static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64MB

_SHEETS = {}  # name -> {sheet_name: DataFrame}

FREQ = {"day": ("D", "%m-%d"), "week": ("W", "%m-%d"), "month": ("MS", "%Y-%m")}


# ---------- 工具 ----------

def _default(o):
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, np.floating):
        v = float(o)
        return v if math.isfinite(v) else None
    if isinstance(o, np.bool_):
        return bool(o)
    if isinstance(o, (pd.Timestamp, datetime, date)):
        return o.isoformat()
    return str(o)


def _clean(obj):
    """递归把 NaN/Inf 转成 None, 保证输出是合法 JSON"""
    if isinstance(obj, float):
        return obj if math.isfinite(obj) else None
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_clean(v) for v in obj]
    return obj


def ok(data=None):
    return Response(
        json.dumps({"ok": True, "data": _clean(data)}, ensure_ascii=False, default=_default),
        mimetype="application/json; charset=utf-8")


def err(msg, code=400):
    return Response(
        json.dumps({"ok": False, "msg": msg}, ensure_ascii=False),
        mimetype="application/json; charset=utf-8", status=code)


def load_workbook(name):
    if name not in _SHEETS:
        path = UPLOAD_DIR / name
        if not path.exists():
            raise FileNotFoundError("文件不存在: " + name)
        raw = pd.read_excel(path, sheet_name=None)
        book = {}
        for k, v in raw.items():
            v.columns = [str(c) for c in v.columns]
            book[str(k)] = v
        _SHEETS[name] = book
    return _SHEETS[name]


def pick_sheet(name, sheet):
    sheets = load_workbook(name)
    if sheet is None or sheet not in sheets:
        sheet = next(iter(sheets))
    return sheet, sheets[sheet]


# ---------- 只保留最新一份 ----------
# 装货报表: 每次上传成功后, uploads 里只留这一份
# 匹配表:   mappings 里只留最新写入的一份(内容已独立存为 JSON, 与原始 xlsx 无关)

def prune_uploads(keep_name):
    """删除 uploads 中除 keep_name 外的所有文件, 返回被删除的文件名列表"""
    removed = []
    for p in UPLOAD_DIR.glob("*"):
        if p.name.startswith(".") or p.name == keep_name:
            continue
        try:
            if p.is_file():
                p.unlink()
                removed.append(p.name)
                _SHEETS.pop(p.name, None)
        except OSError:
            pass
    return removed


def prune_mappings(keep_path=None):
    """匹配表只保留最新的一份(按修改时间); keep_path 指定时优先保留它"""
    files = sorted(MAPPING_DIR.glob("*.json"),
                   key=lambda x: x.stat().st_mtime, reverse=True)
    if keep_path is not None:
        keep_path = Path(keep_path)
        files = [p for p in files if p != keep_path]
        files = ([keep_path] if keep_path.exists() else []) + files
    removed = []
    for p in files[1:]:
        try:
            p.unlink()
            removed.append(p.name)
        except OSError:
            pass
    return removed


# ---------- 匹配表 ----------

_MAPPING_KEY_RE = re.compile(r'[\\/:*?"<>|\s]')

# 缺省匹配设置: 打开页面自动应用; 只有上传新的匹配表时才会被改写
DEFAULT_MAPPING_PATH = BASE / "data" / "default_mapping.json"


def load_default_mapping():
    """读取缺省匹配设置 {name, join, mkey}; 无或匹配表已不存在则返回 None"""
    if not DEFAULT_MAPPING_PATH.exists():
        return None
    try:
        d = json.loads(DEFAULT_MAPPING_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(d, dict) or not d.get("name"):
        return None
    name = d["name"]
    if not _mapping_path(name).exists():
        return None
    return {"name": name, "join": d.get("join", "") or "", "mkey": d.get("mkey", "") or ""}


def save_default_mapping(name, join="", mkey=""):
    DEFAULT_MAPPING_PATH.parent.mkdir(parents=True, exist_ok=True)
    DEFAULT_MAPPING_PATH.write_text(json.dumps(
        {"name": name, "join": join or "", "mkey": mkey or "",
         "updated": datetime.now().strftime("%Y-%m-%d %H:%M")},
        ensure_ascii=False, indent=2), encoding="utf-8")


def latest_upload_name():
    """uploads 里最新的一份数据文件(现在只会有一份)"""
    files = sorted(UPLOAD_DIR.glob("*"), key=lambda x: x.stat().st_mtime, reverse=True)
    for p in files:
        if p.is_file() and not p.name.startswith("."):
            return p.name
    return ""


def suggest_join_candidates(m, mkey, fname, sheet=None):
    """对 fname 的每一列计算与匹配表关键列的命中率, 按 (命中率, 命中行数) 降序"""
    keys = {_norm_key(v) for v in pd.DataFrame(m["data"])[mkey]}
    keys.discard("")
    if not keys:
        return []
    sheet, df = pick_sheet(fname, sheet)
    scored = []
    n = len(df)
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_bool_dtype(s):
            continue
        vals = s.dropna()
        if not len(vals):
            continue
        norm = vals.map(_norm_key)
        hit = int(norm.isin(keys).sum())
        if hit == 0:
            continue
        scored.append({"column": c, "hit": hit, "total": n,
                       "rate": round(hit / len(norm) * 100, 1)})
    scored.sort(key=lambda x: (-x["rate"], -x["hit"]))
    return scored


def _safe_name(name):
    safe = _MAPPING_KEY_RE.sub('_', name).strip('._')
    if not safe:
        raise ValueError("名称不合法")
    return safe


def _mapping_path(name):
    return MAPPING_DIR / (_safe_name(name) + ".json")


def load_mapping(name):
    p = _mapping_path(name)
    if not p.exists():
        raise FileNotFoundError(name)
    return json.loads(p.read_text(encoding="utf-8"))


def save_mapping_obj(m):
    p = _mapping_path(m["name"])
    p.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")


def delete_mapping(name):
    p = _mapping_path(name)
    if p.exists():
        p.unlink()


def _norm_key(x):
    """把关键值归一化成可比较的字符串: 去空白/转字符串/剥 .0 后缀/去前导0."""
    if x is None:
        return ""
    if isinstance(x, float):
        if math.isnan(x) or math.isinf(x):
            return ""
        if x.is_integer():
            s = str(int(x))
        else:
            s = f"{x:g}"
    elif isinstance(x, (int, np.integer)):
        s = str(int(x))
    else:
        s = str(x).strip()
    return s


def apply_mapping_to_df(df, mapping, join_col, mkey=None):
    """把匹配表的数据合并到 df 上. 返回 (合并后的df, 新增列名).

    join_col: 数据侧(df)用于关联的列名.
    mkey:     匹配表侧用于关联的列名. 默认 = mapping["key"].
    """
    mkey = (mkey or mapping.get("key") or "").strip()
    if not mkey:
        raise ValueError("匹配表缺少关键列(key)")
    if join_col not in df.columns:
        raise ValueError("数据中没有字段: " + join_col)
    if mkey not in mapping["columns"]:
        raise ValueError("匹配表中没有字段: " + mkey +
                         "（可选: " + "、".join(mapping["columns"]) + "）")
    mdf = pd.DataFrame(mapping["data"])
    left = df.copy()
    left[join_col] = left[join_col].map(_norm_key)
    right = mdf.copy()
    right[mkey] = right[mkey].map(_norm_key)
    # 同一 join 键的多行匹配会导致数据翻倍, 去重保留首行
    right = right.drop_duplicates(subset=[mkey], keep="first").copy()
    # 关键列重命名为 join_col, 便于 merge
    right = right.rename(columns={mkey: join_col})
    # 不引入与 df 已存在的列同名的列 (避免覆盖原数据)
    new_cols = [c for c in mapping["columns"] if c != mkey and c not in df.columns]
    # 若 mkey 本身和 df 已有列同名, 改名加后缀避免冲突
    safe_new_cols = []
    rename_map = {}
    for c in new_cols:
        if c in left.columns:
            nc = c + "_匹配"
            rename_map[c] = nc
            safe_new_cols.append(nc)
        else:
            safe_new_cols.append(c)
    if rename_map:
        right = right.rename(columns=rename_map)
    aug = left.merge(right[[join_col] + safe_new_cols], on=join_col, how="left")
    return aug, safe_new_cols


def maybe_augment(df):
    """读取 query string 中的 mapping/join/mkey 并增强 df.
    返回 (df, 新增列, mapping_meta, error_tuple)."""
    mapping_name = request.args.get("mapping", "").strip()
    if not mapping_name:
        return df, [], None, None
    join = request.args.get("join", "").strip()
    mkey = request.args.get("mkey", "").strip()
    try:
        m = load_mapping(mapping_name)
        if not mkey:
            mkey = m["key"]
        aug, added = apply_mapping_to_df(df, m, join, mkey=mkey)
        return aug, added, {"name": mapping_name, "join": join or mkey, "mkey": mkey,
                            "added": added}, None
    except FileNotFoundError:
        return df, [], None, ("匹配表不存在: " + mapping_name, 404)
    except ValueError as e:
        return df, [], None, (str(e), 400)
    except Exception as e:
        return df, [], None, ("应用匹配失败: " + str(e), 500)


# ---------- 产品分类配置 ----------

DEFAULT_CATEGORIES = {
    "液氧(含医用氧)": ["液氧", "医用氧"],
    "液氮(含食品氮)": ["液氮", "食品氮"],
    "液氩": ["液氩"],
}
CONFIG_PATH = BASE / "data" / "config.json"


def load_product_categories():
    if CONFIG_PATH.exists():
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            cats = cfg.get("product_categories")
            if isinstance(cats, dict) and cats:
                return cats
        except Exception:
            pass
    return DEFAULT_CATEGORIES


def save_product_categories(cats):
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps({"product_categories": cats}, ensure_ascii=False, indent=2),
        encoding="utf-8")


def categorize_product(value, categories):
    """把原始产品值归类. 支持:
    - 精确匹配（默认产品列表里的值）
    - 包含匹配（产品字符串中包含关键词，如"液氧;99.5%" 含"液氧" → 液氧(含医用氧)）
    """
    if value is None or (isinstance(value, float) and not math.isfinite(value)):
        return "未分类"
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return "未分类"
    # 优先精确匹配
    for cat, products in categories.items():
        if s in products:
            return cat
    # 其次包含匹配（产品名带";"/"%"也能命中）
    for cat, products in categories.items():
        for kw in products:
            if kw and kw in s:
                return cat
    return "未分类"


# 展示用简称: 分组口径仍用全名, 只有呈现给用户的文字用简称
CATEGORY_SHORT = {
    "液氧(含医用氧)": "液氧",
    "液氮(含食品氮)": "液氮",
    "液氩": "液氩",
}


def cat_short(name):
    """液氧(含医用氧)→液氧, 液氮(含食品氮)→液氮; 未配置的品类取括号前文字"""
    if name in CATEGORY_SHORT:
        return CATEGORY_SHORT[name]
    i = str(name).find("(")
    return str(name)[:i].strip() if i > 0 else str(name)


VOLUME_ALIASES = ["装车量(流通)", "装车量(吨)", "装车量", "流通量", "装车", "数量"]
DATE_ALIASES = ["装货日期", "装车日期", "日期"]
PRODUCT_ALIASES = ["产品", "品名", "装车货物", "货物"]
FACTORY_ALIASES = ["工厂", "厂名", "厂家"]
DELIVERY_ALIASES = ["配送类型", "配送方式", "运输方式", "配送", "方式"]


def autodetect_col(df_cols, aliases):
    # 优先精确匹配（按别名列表顺序），避免选了"装车货物编码"而非"装车货物"
    for a in aliases:
        if a in df_cols:
            return a
    # 退化为模糊匹配：选列名最短的命中（更具体的列名优先）
    hits = []
    for a in aliases:
        for c in df_cols:
            if a in c:
                hits.append(c)
    if hits:
        return sorted(hits, key=lambda x: (len(x), x))[0]
    return ""


# ---------- 页面 ----------

@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


# ---------- API ----------

@app.get("/api/files")
def api_files():
    items = []
    for p in sorted(UPLOAD_DIR.glob("*"), key=lambda x: x.stat().st_mtime, reverse=True):
        if p.name.startswith("."):
            continue
        st = p.stat()
        items.append({
            "name": p.name,
            "size": st.st_size,
            "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
        })
    return ok(items)


@app.post("/api/upload")
def api_upload():
    f = request.files.get("file")
    if f is None or not f.filename:
        return err("未收到文件")
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return err("仅支持 .xlsx / .xls 文件")
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    name = "%s-%s" % (ts, Path(f.filename).name)
    f.save(UPLOAD_DIR / name)
    _SHEETS.pop(name, None)
    try:
        sheets = load_workbook(name)
        first = next(iter(sheets))
        df = sheets[first]
        pruned = prune_uploads(name)   # 只保留最新上传的这一份
        return ok({"name": name, "sheets": list(sheets), "sheet": first,
                   "rows": int(len(df)), "cols": int(df.shape[1]), "pruned": pruned})
    except Exception as e:
        try:
            (UPLOAD_DIR / name).unlink()
        except OSError:
            pass
        return err("解析失败: %s" % e)


@app.get("/api/data/<name>")
def api_data(name):
    try:
        sheet = request.args.get("sheet")
        page = max(1, int(request.args.get("page", 1)))
        size = min(500, max(10, int(request.args.get("size", 50))))
        sheet, df = pick_sheet(name, sheet)
        df, added, mmeta, e = maybe_augment(df)
        if e:
            return err(*e)
        total = len(df)
        pages = max(1, math.ceil(total / size))
        page = min(page, pages)
        part = df.iloc[(page - 1) * size: page * size]
        cols = [{"name": c, "dtype": str(df[c].dtype), "from_mapping": c in added}
                for c in df.columns]
        rows = json.loads(json.dumps(part.to_dict(orient="records"),
                                     ensure_ascii=False, default=_default))
        payload = {"sheet": sheet, "sheets": list(load_workbook(name)), "columns": cols,
                   "rows": rows, "total": total, "page": page, "pages": pages, "size": size}
        if mmeta:
            payload["mapping"] = mmeta
        return ok(payload)
    except FileNotFoundError:
        return err("文件不存在", 404)
    except Exception as e:
        return err("读取数据失败: %s" % e, 500)


@app.get("/api/mappings")
def api_mappings_list():
    items = []
    for p in sorted(MAPPING_DIR.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
        m = json.loads(p.read_text(encoding="utf-8"))
        items.append({
            "name": m["name"],
            "key": m["key"],
            "columns": m["columns"],
            "rows": len(m["data"]),
            "updated": m.get("updated", ""),
        })
    return ok(items)


@app.post("/api/mapping/upload")
def api_mapping_upload():
    name = (request.form.get("name") or "").strip()
    key = (request.form.get("key") or "").strip()
    f = request.files.get("file")
    if not f or not f.filename:
        return err("未收到文件")
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return err("仅支持 .xlsx / .xls 文件")
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    tmp = UPLOAD_DIR / ("__map-" + ts + "-" + Path(f.filename).name)
    f.save(tmp)
    try:
        sheets = pd.read_excel(tmp, sheet_name=None)
        df = next(iter(sheets.values()))
        df.columns = [str(c) for c in df.columns]
        if len(df.columns) < 2:
            return err("匹配表至少需要 2 列（关联键 + 目标字段）")
        if not key:
            key = df.columns[0]
        if key not in df.columns:
            return err("关联键字段 %s 不在工作表中" % key)
        try:
            safe = _safe_name(name or Path(f.filename).stem)
        except ValueError:
            return err("匹配表名称不合法")
        m = {
            "name": safe,
            "key": key,
            "columns": list(df.columns),
            "data": df.fillna("").astype(str).to_dict(orient="records"),
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        save_mapping_obj(m)
        pruned = prune_mappings(_mapping_path(safe))   # 匹配表只保留最新的一份
        # 新匹配表自动成为缺省设置; 关联字段按数据里命中率最高的列自动选
        join = ""
        fname = latest_upload_name()
        if fname:
            try:
                cands = suggest_join_candidates(m, key, fname)
                if cands:
                    join = cands[0]["column"]
            except Exception:
                join = ""
        save_default_mapping(safe, join, key)
        return ok({"name": safe, "key": key, "columns": list(df.columns),
                   "rows": len(df), "pruned": pruned,
                   "default": {"name": safe, "join": join, "mkey": key}})
    except Exception as e:
        return err("解析失败: %s" % e)
    finally:
        try:
            tmp.unlink()
        except OSError:
            pass


@app.get("/api/mapping/<name>")
def api_mapping_get(name):
    try:
        m = load_mapping(name)
        preview = m["data"][:50]
        return ok({"name": m["name"], "key": m["key"], "columns": m["columns"],
                   "rows": len(m["data"]), "preview": preview, "updated": m.get("updated", "")})
    except FileNotFoundError:
        return err("匹配表不存在", 404)
    except ValueError as e:
        return err(str(e), 400)


@app.get("/api/mapping/<name>/suggest")
def api_mapping_suggest(name):
    """推荐数据侧的关联列: 对数据每一列计算与匹配表关键列的命中率, 按命中率排序."""
    try:
        m = load_mapping(name)
        mkey = (request.args.get("mkey", "").strip() or m.get("key") or "")
        if mkey not in m["columns"]:
            return err("匹配表中没有字段: " + mkey, 400)
        fname = request.args.get("file", "").strip() or latest_upload_name()
        if not fname:
            return ok({"mkey": mkey, "keys": 0, "candidates": []})
        sheet = request.args.get("sheet") or None
        cands = suggest_join_candidates(m, mkey, fname, sheet)
        keys = len({_norm_key(v) for v in pd.DataFrame(m["data"])[mkey]} - {""})
        return ok({"mkey": mkey, "keys": keys, "candidates": cands[:10]})
    except FileNotFoundError:
        return err("匹配表或数据文件不存在", 404)
    except Exception as e:
        return err("推荐失败: %s" % e, 500)


@app.get("/api/default-mapping")
def api_default_mapping_get():
    """当前缺省匹配设置(打开页面时自动应用的那一份)"""
    return ok({"mapping": load_default_mapping()})


@app.put("/api/default-mapping")
def api_default_mapping_put():
    body = request.get_json(force=True, silent=True) or {}
    name = (body.get("name") or "").strip()
    if not name:
        return err("缺少匹配表名称")
    if not _mapping_path(name).exists():
        return err("匹配表不存在: " + name, 404)
    save_default_mapping(name, body.get("join", ""), body.get("mkey", ""))
    return ok({"mapping": load_default_mapping()})


@app.delete("/api/default-mapping")
def api_default_mapping_delete():
    try:
        DEFAULT_MAPPING_PATH.unlink()
    except OSError:
        pass
    return ok({"mapping": None})


@app.delete("/api/mapping/<name>")
def api_mapping_delete(name):
    try:
        d = load_default_mapping()      # 删除前先看它是不是缺省
        delete_mapping(name)
        if d and d.get("name") == name:
            try:
                DEFAULT_MAPPING_PATH.unlink()   # 缺省指向已删的表 → 一并清掉
            except OSError:
                pass
        return ok({"deleted": name})
    except ValueError as e:
        return err(str(e), 400)


@app.get("/api/config")
def api_config_get():
    return ok({"product_categories": load_product_categories(),
               "default": DEFAULT_CATEGORIES})


@app.put("/api/config")
def api_config_put():
    try:
        body = request.get_json(force=True, silent=False) or {}
        cats = body.get("product_categories")
        if not isinstance(cats, dict):
            return err("product_categories 必须为字典 {类别名: [产品列表]}")
        out = {}
        for k, v in cats.items():
            k = str(k).strip()
            if not k:
                continue
            vs = [str(x).strip() for x in (v or []) if str(x).strip()]
            if vs:
                out[k] = vs
        if not out:
            return err("至少需要一个有效的产品分类")
        save_product_categories(out)
        return ok({"product_categories": out})
    except Exception as e:
        return err("保存失败: %s" % e)


@app.get("/api/analyze/<name>/regions")
def api_analyze_regions(name):
    """返回应用匹配后数据里所有含"区域"列的去重值, 供前端区域下拉使用."""
    try:
        sheet = request.args.get("sheet")
        sheet, df = pick_sheet(name, sheet)
        df, _added, _mmeta, e = maybe_augment(df)
        if e:
            return err(*e)
        region_col = next((c for c in df.columns if "区域" in c), "")
        if not region_col:
            return ok({"column": "", "regions": []})
        vals = sorted({str(v).strip() for v in df[region_col].dropna()
                       if str(v).strip() and str(v).strip().lower() not in ("nan", "none")})
        return ok({"column": region_col, "regions": vals})
    except FileNotFoundError:
        return err("文件不存在", 404)
    except Exception as e:
        return err("获取区域失败: %s" % e, 500)


@app.get("/api/analyze/<name>")
def api_analyze(name):
    try:
        sheet = request.args.get("sheet")
        report = request.args.get("report", "")
        region = request.args.get("region", "").strip()
        freq_key = request.args.get("freq", "day")
        date_col = request.args.get("date_col", "").strip()
        product_col = request.args.get("product_col", "").strip()
        factory_col = request.args.get("factory_col", "").strip()
        delivery_col = request.args.get("delivery_col", "").strip()
        volume_col = request.args.get("volume_col", "").strip()

        sheet, df = pick_sheet(name, sheet)
        df, _added, _mmeta, e = maybe_augment(df)
        if e:
            return err(*e)

        cols = list(df.columns)
        date_col = date_col or autodetect_col(cols, DATE_ALIASES)
        product_col = product_col or autodetect_col(cols, PRODUCT_ALIASES)
        # 工厂列: 优先用匹配表带来的「工厂名称」(同名工厂自动合并), 否则用别名识别
        if not factory_col:
            factory_col = "工厂名称" if "工厂名称" in cols else ""
        factory_col = factory_col or autodetect_col(cols, FACTORY_ALIASES)
        delivery_col = delivery_col or autodetect_col(cols, DELIVERY_ALIASES)
        volume_col = volume_col or autodetect_col(cols, VOLUME_ALIASES)
        required = {"日期字段": date_col, "产品字段": product_col, "工厂字段": factory_col,
                    "配送类型字段": delivery_col, "装车量字段": volume_col}
        missing = [k for k, v in required.items() if not v]
        if missing:
            return err("未能自动识别: " + "、".join(missing) +
                       "（请确认数据列名包含日期/产品/工厂/配送类型/装车量相关字段）")
        cats = load_product_categories()
        df["品类"] = df[product_col].apply(lambda v: categorize_product(v, cats))

        if region:
            region_col = next((c for c in df.columns if "区域" in c), "")
            if not region_col:
                return err("已选择区域「%s」，但需要先在匹配表里把「区域」字段加进来" % region)
            df = df[df[region_col].astype(str) == region]

        detected = {"date_col": date_col, "product_col": product_col,
                    "factory_col": factory_col, "delivery_col": delivery_col,
                    "volume_col": volume_col}
        cat_order = list(cats.keys())

        def last7_mask(dts):
            """近 7 天且不含今天; 若数据整体早于该窗口, 则回退为数据最大日期的前 7 天(含最大日)."""
            today = pd.Timestamp.now().normalize()
            end = today
            dmax = dts.max()
            if pd.notna(dmax) and dmax < today - pd.Timedelta(days=7):
                end = dmax + pd.Timedelta(days=1)
            start = end - pd.Timedelta(days=7)
            return (dts >= start) & (dts < end)

        if report == "product_trend":
            dts = pd.to_datetime(df[date_col], errors="coerce")
            mask = dts.notna() & last7_mask(dts)
            sub = df.loc[mask].copy(); sub["_d"] = dts[mask]
            freq_code, fmt = FREQ.get(freq_key, ("D", "%m-%d"))
            g = sub.groupby([pd.Grouper(key="_d", freq=freq_code), "品类"])[volume_col].sum().reset_index()
            pv = g.pivot(index="_d", columns="品类", values=volume_col).fillna(0).sort_index()
            labels = [d.strftime(fmt) for d in pv.index]
            series = []
            for cat_name in cat_order:
                short = cat_short(cat_name)
                if cat_name in pv.columns:
                    series.append({"name": short, "category": short,
                                   "data": [float(v) for v in pv[cat_name].values]})
                else:
                    series.append({"name": short, "category": short, "data": [0.0] * len(labels)})
            return ok({"labels": labels, "last_label": labels[-1] if labels else "",
                       "series": series, "detected": detected,
                       "region": region, "n": int(len(df))})

        if report == "product_pivot":
            # 与「产品趋势」同口径: 近 7 天(不含今天)
            # 行 = 配送类型 × 产品品类, 列 = 日期
            empty_pv = {"labels": [], "deliveries": [], "cats": [],
                        "cells": [], "date_totals": [], "grand_total": 0.0,
                        "detected": detected, "region": region, "n": int(len(df))}
            dts = pd.to_datetime(df[date_col], errors="coerce")
            mask = dts.notna() & last7_mask(dts)
            sub = df.loc[mask].copy()
            cat_set = set(cat_order)
            sub = sub[sub["品类"].isin(cat_set)]
            if sub.empty:
                return ok(empty_pv)
            sub["_d"] = dts[mask].dt.normalize()   # 只保留日期(装货时间可能带时分秒)
            g = sub.groupby(["_d", delivery_col, "品类"])[volume_col].sum().reset_index()
            dates = sorted(set(g["_d"]))
            labels = [d.strftime("%m-%d") for d in dates]
            idx = {d: i for i, d in enumerate(dates)}
            deliveries = sorted({str(v) for v in g[delivery_col]})
            raw = {}
            for _, r in g.iterrows():
                k = (str(r[delivery_col]), r["品类"])
                arr = raw.setdefault(k, [0.0] * len(dates))
                arr[idx[r["_d"]]] += float(r[volume_col])
            cells, date_totals = [], [0.0] * len(dates)
            for dv in deliveries:
                for cat_name in cat_order:
                    arr = raw.get((dv, cat_name))
                    if arr is None:
                        continue            # 该配送类型下没有这个品类 → 不生成空行
                    cells.append({"delivery": dv, "cat": cat_short(cat_name),
                                  "values": [round(v, 2) for v in arr]})
                    for j, v in enumerate(arr):
                        date_totals[j] += v
            return ok({"labels": labels, "deliveries": deliveries,
                       "cats": [cat_short(c) for c in cat_order],
                       "cells": cells,
                       "date_totals": [round(v, 2) for v in date_totals],
                       "grand_total": round(sum(date_totals), 2),
                       "detected": detected, "region": region, "n": int(len(df))})

        if report == "factory_trend":
            dts = pd.to_datetime(df[date_col], errors="coerce")
            mask = dts.notna() & last7_mask(dts)
            sub = df.loc[mask].copy(); sub["_d"] = dts[mask]
            freq_code, fmt = FREQ.get(freq_key, ("D", "%m-%d"))
            # 工厂列做空值兜底
            raw_f = sub[factory_col]
            fvals = raw_f.where(raw_f.notna(), "").astype(str).str.strip()
            fvals = fvals.where(~fvals.isin(["", "nan", "None", "<NA>"]), "（未匹配）")
            sub["_f"] = fvals
            g = sub.groupby([pd.Grouper(key="_d", freq=freq_code), "_f", "品类"])[volume_col].sum().reset_index()
            # 仅展示已配置产品类别，过滤"未分类"
            cat_set = set(cat_order)
            g = g[g["品类"].isin(cat_set)]
            date_vals = sorted(set(g["_d"]))
            labels = [d.strftime(fmt) for d in date_vals]
            idx = {d: i for i, d in enumerate(date_vals)}
            combos = sorted(set(zip(g["_f"], g["品类"])),
                            key=lambda x: (x[0], cat_order.index(x[1]) if x[1] in cat_order else 99))
            series = []
            for f, cat_name in combos:
                data = [0.0] * len(date_vals)
                sub2 = g[(g["_f"] == f) & (g["品类"] == cat_name)]
                for _, r in sub2.iterrows():
                    data[idx[r["_d"]]] = float(r[volume_col])
                series.append({"name": "%s·%s" % (f, cat_short(cat_name)),
                               "factory": f, "category": cat_short(cat_name), "data": data})
            return ok({"labels": labels, "last_label": labels[-1] if labels else "",
                       "series": series, "detected": detected,
                       "region": region, "n": int(len(df))})

        if report == "factory_pivot":
            # 与「工厂 × 产品趋势」同口径: 近 7 天(不含今天)
            # 逐厂返回: 行 = 配送类型 × 品类, 列 = 日期
            empty_fp = {"labels": [], "deliveries": [], "cats": [], "series": [],
                        "detected": detected, "region": region, "n": int(len(df))}
            dts = pd.to_datetime(df[date_col], errors="coerce")
            mask = dts.notna() & last7_mask(dts)
            sub = df.loc[mask].copy()
            # 工厂列做空值兜底（避免 None 与 str 比较时类型冲突）
            raw_f = sub[factory_col]
            fvals = raw_f.where(raw_f.notna(), "").astype(str).str.strip()
            # pandas str(None) 在不同版本下产生 "None"/"nan"/""，统一归一
            fvals = fvals.where(~fvals.isin(["", "nan", "None", "<NA>"]), "（未匹配）")
            sub = sub.assign(_f=fvals)
            # 仅保留已配置产品类别，过滤掉"未分类"
            cat_set = set(cat_order)
            sub = sub[sub["品类"].isin(cat_set)]
            if sub.empty:
                return ok(empty_fp)
            sub["_d"] = dts[mask].dt.normalize()
            g = sub.groupby(["_f", "_d", delivery_col, "品类"])[volume_col].sum().reset_index()
            dates = sorted(set(g["_d"]))
            labels = [d.strftime("%m-%d") for d in dates]
            idx = {d: i for i, d in enumerate(dates)}
            deliveries = sorted({str(v) for v in g[delivery_col]})
            combos = sorted(set(zip(g["_f"], g["品类"])),
                            key=lambda x: (x[0], cat_order.index(x[1])))
            series = []
            for f, cat_name in combos:
                for dv in deliveries:
                    sub2 = g[(g["_f"] == f) & (g["品类"] == cat_name) &
                             (g[delivery_col].astype(str) == dv)]
                    if sub2.empty:
                        continue
                    data = [0.0] * len(dates)
                    for _, r in sub2.iterrows():
                        data[idx[r["_d"]]] = round(float(r[volume_col]), 2)
                    series.append({"factory": f, "delivery": dv,
                                   "category": cat_short(cat_name), "data": data})
            return ok({"labels": labels, "deliveries": deliveries,
                       "cats": [cat_short(c) for c in cat_order],
                       "series": series,
                       "detected": detected, "region": region, "n": int(len(df))})

        return err("未知的报告类型: " + report)
    except FileNotFoundError:
        return err("文件不存在", 404)
    except ValueError as e:
        return err(str(e), 400)
    except Exception as e:
        return err("分析失败: %s" % e, 500)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8787))
    # 线上发布要求绑定 0.0.0.0; 本地想收紧可设 HOST=127.0.0.1
    host = os.environ.get("HOST", "0.0.0.0")
    print("区域工厂液体装货趋势已启动: http://localhost:%d (监听 %s)" % (port, host))
    app.run(host=host, port=port, threaded=True)
