"""
过程数据分析平台 - Flask 后端
功能：上传Excel/截图 → 自动存数据库 → 回归分析 → 返回结果JSON
数据库：SQLite，所有上传数据自动累积，分析基于全量历史数据
"""
import os
import io
import re
import csv
import json
import hashlib
import hmac
from urllib.parse import quote
import sqlite3
import numpy as np
import pandas as pd
import statsmodels.api as sm
import math
import requests
from flask import Flask, render_template, request, jsonify, Response
from datetime import datetime, timezone, timedelta


def _safe_float(val, nan_default=0.0):
    """将 val 转为 float；NaN/inf 替换为 nan_default（避免 jsonify 输出非法 JSON）"""
    try:
        v = float(val)
        if math.isnan(v) or math.isinf(v):
            return nan_default
        return v
    except (TypeError, ValueError):
        return nan_default

app = Flask(__name__)
CST = timezone(timedelta(hours=8))
BASE_DIR = os.path.dirname(__file__)
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
DB_PATH = os.path.join(BASE_DIR, "data.db")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# 修改/删除数据的访问口令（仅平台主人持有）。口令经 sha256(盐+口令) 哈希后存盘，
# 明文不落库；每次修改/删除请求必须携带正确口令，否则服务端拒绝。
_PASS_SALT = "ai_process_tool_edit_guard_2026"


def _load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _hash_pass(p):
    return hashlib.sha256((_PASS_SALT + str(p)).encode("utf-8")).hexdigest()


def _pass_configured():
    return bool(_load_config().get("edit_pass_hash"))


def _check_pass(p):
    cfg = _load_config()
    h = cfg.get("edit_pass_hash")
    if not h:
        return False
    return hmac.compare_digest(h, _hash_pass(p))
os.makedirs(UPLOAD_DIR, exist_ok=True)

def _ensure_default_pass():
    """防呆：若从未设置过编辑口令，自动用默认口令 zhihao2026 初始化，
    保证开箱即用、不至于忘记口令后彻底进不去。已设置过则不动。"""
    cfg = _load_config()
    if not cfg.get("edit_pass_hash"):
        cfg["edit_pass_hash"] = _hash_pass("zhihao2026")
        _save_config(cfg)

# ── 数据库初始化 ──────────────────────────────────────────
def init_db():
    """初始化 SQLite 数据库"""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period TEXT NOT NULL,
            name TEXT NOT NULL,
            subject TEXT DEFAULT '数学',
            friend_rate REAL,
            group_rate REAL,
            app_rate REAL,
            reply_rate REAL,
            day1_rate REAL,
            conversion_rate REAL,
            day3_rate REAL,
            source TEXT DEFAULT 'upload',
            uploaded_at TEXT DEFAULT (datetime('now', 'localtime')),
            grade TEXT DEFAULT ''
        )
    """)
    # 兼容旧库：补充 subject 列（学科维度：数学/语文/英语…）
    try:
        conn.execute("ALTER TABLE records ADD COLUMN subject TEXT DEFAULT '数学'")
    except Exception:
        pass
    # 唯一约束升级：加入 subject，避免不同学科同名同人冲突
    conn.execute("DROP INDEX IF EXISTS idx_grade_period_name")
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_sgp_name 
        ON records(subject, grade, period, name)
    """)
    # 下期预测记录：对某期(可选某人)的预测值，待真实数据上传后做对比
    conn.execute("""
        CREATE TABLE IF NOT EXISTS predictions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period TEXT NOT NULL,
            name TEXT DEFAULT '',
            subject TEXT DEFAULT '数学',
            predicted_day3 REAL,
            predicted_final REAL,
            metrics_json TEXT DEFAULT '{}',
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)
    try:
        conn.execute("ALTER TABLE predictions ADD COLUMN subject TEXT DEFAULT '数学'")
    except Exception:
        pass
    conn.execute("DROP INDEX IF EXISTS idx_pred_period_name")
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_pred_s_period_name 
        ON predictions(subject, period, name)
    """)
    conn.commit()
    conn.close()

init_db()
_ensure_default_pass()

def get_db():
    """获取数据库连接"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# ── 数据库操作 ──────────────────────────────────────────

def _normalize_period(p):
    """把期次统一成纯数字串，避免 '0717'/'717'/'00717'/'7.0' 这类格式共存。
    正常返回如 '717'，非法输入原样返回。"""
    try:
        return str(int(round(float(str(p).strip()))))
    except (ValueError, TypeError):
        return str(p).strip()

def check_grade_conflicts(rows, default_grade="", default_subject=None):
    """录入前预检：同一人（同姓名、同学科）在库中只能有一个年级。
    若本次录入会把该人挂到与已有记录不同的年级，返回冲突文案列表（空=无冲突）。
    同时检查本批次内是否出现同一人多年级。"""
    default_subject = default_subject or "数学"
    batch = {}  # (姓名, 学科) -> set(年级)
    for r in rows:
        name = str(r.get("姓名", "")).strip()
        if not name:
            continue
        subj = str(r.get("学科", default_subject) or default_subject).strip() or default_subject
        g = str(r.get("年级", default_grade) or default_grade).strip()
        batch.setdefault((name, subj), set())
        if g:
            batch[(name, subj)].add(g)
    conflicts = []
    for (name, subj), gs in batch.items():
        if len(gs) > 1:
            conflicts.append(f"「{name}」（{subj}）本次同时出现在多个年级（{', '.join(sorted(gs))}），请确认其唯一年级")
        elif not gs and not default_grade:
            conflicts.append(f"「{name}」（{subj}）未填写年级且未选择默认年级，无法确定其年级，已阻止录入")
    conn = get_db()
    for (name, subj), gs in batch.items():
        existing = conn.execute(
            "SELECT DISTINCT grade FROM records WHERE name=? AND subject=?", (name, subj)
        ).fetchall()
        ex = set(str(r[0]) for r in existing if r[0])
        for g in gs:
            if g and ex and g not in ex:
                conflicts.append(
                    f"「{name}」（{subj}）已存在于【{', '.join(sorted(ex))}】，不能在【{g}】录入——"
                    f"每人年级应唯一，请先确认正确年级，或删除该人旧年级记录后再传"
                )
    conn.close()
    return conflicts


def save_to_db(rows, source="upload", grade="", subject=None):
    """将解析后的数据行插入 SQLite，重复的 (学科, 年级, 期次, 姓名) 会更新。
    若本次未指定学科/年级，自动沿用该(期次,姓名)已有记录的学科/年级，避免产生重复行。"""
    conn = get_db()
    # 预取已有 (期次,姓名) -> (学科, 年级) 映射，用于空学科/空年级时回填
    existing_grade = {}
    existing_subject = {}
    for r in conn.execute("SELECT period, name, grade, subject FROM records").fetchall():
        existing_grade.setdefault((str(r["period"]), str(r["name"])), str(r["grade"]))
        existing_subject.setdefault((str(r["period"]), str(r["name"])), str(r["subject"]))
    inserted = 0
    updated = 0
    for row in rows:
        period = _normalize_period(row.get("期次", ""))
        name = str(row.get("姓名", ""))
        row_subject = str(row.get("学科", "") or subject or "数学").strip() or "数学"
        row_grade = str(row.get("年级", grade or "")).strip()
        # 空年级时，沿用该人已有年级（优先非空）
        if not row_grade:
            eg = existing_grade.get((period, name), "")
            if eg:
                row_grade = eg
        # 行内未指定学科时，沿用该人已有学科，避免把数据误标到别的学科
        if not str(row.get("学科", "")).strip():
            es = existing_subject.get((period, name), "")
            if es:
                row_subject = es
        if not period or not name:
            continue
        try:
            conn.execute("""
                INSERT INTO records 
                    (period, name, subject, friend_rate, group_rate, app_rate, reply_rate, day1_rate, conversion_rate, day3_rate, source, grade)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(subject, grade, period, name) DO UPDATE SET
                    friend_rate=excluded.friend_rate,
                    group_rate=excluded.group_rate,
                    app_rate=excluded.app_rate,
                    reply_rate=excluded.reply_rate,
                    day1_rate=excluded.day1_rate,
                    conversion_rate=excluded.conversion_rate,
                    day3_rate=excluded.day3_rate,
                    subject=excluded.subject,
                    source=excluded.source,
                    grade=excluded.grade,
                    uploaded_at=datetime('now', 'localtime')
            """, (
                period, name, row_subject,
                float(row.get("好友率", 0) or 0),
                float(row.get("入群率", 0) or 0),
                float(row.get("APP下载率", 0) or 0),
                float(row.get("开营回复率", 0) or 0),
                float(row.get("Day1到课率", 0) or 0),
                float(row.get("最终转化率", 0) or 0),
                float(row.get("Day3出直播间转化率", 0) or 0),
                source,
                row_grade
            ))
            if conn.total_changes > inserted + updated:
                inserted += 1
        except Exception as e:
            print(f"[DB] 插入失败 {row_subject}/{row_grade}/{period}/{name}: {e}")
    conn.commit()
    conn.close()
    return inserted

def get_all_data(grade=None, subject=None):
    """训练集唯一入口：仅读取 records 表中真实录入的数据（可按年级/学科筛选）。

    重要约束：AI 预测结果存于独立的 predictions 表，绝不会被读入训练集，
    因此预测值不会"回流"污染下次预测。任何预测/分析/回归都必须经此函数取数。
    """
    conn = get_db()
    where = []
    params = []
    if grade:
        where.append("grade = ?"); params.append(grade)
    if subject:
        where.append("subject = ?"); params.append(subject)
    wsql = ("WHERE " + " AND ".join(where)) if where else ""
    rows = conn.execute(
        "SELECT period, name, subject, friend_rate, group_rate, app_rate, "
        "reply_rate, day1_rate, conversion_rate, day3_rate, "
        "source, uploaded_at, grade "
        "FROM records " + wsql + " ORDER BY subject, grade, period, name",
        params
    ).fetchall()
    conn.close()
    
    if not rows:
        return None
    
    data = []
    for r in rows:
        data.append({
            "期次": r["period"],
            "姓名": r["name"],
            "学科": r["subject"],
            "好友率": r["friend_rate"],
            "入群率": r["group_rate"],
            "APP下载率": r["app_rate"],
            "开营回复率": r["reply_rate"],
            "Day1到课率": r["day1_rate"],
            "最终转化率": r["conversion_rate"],
            # day3_rate 与 conversion_rate 同口径：均为百分比数值（分母=全部好友），不做额外换算。
            # 验证：全量各年级 Day3 原始值均 < 最终转化率，符合同分母漏斗；此前×100是错误放大。
            "Day3出直播间转化率": (round(float(r["day3_rate"]), 2) if r["day3_rate"] is not None else None),
            "年级": r["grade"],
        })
    return pd.DataFrame(data)

def get_db_stats(grade=None, subject=None):
    """获取数据库统计信息，可按年级/学科筛选"""
    conn = get_db()
    stats = {}
    where = []
    params = []
    if grade:
        where.append("grade = ?"); params.append(grade)
    if subject:
        where.append("subject = ?"); params.append(subject)
    wsql = ("WHERE " + " AND ".join(where)) if where else ""
    
    row = conn.execute(f"SELECT COUNT(*) as cnt FROM records {wsql}", params).fetchone()
    stats["total_rows"] = row["cnt"]
    row = conn.execute(f"SELECT COUNT(DISTINCT period) as cnt FROM records {wsql}", params).fetchone()
    stats["total_batches"] = row["cnt"]
    row = conn.execute(f"SELECT COUNT(DISTINCT name) as cnt FROM records {wsql}", params).fetchone()
    stats["total_people"] = row["cnt"]
    row = conn.execute(f"SELECT source, COUNT(*) as cnt FROM records {wsql} GROUP BY source", params).fetchall()
    stats["sources"] = {r["source"]: r["cnt"] for r in row}
    row = conn.execute(f"SELECT MAX(uploaded_at) as cnt FROM records {wsql}", params).fetchone()
    stats["last_upload"] = row["cnt"]
    
    # 按年级分组统计（受 subject 过滤影响）
    gsql = "SELECT grade, COUNT(*) as cnt, COUNT(DISTINCT period) as periods, COUNT(DISTINCT name) as people FROM records " + wsql + " GROUP BY grade ORDER BY grade"
    grades = conn.execute(gsql, params).fetchall()
    stats["grades"] = {r["grade"] or "未设置": {"count": r["cnt"], "periods": r["periods"], "people": r["people"]} for r in grades}
    
    # 所有年级列表
    all_grades = conn.execute("SELECT DISTINCT grade FROM records WHERE grade != '' ORDER BY grade").fetchall()
    stats["grade_list"] = [r["grade"] for r in all_grades]
    
    # 按学科分组统计（受 grade 过滤影响）
    subj_sql = "SELECT subject, COUNT(*) as cnt FROM records " + wsql + " GROUP BY subject ORDER BY subject"
    subjs = conn.execute(subj_sql, params).fetchall()
    stats["subjects"] = {r["subject"] or "未设置": r["cnt"] for r in subjs}
    all_subjects = conn.execute("SELECT DISTINCT subject FROM records ORDER BY subject").fetchall()
    stats["subject_list"] = [r["subject"] for r in all_subjects]
    
    conn.close()
    return stats


# ── 下期预测（杠杆模型）──────────────────────────────────
# 用全量历史数据拟合 OLS(含暑假哑变量)，由过程指标预测结果指标
PREDICT_X = ["好友率", "入群率", "APP下载率", "开营回复率", "Day1到课率"]
PREDICT_Y = ["最终转化率", "Day3出直播间转化率"]


def _summer_flag(period):
    """期次 >= 0622 视为暑假（与 run_analysis 口径一致）"""
    try:
        return 1 if str(int(float(str(period)))).zfill(4) >= "0622" else 0
    except Exception:
        return 0


def predict_conversion(period, metrics, grade=None, subject=None, name=None):
    """
    根据输入的过程指标，预测 Day3/最终转化率。
    metrics: {好友率, 入群率, APP下载率, 开营回复率, Day1到课率}（数值）

    预测口径：
    - name 非空：只用该人「个人历史」，不被团队其他人数据淹没（贴合个人闭环）。
                个人样本少，不用全变量 OLS（会过拟合到 R²=1 且系数反号），
                改用「近邻法」：找历史上过程指标与你输入最像的几期，用它们
                的真实转化率均值当预测——单调、符合直觉、不会反直觉。
    - name 为空：用全年级学科「团队整体」做 OLS 回归（样本足，较稳）。

    Day3 与最终转化率同口径（百分比数值，分母=全部好友），正常区间 0~几个百分点；
    先前×100误判的"翻车值"已不存在，无需特殊剔除。

    返回 {最终转化率, Day3出直播间转化率, 个人模式, 样本数}
    """
    base = get_all_data(grade=grade if grade else None, subject=subject if subject else None)
    if base is None or len(base) == 0:
        return None
    if name:
        return _predict_by_person(base, name, metrics)
    # 团队模式：OLS + Day3 鲁棒
    df = base.copy()
    df["暑假"] = df["期次"].apply(
        lambda x: 1 if str(int(float(str(x)))).zfill(4) >= "0622" else 0
    )
    summer = _summer_flag(period)
    result = {"个人模式": False, "样本数": int(len(df))}
    for y_var in PREDICT_Y:
        X_all = df[[*PREDICT_X, "暑假"]]
        y_all = df[y_var]
        combo = pd.concat([X_all, y_all], axis=1).dropna()
        if len(combo) < 5:
            result[y_var] = None
            continue
        # Day3 与最终同口径(百分比数值,分母=全部好友)，正常区间 0~几个百分点，无需翻车剔除
        X = combo[[*PREDICT_X, "暑假"]]
        y = combo[y_var]
        X_sm = sm.add_constant(X)
        try:
            model = sm.OLS(y, X_sm).fit()
            const = _safe_float(model.params.get("const", 0))
            summer_coef = _safe_float(model.params.get("暑假", 0))
            val = const + summer_coef * summer
            for x in PREDICT_X:
                val += _safe_float(model.params.get(x, 0)) * _safe_float(metrics.get(x, 0) or 0)
            val = max(0, min(100, val))
            result[y_var] = round(float(val), 2)
        except Exception:
            result[y_var] = None
    return result


def _predict_by_person(base, name, metrics):
    """个人模式：近邻法。找过程指标最相似的 K 期真实值均值作为预测。"""
    df = base[base["姓名"] == name].copy()
    if len(df) == 0:
        return None
    df["暑假"] = df["期次"].apply(
        lambda x: 1 if str(int(float(str(x)))).zfill(4) >= "0622" else 0
    )
    result = {"个人模式": True, "样本数": int(len(df))}
    for y_var in PREDICT_Y:
        sub = df[[*PREDICT_X, y_var]].dropna()
        # Day3 同口径百分比数值，正常区间 0~几个百分点，无需翻车剔除
        if len(sub) < 3:
            result[y_var] = None
            continue
        # 与输入指标欧氏距离最小的 K 期
        K = min(3, len(sub))
        def _dist(row):
            return sum((float(row[x]) - float(metrics.get(x, 0) or 0)) ** 2 for x in PREDICT_X)
        dist = sub.apply(_dist, axis=1)
        neigh = sub.loc[dist.nsmallest(K).index]
        val = float(neigh[y_var].mean())
        val = max(0, min(100, round(val, 2)))
        result[y_var] = val
    return result


def latest_metrics(grade=None, subject=None):
    """返回最新一期的团队均值过程指标，用于预测表单预填"""
    df = get_all_data(grade=grade if grade else None, subject=subject if subject else None)
    if df is None or len(df) == 0:
        return None
    # 数值排序，避免 "0717" 被错误地排在最前
    def _pk(p):
        try:
            return int(float(str(p)))
        except (ValueError, TypeError):
            return -1
    periods = sorted(df["期次"].astype(str).unique().tolist(), key=_pk)
    latest = periods[-1]
    sub = df[df["期次"].astype(str) == latest]
    metrics = {x: round(float(sub[x].dropna().mean()), 2) for x in PREDICT_X}
    return {"period": latest, "metrics": metrics}


# ── 分析核心函数 ──────────────────────────────────────────

def run_analysis(df):
    """对所有X变量跑单变量回归（控制暑假哑变量），返回完整结果。

    入参 df 必须来自 get_all_data()，即仅含真实录入数据；预测表不参与。
    """
    x_vars = ["好友率", "入群率", "APP下载率", "开营回复率", "Day1到课率"]
    y_vars = ["最终转化率", "Day3出直播间转化率"]
    
    # 识别暑假期次
    df["暑假"] = df["期次"].apply(
        lambda x: 1 if str(int(float(str(x)))).zfill(4) >= "0622" else 0
    )
    
    results = {"final": [], "day3": [], "path": [], "descriptive": {}, "summer_effect": {}}
    
    # 描述性统计
    for col in x_vars + y_vars:
        vals = df[col].dropna()
        results["descriptive"][col] = {
            "mean": round(float(vals.mean()), 2),
            "std": round(float(vals.std(ddof=1)), 3),
            "min": round(float(vals.min()), 2),
            "max": round(float(vals.max()), 2),
            "n": int(len(vals))
        }
    
    # 暑假效应
    for y in y_vars:
        ns = df[df["暑假"] == 0][y].mean()
        sm_val = df[df["暑假"] == 1][y].mean()
        results["summer_effect"][y] = {
            "nonsummer": round(_safe_float(ns), 2),
            "summer": round(_safe_float(sm_val), 2),
            "gap": round(_safe_float(sm_val - ns), 2)
        }
    
    # 单变量回归
    for y_var in y_vars:
        key = "final" if y_var == "最终转化率" else "day3"
        for x_var in x_vars:
            # 同时清除X和y的NaN（y含缺失值会导致OLS全量返回NaN）
            X = df[[x_var, "暑假", y_var]].dropna()
            y = X.pop(y_var)
            X_sm = sm.add_constant(X)
            try:
                model = sm.OLS(y, X_sm).fit()
                coef = model.params.get(x_var, 0)
                pval = model.pvalues.get(x_var, 1)
                r2 = model.rsquared
                results[key].append({
                    "variable": x_var,
                    "coefficient": round(_safe_float(coef), 4),
                    "p_value": round(_safe_float(pval, nan_default=1.0), 4),
                    "significant": bool(_safe_float(pval, nan_default=1.0) < 0.05),
                    "r_squared": round(_safe_float(r2), 3)
                })
            except Exception:
                results[key].append({
                    "variable": x_var, "coefficient": 0, "p_value": 1,
                    "significant": False, "r_squared": 0
                })
        results[key].sort(key=lambda x: abs(x["coefficient"]), reverse=True)
    
    # 路径分析
    for x_var in ["好友率", "入群率", "APP下载率", "开营回复率"]:
        X = df[[x_var, "暑假", "Day1到课率"]].dropna()
        y = X.pop("Day1到课率")
        X_sm = sm.add_constant(X)
        try:
            model = sm.OLS(y, X_sm).fit()
            coef_to_attend = model.params.get(x_var, 0)
            pval = model.pvalues.get(x_var, 1)
            attend_effect = find_coef(results["final"], "Day1到课率")
            direct_effect = find_coef(results["final"], x_var)
            indirect_effect = coef_to_attend * attend_effect if attend_effect else 0
            total_effect = direct_effect + indirect_effect if direct_effect else indirect_effect
            results["path"].append({
                "variable": x_var,
                "coef_to_attend": round(_safe_float(coef_to_attend), 4),
                "p_value_to_attend": round(_safe_float(pval, nan_default=1.0), 4),
                "sig_to_attend": bool(_safe_float(pval, nan_default=1.0) < 0.05),
                "direct_effect": round(_safe_float(direct_effect), 4) if direct_effect else 0,
                "indirect_effect": round(_safe_float(indirect_effect), 4),
                "total_effect": round(_safe_float(total_effect), 4)
            })
        except:
            pass
    
    attend_coef = find_coef(results["final"], "Day1到课率")
    results["path"].append({
        "variable": "Day1到课率",
        "coef_to_attend": 1.0, "p_value_to_attend": 0, "sig_to_attend": True,
        "direct_effect": round(_safe_float(attend_coef), 4) if attend_coef else 0,
        "indirect_effect": 0,
        "total_effect": round(_safe_float(attend_coef), 4) if attend_coef else 0
    })
    results["path"].sort(key=lambda x: abs(x["total_effect"]), reverse=True)
    
    return results


def find_coef(results_list, var_name):
    for r in results_list:
        if r["variable"] == var_name:
            return r["coefficient"]
    return None


# ── AI建议生成 ──────────────────────────────────────────

# ───────────────────────────────────────────────────────────────
# 结构化建议库：每个短板环节给不同的、可直接落地的方案
# 字段：control=可控/上游不可控；note=过程定位；
#       sop=可抄进SOP的动作；materials=需做的物料图；ai=借助AI能做的
# 原则：可控环节给具体SOP+物料+AI；上游不可控环节(好友率)只标注+引导下游
# ───────────────────────────────────────────────────────────────
# 每环节准备多套不同角度的建议（变体），按期次稳定轮换：
# 同一期永远显示同一条，换一期自动换一套，避免"每次都是那一个建议"。
# 上游不可控的过程指标：运营无法决定家长加不加、也不能主动加，
# 平台不给"提升该指标"的改善方案，只展示数据并标注"不可控"。
UNCONTROLLABLE_VARS = {"好友率"}

ADVICE_LIB = {
    "好友率": [
        {
            "control": "上游不可控",
            "note": "好友率由前端BD引流质量决定（家长主动加你、你只通过），你控制不了家长加不加，也不能主动加。",
            "sop": [
                "好友申请当班内通过，通过好友后立刻拉群（拉群算入群率，是你发力的第一环）",
                "每周把『家长为什么不加』的真实原因（嫌麻烦/怕骚扰/忘了）整理Top3，反馈给BD优化加微引导卡",
                "用『家长画像分类表』把家长分焦虑/放养/观望三类，下游各环节话术按类型切换",
            ],
            "materials": [
                "【家长画像分类表】焦虑型/放养型/观望型 各特征 + 对应话术方向",
            ],
            "ai": [
                "让AI按历史对话给家长分三类画像，并归纳每类最常问的3个问题，预写『秒回话术库』——通过瞬间直接贴",
                "⚠️ 此环上游决定，别在加好友上耗动作，重心移下游",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，你唯一可控的是『及时通过+不漏接』，通过后立刻拉群。",
            "sop": [
                "每天早/午/晚3个固定时间点打开企微『新的朋友』，申请当班内通过",
                "通过好友后10分钟内必拉群（拉群计入入群率，是你能发力的第一环）",
                "通过率偏低的期次，重点复盘是不是漏接了申请、卡在哪一刻",
            ],
            "materials": [
                "【漏接检查清单】早/午/晚3次查『新的朋友』打卡表",
            ],
            "ai": [
                "让AI生成『通过即拉群』SOP检查清单，一步步照做不漏接",
                "⚠️ 此环上游决定，别耗动作，重心移下游",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，你只做及时通过；通过后第一句话就给价值，把『被加』变『被服务』。",
            "sop": [
                "通过好友后立刻发『资料已为你留好』价值锚点，别只回『你好』",
                "第一句话按家长画像分3版（焦虑/放养/观望），通过瞬间贴",
                "通过即拉群（计入入群率），别漏接",
            ],
            "materials": [
                "【通过即发价值话术】3版第一句话模板（资料已留好/本周重点/孩子专属）",
            ],
            "ai": [
                "让AI把通过好友后的『第一句话』从客套改成价值钩子（如『孩子本期资料已留好，进群领』）",
                "⚠️ 此环上游决定，别耗动作在加好友上，守住通过速度即可",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，你唯一可控的是『通过速度』；找出你拖的时段，优化值守。",
            "sop": [
                "每笔申请记录『申请时间→通过时间』间隔，找平均通过最慢的时段",
                "把最慢时段排进固定值守（如午休/晚9点后容易漏）",
                "通过即拉群，别漏接",
            ],
            "materials": [
                "【通过时长记录表】日期/时段/申请→通过间隔，标红最慢",
            ],
            "ai": [
                "让AI生成『通过时长看板』：按小时统计平均通过间隔，标红最慢时段",
                "⚠️ 此环上游决定，重心放下游",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，但你可以用『已通过家长』的真实动机反推引流优化，反馈BD。",
            "sop": [
                "整理已通过家长最常说的『为什么加你』（如『朋友推荐/想看资料』）",
                "把这些真实动机反馈给BD，推动引流端强化对应钩子",
                "通过即拉群，守住通过速度",
            ],
            "materials": [
                "【加微动机归纳表】『为什么加你』Top原因 + 对应BD优化动作",
            ],
            "ai": [
                "让AI从已通过家长对话里归纳『加微Top动机』，生成给BD的优化建议",
                "⚠️ 此环上游决定，别主动加家长",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，用企微「好友来源渠道」分组统计通过率，把高通过渠道反馈BD复制放大。",
            "sop": [
                "在企微后台按来源标签（扫码/搜索/名片）分组看通过率",
                "把通过率最高的渠道特征反馈BD（如『XX渠道来的家长最愿意加』）",
                "通过即拉群，别漏接",
            ],
            "materials": [
                "【渠道通过率看板】按来源统计通过率对比",
            ],
            "ai": [
                "让AI按来源渠道算通过率对比，标出最优质渠道，给BD写『高通过渠道特征』报告",
                "⚠️ 此环上游决定，你控制不了家长加不加，重心移下游",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，但对「已通过但沉默」家长用一条轻量干货维持联系，不催不扰。",
            "sop": [
                "通过好友后第3天发一条孩子相关干货（如『本周易错点提示』），不带任何要求",
                "只维护温度，不催下载不催课",
                "通过即拉群，别漏接",
            ],
            "materials": [
                "【轻量干货卡】一条无压力的孩子相关干货",
            ],
            "ai": [
                "让AI按孩子年级生成一条无压力干货，不发要求；标记『已通过但零互动』家长单独轻触",
                "⚠️ 不催不扰，守住通过速度即可",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，但能用企微自动化减少漏接：自动通过+欢迎语模板，仍不主动加。",
            "sop": [
                "在企微设『自动通过好友』+ 统一欢迎语模板，减少人工漏看",
                "欢迎语里直接带『资料已留好，进群领』价值锚点",
                "自动通过仅限已引流场景，绝不主动加家长",
            ],
            "materials": [
                "【自动欢迎语模板】通过即触发的价值钩子话术",
            ],
            "ai": [
                "让AI写一版通用自动欢迎语，通过瞬间发出不漏；按家长来源给不同欢迎语分支",
                "⚠️ 此环上游决定，自动化只减漏接，不主动加",
            ],
        },
        {
            "control": "上游不可控",
            "note": "好友率上游不可控，用『当日申请当日清』纪律防漏接。",
            "sop": [
                "每天固定3个时段（早/午/晚）清空企微『新的朋友』，隔夜申请不过夜",
                "漏接的申请当日补通过+拉群",
                "每周复盘漏接数，找最容易忘的时段加提醒",
            ],
            "materials": [
                "【漏接清零打卡表】早/午/晚3次清空『新的朋友』打卡",
            ],
            "ai": [
                "让AI生成『漏接清零』打卡清单，到点提醒你查；用看板盯每日漏接数倒逼值守",
                "⚠️ 此环上游决定，重心移下游",
            ],
        },
    ],
    "入群率": [
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，引导讲清群价值。",
            "sop": [
                "通过好友后立即私信发入群邀请（群二维码），别等攒一批一起发",
                "入群欢迎语一次讲清：群里有什么（资料/答疑/直播提醒）",
                "对24小时内未进群的家长，私信简单提醒1次",
                "把群价值具象成『每周X题精讲 + 每月诊断报告』",
            ],
            "materials": [
                "【入群引导卡】竖屏1080×1920：群二维码 + 三句话群价值 + 你的企微头像",
                "【往期家长真实好评截图卡】入群引导卡底部加2条真实好评，先建立信任再引导进群",
            ],
            "ai": [
                "把『进群』包装成『领开学礼』：让AI自动生成『进群即得：资料包+试听+1v1诊断』的群公告+私信（AI按年级写3版做A/B）",
                "让AI把入群欢迎语从『广播』改造成『1个让家长愿意回的钩子』（如『孩子专属资料已留好，进群领取』）",
                "按家长画像写不同入群钩子：焦虑型→『孩子专属资料已留好』；放养型→『每周一题精讲』；观望型→『先看看往期家长好评』",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用真实好评建立信任。",
            "sop": [
                "入群引导卡直接放『往期家长真实好评』，先建立信任再引导进群",
                "用『已进群家长』的真实反馈做钩子（如『XX妈妈：资料太实用了』）",
                "对24小时内未进群的家长，私信简单提醒1次",
            ],
            "materials": [
                "【往期家长好评截图卡】2-3条真实好评 + 群二维码",
                "【入群引导卡】竖屏：群二维码 + 群价值 + 真实好评",
            ],
            "ai": [
                "让AI从往期家长对话里挑出最打动人的3条好评，做成引导素材",
                "把『进群』包装成『领开学礼』（AI按年级写3版A/B）",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用限时福利制造紧迫感。",
            "sop": [
                "入群引导卡加『今晚8点前入群领XX』限时钩子",
                "对24小时内未进群的家长，私信简单提醒1次",
                "限时福利到期前1小时再补1次轻提醒（不反复@）",
            ],
            "materials": [
                "【入群引导卡(限时版)】竖屏：群二维码 + 『今晚8点前入群领XX』倒计时",
            ],
            "ai": [
                "让AI写限时福利话术（稀缺感，不假紧迫）：『名额/资料今晚截止』",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用群氛围和同伴进度促入群。",
            "sop": [
                "展示群内往期热闹截图 / 孩子作品，建立『大家都在』的氛围",
                "用『群里已有XX位家长』轻同伴压力",
                "对24小时内未进群的家长，私信简单提醒1次",
            ],
            "materials": [
                "【群氛围截图卡】往期群聊热闹/孩子作品截图 + 群二维码",
            ],
            "ai": [
                "让AI生成群氛围截图文案（突出『大家都在、孩子有收获』）",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用孩子真实作品建立『群里有干货』的预期。",
            "sop": [
                "入群引导卡放1张孩子真实作品/优秀作业截图",
                "私信讲清『群里每周发孩子作品精讲』",
                "24h内未进群轻提醒1次",
            ],
            "materials": [
                "【孩子作品展示卡】1张真实作业/作品 + 群二维码",
            ],
            "ai": [
                "让AI挑往期孩子优秀作业做成展示素材",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用「孩子本周学习日历」钩子。",
            "sop": [
                "入群引导卡放本周学习日历（哪天讲啥）",
                "私信讲『进群看完整安排，别错过孩子弱项那天的课』",
                "24h内未进群轻提醒1次",
            ],
            "materials": [
                "【本周学习日历卡】横版：每日主题 + 孩子弱项标注",
            ],
            "ai": [
                "让AI生成按年级的学习日历钩子（标出弱项那天的课）",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用「限时答疑名额」钩子。",
            "sop": [
                "入群引导突出『前XX名进群抢老师1v1答疑位』",
                "对24h内未进群轻提醒1次",
                "名额用真实库存（如『本周只留5个答疑位』）",
            ],
            "materials": [
                "【答疑名额卡】竖屏：限时答疑位 + 群二维码",
            ],
            "ai": [
                "让AI写答疑名额钩子（真实稀缺，不假紧迫）",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用『群快满』真实稀缺感促入群。",
            "sop": [
                "入群引导卡标注『本群限定200人，已190』真实进度",
                "对24h内未进群轻提醒1次，强调名额真实",
                "不用假数字，用真实剩余名额",
            ],
            "materials": [
                "【群满进度卡】竖屏：群二维码 + 真实剩余名额",
            ],
            "ai": [
                "让AI写『群快满』真实稀缺钩子（不假紧迫）",
            ],
        },
        {
            "control": "可控",
            "note": "通过好友后第一时间拉群，用『进群解锁阶梯礼包』分步钩子。",
            "sop": [
                "入群引导讲清『进群领资料→看1讲→答1题→领学情报告』四步礼包",
                "对24h内未进群轻提醒1次",
                "礼包每步都是真实价值，不画饼",
            ],
            "materials": [
                "【阶梯礼包卡】竖屏：四步解锁路径 + 群二维码",
            ],
            "ai": [
                "让AI写阶梯礼包钩子（每步真实可得）",
            ],
        },
    ],
    "APP下载率": [
        {
            "control": "可控",
            "note": "下载靠讲清必要性 + 简单提醒，不高频催（防投诉）。APP=上课 + 课堂互动答题，作业发群里、外包AI批改，不在APP。",
            "sop": [
                "加好友后10分钟内私信发下载引导（APP图标 + 用途）",
                "当天晚8点在班级群发1次下载引导图文（置顶）",
                "第2天上午对未下载家长私信补1次，换角度：明天D1要在APP上课 + 参加课堂互动答题",
                "全期提醒不超过2次",
                "下载引导卡分两版钩子：开课前→『上课 + 课堂互动答题都在这里』；开课后→『孩子课堂答题表现已生成，打开APP查看』",
            ],
            "materials": [
                "【下载引导私信配图】竖屏1080×1920：APP图标 + 上课/课堂互动答题都在这里",
                "【群内下载长图】应用商店搜→下载→注册→进班级 4步截图",
                "【孩子课堂答题表现卡】1张真实答题情况 + 一句家长感言",
            ],
            "ai": [
                "把下载入口变成『看孩子课堂答题表现』：让AI用本期过程数据给每位家长生成『XX的课堂答题表现已生成，打开APP查看』诱饵卡",
                "⚠️ 全期提醒≤2次，多了=投诉/封号；让AI把『要你下载』改写成『送你价值』，靠单次高价值触达",
                "用孩子真实课堂表现（如『本周答对18题、互动活跃』）做下载诱饵",
            ],
        },
        {
            "control": "可控",
            "note": "下载靠轻同伴压力 + 真实反馈，仍≤2次提醒（防投诉）。APP=上课 + 课堂互动答题。",
            "sop": [
                "在班级群晒『已有XX%家长下载并进APP上课答题』进度，制造轻同伴压力",
                "晚8点群内发1次，带1条已下载家长的真实反馈",
                "第2天上午对未下载家长私信补1次，全期≤2次",
            ],
            "materials": [
                "【下载引导图】突出『班里大部分人都下了』",
                "【已下载家长感言卡】1条真实感言",
            ],
            "ai": [
                "让AI汇总已下载家长的真实好评，做成引导文案（『班里XX位家长已下载，反馈都说…』）",
                "用『大部分人都下了』轻促，但仍≤2次，不骚扰",
                "⚠️ 全期提醒≤2次，多了=投诉/封号",
            ],
        },
        {
            "control": "可控",
            "note": "下载绑『上课 + 课堂互动答题刚需』，价值导向，≤2次提醒。",
            "sop": [
                "下载引导直接绑刚需：『明天D1上课 + 课堂互动答题都在APP，不下就参与不了』",
                "开课前1天发引导，第2天上午补1次强调刚需，全期≤2次",
                "下载引导图突出『上课 + 课堂互动答题都在APP』",
            ],
            "materials": [
                "【下载引导图】突出『不下APP=孩子上不了课、答不了互动题』",
            ],
            "ai": [
                "让AI写『不下APP=孩子上不了课/答不了题』的刚需提醒（价值导向，不是硬催）",
                "⚠️ 全期提醒≤2次，多了=投诉/封号；靠『刚需』一次说清就够",
                "用孩子真实课堂表现做诱饵卡",
            ],
        },
        {
            "control": "可控",
            "note": "下载绑『直播间/课前置』价值，≤2次提醒。",
            "sop": [
                "下载引导绑『D3直播间要用的APP功能/抽奖』，开课前1天发",
                "当天晚8点群内1次，第2天上午补1次，全期≤2次",
                "下载引导图突出『直播间抽奖需APP参与』",
            ],
            "materials": [
                "【下载引导图(直播间版)】突出『直播间抽奖/互动需APP』",
            ],
            "ai": [
                "让AI写『直播间抽奖需APP参与』的钩子（价值导向，不是硬催）",
                "⚠️ 全期提醒≤2次，多了=投诉/封号",
            ],
        },
        {
            "control": "可控",
            "note": "下载绑『孩子课堂答题表现』可视化价值，≤2次提醒。",
            "sop": [
                "下载引导突出『每周自动生成孩子课堂答题表现』",
                "加好友后10分钟内私信发引导，晚8点群内1次，次日补1次，全期≤2次",
                "下载卡突出『不下APP看不到孩子课堂答题表现』",
            ],
            "materials": [
                "【下载引导图(答题表现版)】突出『孩子课堂答题表现』",
                "【孩子课堂答题表现卡】1张真实答题情况 + 一句家长感言",
            ],
            "ai": [
                "让AI把APP价值包装成『孩子课堂答题表现』，用真实表现做诱饵",
                "⚠️ 全期提醒≤2次，多了=投诉/封号",
            ],
        },
        {
            "control": "可控",
            "note": "下载绑『孩子学情报告/错题本』刚需，价值导向，≤2次提醒。",
            "sop": [
                "下载引导突出『每周自动生成孩子错题本/学情报告，只在APP看』",
                "加好友后10分钟私信，晚8点群内1次，次日补1次，全期≤2次",
                "下载卡突出『不下APP看不到孩子学情』",
            ],
            "materials": [
                "【下载引导图(学情版)】突出『孩子错题本/学情报告』",
            ],
            "ai": [
                "让AI把APP价值包装成『孩子学情报告』，用真实表现做诱饵",
                "⚠️ 全期提醒≤2次，多了=投诉/封号",
            ],
        },
        {
            "control": "可控",
            "note": "下载靠降低操作摩擦，扫码/短链一步直达，仍≤2次提醒。",
            "sop": [
                "下载引导用短链/扫码直达应用商店搜索结果，省去手动搜",
                "加好友后10分钟私信发短链，晚8点群内1次，次日补1次，全期≤2次",
                "引导图突出『点一下就到下载页』",
            ],
            "materials": [
                "【一键下载卡】扫码/短链直达 + APP图标",
            ],
            "ai": [
                "让AI把下载步骤压成『点一下』，降低操作门槛",
                "⚠️ 全期提醒≤2次，多了=投诉/封号",
            ],
        },
        {
            "control": "可控",
            "note": "下载靠『首登即生成孩子专属开课提醒』价值，≤2次提醒。",
            "sop": [
                "下载引导突出『首次打开APP即生成孩子专属开课提醒，到点提醒你』",
                "加好友后10分钟私信，晚8点群内1次，次日补1次，全期≤2次",
                "首登价值真实，不夸大",
            ],
            "materials": [
                "【首登价值卡】突出『首次打开即领孩子专属开课提醒』",
            ],
            "ai": [
                "让AI把首登包装成『孩子专属开课闹钟』，用真实价值做诱饵",
                "⚠️ 全期提醒≤2次，多了=投诉/封号",
            ],
        },
    ],
    "开营回复率": [
        {
            "control": "可控",
            "note": "开营触达讲清时间 + 用处，未回复简单提醒。",
            "sop": [
                "开营前1天发开营通知（时间/安排/家长为什么要回复）",
                "开营当天早上再提醒1次",
                "对未回复家长私信补1次，不反复@",
                "开营通知按家长画像分两版：放养型用『孩子被点名表扬』钩子；观望型用『往期开营后孩子变化』钩子",
            ],
            "materials": [
                "【开营通知卡】横版：开营时间 + 3个家长能得到什么 + 按钮式回复收到",
            ],
            "ai": [
                "沉默家长『一次精准唤醒』：让AI把已到课未回复的家长单列一类，只发1条『孩子被点名表扬了，回复领专属资料』，不群发、不反复@",
                "让AI把开营通知从『广播』改造成『1个让家长忍不住回的提问』（如『孩子目前最怕的数学点是哪个？回复我帮你把关』）",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达用『孩子被点名』钩子，未回复简单提醒。",
            "sop": [
                "开营通知突出『孩子会被点名 / 被表扬』，放养型家长最吃这套",
                "开营前1天发，当天早提醒1次，未回复补1次",
                "用孩子相关提问代替广播通知",
            ],
            "materials": [
                "【开营通知卡(孩子视角)】横版：『开营XX会被点名讲XX』+ 回复收到按钮",
            ],
            "ai": [
                "让AI给每个孩子生成『开营会被点名讲XX』的个性化提醒，家长为孩子愿意回",
                "用孩子相关提问代替广播（如『孩子目前最怕的数学点是哪个？回复我帮你把关』）",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达用『往期变化』背书，未回复简单提醒。",
            "sop": [
                "开营通知展示往期开营后孩子的真实变化（进步案例），观望型家长最吃这套",
                "开营前1天发，当天早提醒1次，未回复补1次",
                "用『别人家孩子开营后怎样』建立预期",
            ],
            "materials": [
                "【往期开营变化案例卡】1个真实前后对比 + 一句家长感言",
            ],
            "ai": [
                "让AI挑往期孩子开营前后的真实变化做素材，建立『来了就有用』的预期",
                "用『别人家孩子开营后怎样』轻背书，代替干巴巴通知",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达把回复门槛降到最低，未回复简单提醒。",
            "sop": [
                "开营通知只求『回复1』最低门槛，用按钮式回复",
                "开营前1天发，当天早提醒1次，未回复补1次",
                "用『扣1即可』降低家长动作成本",
            ],
            "materials": [
                "【开营通知卡】横版：开营时间 + 按钮式『回复1收到』",
            ],
            "ai": [
                "让AI把回复门槛降到『扣1即可』，降低家长动作成本",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达带班主任温度人设，未回复简单提醒。",
            "sop": [
                "开营通知带『我是XX老师，这期专门盯孩子』人设温度",
                "开营前1天发，当天早提醒1次，未回复补1次",
                "放养型家长最吃『有人专门管孩子』这套",
            ],
            "materials": [
                "【开营通知卡(老师视角)】横版：『我是XX老师，这期盯孩子』+ 回复收到",
            ],
            "ai": [
                "让AI写带温度的1对1口吻开营私信（不是群发广播）",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达绑『开营即领资料』限时钩子，未回复简单提醒。",
            "sop": [
                "开营通知突出『开营当天领专属资料包』，制造轻紧迫",
                "开营前1天发，当天早提醒1次，未回复补1次",
                "用『领资料』降低回复心理门槛",
            ],
            "materials": [
                "【开营通知卡(资料钩子)】横版：开营领专属资料 + 回复收到",
            ],
            "ai": [
                "让AI写『开营即领资料』钩子，降低回复门槛",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达带「孩子专属学习画像预告」，未回复简单提醒。",
            "sop": [
                "开营通知突出『回复即领孩子专属学习画像（薄弱点+建议）』",
                "开营前1天发，当天早提醒1次，未回复补1次",
                "用价值降低回复门槛",
            ],
            "materials": [
                "【学习画像预告卡】横版：回复领孩子专属薄弱点+建议",
            ],
            "ai": [
                "让AI给每个孩子生成『专属学习画像预告』（薄弱点+一句建议）",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达用「开营小福利」低门槛钩子，未回复简单提醒。",
            "sop": [
                "开营通知突出『开营当天抽X名送资料/小课』，回复参与",
                "开营前1天发，当天早提醒1次，未回复补1次",
                "福利真实不夸大",
            ],
            "materials": [
                "【开营福利卡】横版：开营抽奖 + 回复参与",
            ],
            "ai": [
                "让AI写开营福利钩子（真实不夸大）",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达先1对1私信预热再发群，私信开个头家长更愿回。",
            "sop": [
                "开营前1天先给重点家长私信预热（『这期孩子这块弱，开营来讲』）",
                "再发群通知，私信过的家长更易回",
                "未回复补1次，不反复@",
            ],
            "materials": [
                "【私信预热卡】1对1：孩子弱项 + 开营时间",
            ],
            "ai": [
                "让AI给重点家长写1对1私信预热（按孩子薄弱点）",
            ],
        },
        {
            "control": "可控",
            "note": "开营触达把回复变成『参与调研』，不是干收到。",
            "sop": [
                "开营通知末尾加『孩子最想提升哪块？回复1计算/2应用/3概念』",
                "把干回『收到』变成选弱项，家长更愿意动",
                "开营前1天发，当天早提醒1次，未回复补1次",
            ],
            "materials": [
                "【开营调研卡】横版：开营时间 + 选孩子弱项回复",
            ],
            "ai": [
                "让AI把开营通知末尾改造成『选孩子弱项』调研，提升回复率",
            ],
        },
    ],
    "Day1到课率": [
        {
            "control": "可控",
            "note": "提醒可多次，但每次带内容价值，靠内容吸引到课。",
            "sop": [
                "开课前1天发明天讲什么 + 孩子能收获什么预热",
                "开课前1小时发上课链接 + 今天重点提醒",
                "当天中午发回放/作业提醒",
                "对未到课家长私信1对1确认，不发群发干巴巴催课",
                "到课后立刻发『今日作业即打卡』提醒，形成正反馈闭环",
            ],
            "materials": [
                "【到课预热海报】横版：课程主题 + 3个收获点 + 时间",
                "【开课提醒卡】竖屏：今天X点 + 上课链接按钮",
            ],
            "ai": [
                "AI『错题归因弹药』：作业是AI批改的，AI能汇总全班最易错3个知识点，写成『今晚直播重点讲XX，孩子刚好这块弱』的预告",
                "让AI给每个孩子生成『课前1分钟预告卡』（今天讲什么+孩子能收获什么+一个小钩子），转发到群",
                "用平台预测功能提前看本期哪类孩子最可能不到课，课前1小时发『孩子专属薄弱点预告』",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠『孩子专属预告』吸引到课。",
            "sop": [
                "课前1小时发『孩子今天要学XX（他刚好弱的）』个人化预告",
                "开课前1天预热，当天中午回放/作业提醒",
                "未到课家长私信1对1确认",
            ],
            "materials": [
                "【课前预告卡(个人化)】竖屏：今天讲XX + 孩子这块弱 + 上课链接",
            ],
            "ai": [
                "按孩子薄弱点（来自AI批改）生成个人化预告卡，转发到群——孩子催爸妈来",
                "用平台预测看掉队人群，课前1小时精准发",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠『到课→打卡』正反馈闭环吸引。",
            "sop": [
                "到课后立刻发『今日作业即打卡』提醒，形成正反馈（作业完成=打卡，家长有成就感）",
                "开课前1天预热，课前1小时提醒，中午回放/作业",
                "未到课家长私信1对1确认",
            ],
            "materials": [
                "【到课预热海报】横版：课程主题 + 3个收获点 + 时间",
                "【打卡成就卡】『今日打卡完成』小奖状风格",
            ],
            "ai": [
                "让AI把『到课→作业→打卡』做成一条龙提醒模板，自动触发",
                "用『打卡成就』激励下期继续来（正反馈闭环）",
                "错题归因弹药：用AI批改汇总易错点写预告",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠『同伴榜样』吸引到课。",
            "sop": [
                "课前发『上期XX同学用这招提了X分』同伴榜样",
                "开课前1天预热，课前1小时提醒，中午回放/作业",
                "未到课家长私信1对1确认",
            ],
            "materials": [
                "【同伴榜样卡】往期进步最大同学案例 + 上课链接",
            ],
            "ai": [
                "让AI挑往期进步最大的同学做榜样素材（轻同伴压力，不点名压力）",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠『课前悬念』吸引到课。",
            "sop": [
                "课前发『今天有个让孩子惊呼的小实验』悬念钩子",
                "开课前1小时发上课链接 + 悬念预告",
                "未到课家长私信1对1确认",
            ],
            "materials": [
                "【课前悬念卡】竖屏：『今天直播有个XX彩蛋』+ 上课链接",
            ],
            "ai": [
                "让AI写『今天直播有个XX彩蛋/小实验』悬念预告，孩子催爸妈来",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠『孩子到课成就』正反馈吸引。",
            "sop": [
                "课前发『上期你孩子已集齐X枚到课勋章』成就提醒",
                "开课前1天预热，课前1小时提醒，中午回放/作业",
                "未到课家长私信1对1确认",
            ],
            "materials": [
                "【到课成就卡】『孩子已集齐X枚到课勋章』小奖状",
            ],
            "ai": [
                "让AI把到课做成『集勋章』游戏化，孩子催爸妈来",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠「家长社群打卡接龙」氛围吸引到课。",
            "sop": [
                "课前发『上期打卡接龙XX人参与』氛围",
                "开课前1天预热，课前1小时提醒，中午回放/作业",
                "未到课家长私信1对1确认",
            ],
            "materials": [
                "【打卡接龙氛围卡】往期接龙截图 + 上课链接",
            ],
            "ai": [
                "让AI生成接龙氛围文案（轻同伴压力，不点名）",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠「课前1对1关怀短消息」高触达吸引到课。",
            "sop": [
                "课前1小时给未互动家长发1对1关怀（『孩子今天讲XX，他这块弱，建议来』）",
                "开课前1天预热，中午回放/作业",
                "私信不群发，不反复@",
            ],
            "materials": [
                "【1对1关怀卡】竖屏：孩子弱项 + 上课链接",
            ],
            "ai": [
                "让AI生成个性化关怀短消息（按孩子薄弱点）",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，靠『今日课衔接昨日作业』连续性钩子吸引到课。",
            "sop": [
                "课前发『今天讲XX，正好接着你昨天作业错的那个点』",
                "开课前1天预热，课前1小时提醒，中午回放/作业",
                "未到课家长私信1对1确认",
            ],
            "materials": [
                "【衔接预告卡】竖屏：今日课 + 昨日作业衔接点 + 链接",
            ],
            "ai": [
                "让AI把今日课和昨日作业错点串成『承上启下』预告",
            ],
        },
        {
            "control": "可控",
            "note": "提醒可多次，对『往期到课高、本期掉』的家长单独1对1召回。",
            "sop": [
                "筛出往期到课好但本期没动静的家长，课前1小时1对1唤醒",
                "话术：『孩子上期课都来了，这期讲XX他刚好弱，别落下』",
                "私信不群发，不反复@",
            ],
            "materials": [
                "【流失召回卡】1对1：往期到课好 + 本期弱项 + 链接",
            ],
            "ai": [
                "让AI生成流失召回1对1话术（点名往期好、本期别落）",
            ],
        },
    ],
}

# 内部示范家长（托号）打法：在群里用内部号演真实家长带节奏，降低真实家长行动门槛。
# 字眼用「内部示范家长」，标清「人演 / AI帮」；不涉及销售话术、不替代真实互动。
SEED_LIB = {
    "好友率": "好友率由前端BD决定（家长主动加你），不适用内部示范家长带节奏——不能fake真实加微。把精力放下游环节。",
    "入群率": "内部示范家长『晨晨妈』通过好友后半天发：『老师拉群没？我想进群看资料，怕错过』→真实家长见有人主动想进，跟着进；进群后补：『资料收到了，正好孩子乘法分配律总错，这有个专项练』证明群真有料。AI帮：生成内部号人设卡(焦虑/放养/观望三类)+话术3版+排期表，你照演，错峰带瑕疵像真人。",
    "APP下载率": "内部示范家长『桐桐爸』晚8点发：『刚下了APP，陪娃做了几道互动题，他居然说比做题册有意思😂』；『米粒妈』隔10分钟问：『那个答题排行榜是实时的不？想看看娃排第几』→引发真实家长好奇去搜(口播应用商店搜，不印码)。AI帮：生成不同人设话术多版+排期，避免复制粘贴感。",
    "开营回复率": "内部示范家长『轩轩妈』开营通知一发立刻回：『收到！孩子说想上，他数学一直没信心』；『兜兜爸』问：『明天讲的计算是几年级难度？想提前给他预热』→把广播变讨论，真实家长才接话。AI帮：生成不同人设回复话术多版。",
    "Day1到课率": "内部示范家长课前发：『孩子说今天讲小数，他这块弱，坐等开课』；课中发：『刚答了第3题，孩子说比学校讲得透』；课后发：『作业发群里了，我先对下娃写得对不对』→带真实家长去看群里的作业。AI帮：按人设生成课前/课中/课后话术多版+排期。",
}


# 卡点筛选 SOP：按"家长卡在7环漏斗哪一环"筛选人群，缺哪补哪，避免一刀切群发、守住克制触达约束。
# 挂到对应环节；好友率上游不可控不参与筛选。
FILTER_LIB = {
    "入群率": "卡点A【加了好友没进群】：企微/表格筛出『已通过好友但不在群里』的家长 → 私信补发入群邀请(群二维码)并讲清群价值(资料/答疑/直播提醒)；24h内未进群轻提醒1次，不反复@。",
    "APP下载率": "卡点B【进了群没下载】：筛出『在群但APP无孩子账号/无答题记录』的家长 → 换『上课+课堂互动答题』钩子发1次；全期≤2次提醒(守红线)；用内部示范家长演『下了答题有意思』带节奏。",
    "开营回复率": "卡点C【下载了没回开营】：筛出『APP有账号但未回开营通知』的家长 → 私信1次换角度(『孩子被点名表扬了，回复领专属资料』)；不群发催、不反复@。",
    "Day1到课率": "卡点D【回了开营没到课】：筛出『回复开营但Day1无孩子答题记录』的家长 → 课前1小时发『娃专属薄弱点预告卡』(来自课堂答题表现)；内部示范家长课前带『坐等开课』氛围。",
}

# 通用筛选维度方向（简单提及，不展开）：画像/学情/响应行为/转化意向，供运营叠加使用。
FILTER_DIMS = "①家长画像：按焦虑/放养/观望分群，同一动作换不同话术；②孩子学情：按薄弱点(计算/应用/概念)分组发预告卡；③响应行为：按活跃/潜水/已读不回分层——潜水靠价值内容钩、已读不回私信换角度1次；④转化意向：直播间后筛主动问课者重点跟(你个人业务，平台不代写话术)。"

def _period_seq(period):
    """期次在全部期次中的顺序序号（数值排序），用于稳定轮转。"""
    import re, sqlite3
    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.execute("SELECT DISTINCT period FROM records ORDER BY period")
        periods = [str(r[0]) for r in cur.fetchall()]
        conn.close()
        if str(period) in periods:
            return periods.index(str(period))
    except Exception:
        pass
    m = re.search(r"\d+", str(period))
    return int(m.group()) if m else 0


def _name_key(name):
    """人名稳定哈希 -> 非负整数，用于按人分流（方案1：保证不同人不重）。"""
    import hashlib
    if not name:
        return 0
    return int(hashlib.md5(str(name).encode("utf-8")).hexdigest(), 16)


def _reason_key(cur, base):
    """按本期值档位 + 环比方向给原因码(0..8)；数据不足返回 None（方案3：按原因分流）。"""
    if cur is None or base is None:
        return None
    band = 0 if cur < 60 else (1 if cur < 85 else 2)        # 低 / 中 / 高
    delta = cur - base
    direction = 0 if delta <= -3 else (2 if delta >= 3 else 1)  # 降 / 平 / 升
    return band * 3 + direction


def get_advice(var, period="", name="", grade="", subject="", cur_val=None, base_val=None):
    """方案3+方案1组合选建议变体：
    - 方案3 按原因分流：能算原因(档位+环比)时原因参与轮转，同类原因给同类建议
    - 方案1 按人名取模：人名稳定参与，保证不同人看到的不重
    - 保留期次顺序轮转：同人同期固定、相邻期次必换、团队模式(无人名)退回纯期次轮转
    """
    entry = ADVICE_LIB.get(var)
    if not entry:
        return {}
    if isinstance(entry, dict):
        return entry  # 兼容旧结构
    if not isinstance(entry, list) or len(entry) == 0:
        return {}
    n = len(entry)
    p_idx = _period_seq(period)
    name_key = _name_key(name)
    # 团队模式（无人名）不按原因分流，退回纯期次轮转，保持原行为
    reason_key = _reason_key(cur_val, base_val) if name else None

    base = p_idx
    if reason_key is not None:
        base = base * 31 + reason_key
    base = base * 31 + name_key
    idx = base % n

    v = dict(entry[idx])
    seed = SEED_LIB.get(var, "")
    if seed:
        v["seed"] = seed
    flt = FILTER_LIB.get(var, "")
    if flt:
        v["filter"] = flt
    v["filter_more"] = FILTER_DIMS
    return v


def generate_suggestions(df, analysis_results, grade="", period=""):
    """基于分析结果生成优化建议，返回优先级排序的建议列表"""
    suggestions = []
    desc = analysis_results.get("descriptive", {})
    final_reg = analysis_results.get("final", [])
    path_analysis = analysis_results.get("path", [])
    
    # 构建系数查找表
    coef_map = {r["variable"]: r for r in final_reg}
    path_map = {r["variable"]: r for r in path_analysis}
    
    y_mean = desc.get("最终转化率", {}).get("mean", 0)
    day3_mean = desc.get("Day3出直播间转化率", {}).get("mean", 0)
    
    # ── 1. 计算每个变量的改进空间 × 影响力 ──
    items = []
    for x_var in ["好友率", "入群率", "APP下载率", "开营回复率", "Day1到课率"]:
        var_desc = desc.get(x_var, {})
        current_mean = var_desc.get("mean", 0)
        coef_info = coef_map.get(x_var, {})
        coef = abs(coef_info.get("coefficient", 0))
        p_val = coef_info.get("p_value", 1)
        significant = coef_info.get("significant", False)
        
        # 改进空间：理论上限100% - 当前均值
        room = max(0, 100 - current_mean)
        # 每提升10个百分点带来的转化率增益
        gain_per_10pct = coef * 10
        # 优先级分 = 改进空间 × 影响力（归一化）
        priority = room * coef
        
        items.append({
            "variable": x_var,
            "current_mean": round(current_mean, 1),
            "room": round(room, 1),
            "coefficient": round(coef, 4),
            "gain_per_10pct": round(gain_per_10pct, 2),
            "significant": significant,
            "p_value": round(p_val, 4),
            "priority": round(priority, 2),
        })
    
    # 按优先级排序
    items.sort(key=lambda x: x["priority"], reverse=True)
    
    # ── 2. 生成自然语言建议 ──
    rank_labels = ["第一优先", "第二优先", "第三优先", "第四优先", "第五优先"]
    variable_labels = {
        "Day1到课率": "Day1到课率",
        "APP下载率": "APP下载率",
        "开营回复率": "开营回复率",
        "入群率": "入群率",
        "好友率": "好友率",
    }
    # 建议使用全局 ADVICE_LIB（模块级结构化方案）
    
    for i, item in enumerate(items):
        var = item["variable"]
        label = variable_labels.get(var, var)
        rank = rank_labels[i] if i < len(rank_labels) else f"第{i+1}优先"
        
        # 判断当前水平
        if item["current_mean"] >= 85:
            level = "已处于高位"
            level_note = "提升空间有限，保持即可"
        elif item["current_mean"] >= 60:
            level = "中等水平"
            level_note = f"仍有 {item['room']:.0f}% 的提升空间"
        else:
            level = "偏低"
            level_note = f"有 {item['room']:.0f}% 的巨大提升空间，是当前最大短板"
        
        # 预期效果
        if item["significant"] and item["coefficient"] > 0:
            effect = f"每提升10个百分点，预计转化率增加 {item['gain_per_10pct']:.2f}%"
            confidence = "统计显著，可信度高"
        elif item["coefficient"] > 0:
            effect = f"每提升10个百分点，预计转化率增加 {item['gain_per_10pct']:.2f}%"
            confidence = "统计不显著（p={:.3f}），趋势参考".format(item["p_value"])
        else:
            effect = "与转化率关联较弱"
            confidence = "不建议作为优先发力点"
        
        # 瓶颈分析：如果在路径分析中
        path_info = path_map.get(var, {})
        bottleneck_note = ""
        if var != "Day1到课率" and path_info.get("indirect_effect", 0) > 0:
            indirect = path_info.get("indirect_effect", 0)
            total = path_info.get("total_effect", 0)
            if total > indirect * 1.5:
                bottleneck_note = f"该环节通过影响到课率间接拉动转化（间接效应 {indirect:.4f}），是漏斗上游关键节点"
        
        suggestion = {
            "rank": rank,
            "rank_index": i + 1,
            "variable": var,
            "current_level": level,
            "level_note": level_note,
            "effect": effect,
            "confidence": confidence,
            "action": get_advice(var, period),
            "bottleneck_note": bottleneck_note,
            "current_mean": item["current_mean"],
            "room": item["room"],
            "gain_per_10pct": item["gain_per_10pct"],
            "significant": item["significant"],
        }
        suggestions.append(suggestion)
    
    # ── 3. 场景模拟：Top-1 变量提升后的预期 ──
    if items:
        top = items[0]
        if top["coefficient"] > 0:
            # 模拟提升5%、10%、15%的效果
            scenarios = []
            for pct in [5, 10, 15]:
                new_val = min(top["current_mean"] + pct, 100)
                actual_gain = pct * top["coefficient"]
                new_conversion = y_mean + actual_gain
                scenarios.append({
                    "improve_by": pct,
                    "new_value": round(new_val, 1),
                    "predicted_conversion_gain": round(actual_gain, 2),
                    "predicted_conversion": round(new_conversion, 2),
                })
            
            top_suggestion_extra = {
                "scenarios": scenarios,
                "summary": f"当前{top['variable']}为 {top['current_mean']:.1f}%，如果能提升至 {min(top['current_mean'] + 10, 100):.1f}%，预计转化率可从 {y_mean:.2f}% 提升至 {y_mean + top['gain_per_10pct']:.2f}%",
            }
        else:
            top_suggestion_extra = None
    else:
        top_suggestion_extra = None
    
    # ── 4. 漏斗健康度诊断 ──
    funnel_health = []
    # 转化漏斗：好友率 → 入群率 → APP下载率 → Day1到课率 → 转化率
    funnel_steps = [
        ("好友率", "入群率", "加好友 → 入群"),
        ("入群率", "APP下载率", "入群 → APP下载"),
        ("APP下载率", "Day1到课率", "APP下载 → Day1到课"),
        ("Day1到课率", "最终转化率", "Day1到课 → 最终转化"),
    ]
    for from_var, to_var, label in funnel_steps:
        from_val = desc.get(from_var, {}).get("mean", 0)
        to_val = desc.get(to_var, {}).get("mean", 0)
        drop = from_val - to_val
        if drop > 20:
            status = "严重流失"
            advice = f"从{from_var}到{to_var}流失 {drop:.1f}%，此处是最大断崖，优先修补"
        elif drop > 10:
            status = "注意"
            advice = f"从{from_var}到{to_var}流失 {drop:.1f}%，有优化空间"
        else:
            status = "健康"
            advice = f"转化衔接良好，流失仅 {drop:.1f}%"
        funnel_health.append({
            "step": label,
            "from_var": from_var,
            "to_var": to_var,
            "from_value": round(from_val, 1),
            "to_value": round(to_val, 1),
            "drop": round(drop, 1),
            "status": status,
            "advice": advice,
        })
    
    return {
        "suggestions": suggestions,
        "top_action": top_suggestion_extra,
        "funnel_health": funnel_health,
        "grade": grade or "全年级",
        "y_mean": round(y_mean, 2),
        "day3_mean": round(day3_mean, 2),
    }


# ── 路由 ──────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    """接收CSV或Excel文件 → 存入数据库 → 用全量数据做分析"""
    if "file" not in request.files:
        return jsonify({"error": "请上传CSV或Excel文件"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "文件名为空"}), 400
    
    try:
        content = file.read()
        fname = file.filename.lower()
        
        # Excel (.xlsx / .xls)
        if fname.endswith(('.xlsx', '.xls')):
            import openpyxl
            df = pd.read_excel(io.BytesIO(content), engine='openpyxl')
        else:
            # CSV
            for encoding in ["utf-8", "gbk", "gb2312", "utf-8-sig"]:
                try:
                    df = pd.read_csv(io.BytesIO(content), encoding=encoding)
                    break
                except:
                    continue
            else:
                return jsonify({"error": "无法解析CSV文件，请检查编码格式"}), 400
        
        # 存入数据库（所有列一次传齐，覆盖式写入）
        required_cols = ["期次", "好友率", "入群率", "APP下载率", "开营回复率", "Day1到课率", "最终转化率"]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            found = [c for c in df.columns]
            return jsonify({
                "error": f"还缺这几列：{', '.join(missing)}。"
                         f"你表格里我认出来的列有：{', '.join(found) if found else '（没认出任何列）'}。"
                         f"把缺的列补上，或把你的表头第一行发我，我帮你对齐。"
            }), 400
        
        grade = request.form.get("grade", "")
        subject = request.form.get("subject", "数学")
        conflicts = check_grade_conflicts(df.to_dict("records"), grade, subject)
        if conflicts:
            return jsonify({"error": "⚠️ 已阻止本次录入（发现年级冲突）：\n" + "\n".join(conflicts)}), 400
        rows_saved = save_to_db(df.to_dict("records"), source=fname, grade=grade, subject=subject)
        
        # 用全量数据库数据分析（默认按上传的年级/学科过滤，空表示全年级/全部）
        return _analyze_from_db(rows_saved, grade=grade if grade else None, subject=subject)
    except Exception as e:
        return jsonify({"error": f"分析出错: {str(e)}"}), 500


@app.route("/analyze-json", methods=["POST"])
def analyze_json():
    """接收前端传来的 JSON 数据（手动录入/Excel导出）→ 存入数据库 → 用全量数据分析"""
    data = request.get_json()
    if not data or "rows" not in data:
        return jsonify({"error": "请提供 rows 数据"}), 400
    
    try:
        df = pd.DataFrame(data["rows"])
        # 确保数值列是数字类型
        for col in ["好友率", "入群率", "APP下载率", "开营回复率", "Day1到课率", "最终转化率"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        if "Day3出直播间转化率" not in df.columns:
            df["Day3出直播间转化率"] = np.nan
        else:
            df["Day3出直播间转化率"] = pd.to_numeric(df["Day3出直播间转化率"], errors='coerce')
        
        # 存入数据库
        grade = data.get("grade", "")
        subject = data.get("subject", "数学")
        conflicts = check_grade_conflicts(data["rows"], grade, subject)
        if conflicts:
            return jsonify({"error": "⚠️ 已阻止本次录入（发现年级冲突）：\n" + "\n".join(conflicts)}), 400
        rows_saved = save_to_db(df.to_dict("records"), source="manual-json", grade=grade, subject=subject)
        
        # 用全量数据库数据分析
        return _analyze_from_db(rows_saved, grade=grade if grade else None, subject=subject)
    except Exception as e:
        return jsonify({"error": f"分析出错: {str(e)}"}), 500


@app.route("/db-stats")
def db_stats():
    """返回数据库统计信息"""
    return jsonify(get_db_stats())


@app.route("/analyze-grade", methods=["POST"])
def analyze_grade():
    """按指定年级/学科重新分析"""
    data = request.get_json() or {}
    grade = data.get("grade", "")
    subject = data.get("subject", "")
    return _analyze_from_db(0, grade=grade if grade else None, subject=subject if subject else None)


def _analyze_from_db(rows_saved=0, grade=None, subject=None):
    """从数据库读取全部数据进行分析，可按年级/学科筛选"""
    df = get_all_data(grade=grade, subject=subject)
    if df is None:
        return jsonify({"error": f"数据库中暂无{('（' + subject + '）') if subject else ''}{grade + '年级' if grade else ''}数据"}), 400
    
    results = run_analysis(df)
    results["rows"] = len(df)
    results["batches"] = int(df["期次"].nunique())
    results["just_saved"] = rows_saved
    results["current_grade"] = grade or "全年级"
    results["current_subject"] = subject or "全部"
    results["db_stats"] = get_db_stats(grade=grade, subject=subject)
    return jsonify(results)


@app.route("/suggestions", methods=["POST"])
def suggestions():
    """基于当前年级/学科的全量数据生成AI优化建议"""
    data = request.get_json() or {}
    grade = data.get("grade", "")
    subject = data.get("subject", "")
    df = get_all_data(grade=grade if grade else None, subject=subject if subject else None)
    if df is None:
        return jsonify({"error": f"数据库中暂无{('（' + subject + '）') if subject else ''}{grade + '年级' if grade else ''}数据"}), 400
    
    analysis_results = run_analysis(df)
    suggestions = generate_suggestions(df, analysis_results, grade=grade)
    suggestions["db_stats"] = get_db_stats(grade=grade if grade else None, subject=subject if subject else None)
    suggestions["rows"] = len(df)
    return jsonify(suggestions)


@app.route("/names")
def get_names():
    """返回数据库中所有人员姓名列表及各自的可用期次（可按年级/学科筛选）"""
    grade = request.args.get("grade", "")
    subject = request.args.get("subject", "")
    conn = get_db()
    where = []
    params = []
    if grade:
        where.append("grade = ?"); params.append(grade)
    if subject:
        where.append("subject = ?"); params.append(subject)
    wsql = ("WHERE " + " AND ".join(where)) if where else ""
    rows = conn.execute(
        "SELECT DISTINCT name, period FROM records " + wsql + " ORDER BY name, period", params
    ).fetchall()
    conn.close()
    
    # 按姓名聚合期次
    name_periods = {}
    for r in rows:
        name_periods.setdefault(r["name"], []).append(r["period"])
    
    return jsonify({
        "names": list(name_periods.keys()),
        "periods": name_periods,
    })


@app.route("/personal-suggestions", methods=["POST"])
def personal_suggestions():
    """
    个人诊断：选人+选期次 → 该人在当期 vs 同组同期其他人对比
    """
    data = request.get_json() or {}
    name = data.get("name", "").strip()
    period = _normalize_period(data.get("period", ""))
    grade = data.get("grade", "")
    subject = data.get("subject", "")
    
    if not name:
        return jsonify({"error": "请输入姓名"}), 400
    if not period:
        return jsonify({"error": "请选择期次"}), 400
    
    df = get_all_data(grade=grade if grade else None, subject=subject if subject else None)
    if df is None:
        return jsonify({"error": "数据库中暂无数据"}), 400
    
    # 只在同一期次内做对比
    same_period = df[df["期次"].astype(str) == str(period)]
    if len(same_period) == 0:
        return jsonify({"error": f"未找到期次「{period}」的数据"}), 400

    personal = same_period[same_period["姓名"] == name]
    if len(personal) == 0:
        return jsonify({"error": f"「{name}」在「{period}」期中没有数据"}), 400

    # 同组 = 同一期次 + 同学科 + 同年级（用这个人自己的年级/学科锁范围）
    others = same_period[same_period["姓名"] != name]
    p_row = personal.iloc[0]
    p_grade = str(p_row.get("年级", "") or "").strip()
    p_subject = str(p_row.get("学科", "") or "").strip()
    if p_grade:
        others = others[others["年级"].astype(str).str.strip() == p_grade]
    if p_subject:
        others = others[others["学科"].astype(str).str.strip() == p_subject]

    if len(others) == 0:
        return jsonify({"error": f"「{period}」期只有你一个人的数据，无法做同期对比"}), 400
    
    x_vars = ["好友率", "入群率", "APP下载率", "开营回复率", "Day1到课率"]
    y_vars = ["最终转化率", "Day3出直播间转化率"]
    
    # 同组同期统计（排除自己）
    others_stats = {}
    for col in x_vars + y_vars:
        vals = others[col].dropna()
        others_stats[col] = {
            "mean": round(float(vals.mean()), 2),
            "median": round(float(vals.median()), 2),
            "top25": round(float(vals.quantile(0.75)), 2),
            "min": round(float(vals.min()), 2),
            "max": round(float(vals.max()), 2),
            "n": int(len(vals)),
        }
    
    # 个人该期数据
    personal_row = personal.iloc[0]
    personal_data = {}
    for col in x_vars + y_vars:
        personal_data[col] = round(float(personal_row[col]) if pd.notna(personal_row[col]) else 0, 2)
    
    # 该期所有人名单
    others_names = others["姓名"].tolist()
    
    variable_labels = {
        "好友率": "好友率",
        "入群率": "入群率",
        "APP下载率": "APP下载率",
        "开营回复率": "开营回复率",
        "Day1到课率": "Day1到课率",
        "最终转化率": "最终转化率",
        "Day3出直播间转化率": "Day3直播间转化率",
    }
    
    # ── 逐项对比（你 vs 同组同期均值）──
    comparisons = []
    for col in x_vars:
        os = others_stats[col]
        pv = personal_data[col]
        gap = round(pv - os["mean"], 2)
        gap_pct = round((gap / os["mean"]) * 100, 1) if os["mean"] > 0 else 0
        controllable = col not in UNCONTROLLABLE_VARS

        if not controllable:
            # 上游不可控（如好友率由BD引流质量决定），不给"提升该指标"的改善方案
            level, level_icon, potential_gain = "不可控", "⚪", 0
        elif gap >= 2:
            level, level_icon, potential_gain = "优势", "🟢", 0
        elif gap >= -2:
            level, level_icon, potential_gain = "持平", "🟡", 0
        elif gap >= -5:
            level, level_icon = "略低", "🟠"
            coef = _get_personal_coef(df, col, grade)
            potential_gain = abs(min(gap, 0)) * coef if gap < 0 else 0
        else:
            level, level_icon = "短板", "🔴"
            coef = _get_personal_coef(df, col, grade)
            potential_gain = abs(min(gap, 0)) * coef if gap < 0 else 0

        comparisons.append({
            "variable": col,
            "label": variable_labels.get(col, col),
            "personal": pv,
            "benchmark_mean": os["mean"],
            "benchmark_top25": os["top25"],
            "gap": gap,
            "gap_pct": gap_pct,
            "level": level,
            "level_icon": level_icon,
            "controllable": controllable,
            "potential_gain": round(potential_gain, 4),
        })
    
    comparisons.sort(key=lambda x: x["gap"])
    
    # Y变量对比
    y_comparisons = []
    for col in y_vars:
        os = others_stats[col]
        pv = personal_data[col]
        gap = round(pv - os["mean"], 2)
        y_comparisons.append({
            "variable": col,
            "label": variable_labels.get(col, col),
            "personal": pv,
            "benchmark_mean": os["mean"],
            "gap": gap,
        })
    
    # ── 生成建议 ──
    weaknesses = [c for c in comparisons if c["level"] in ("短板", "略低")]
    strengths = [c for c in comparisons if c["level"] == "优势"]
    
    advice_items = []
    if weaknesses:
        weaknesses.sort(key=lambda x: x["potential_gain"], reverse=True)
        top = weaknesses[0]
        advice_items.append({
            "type": "priority",
            "title": f"本期最该优先提升的是「{top['label']}」",
            "detail": f"你 {top['personal']}%，同组均值 {top['benchmark_mean']}%，落后 {abs(top['gap'])} 个百分点。如果追平同组均值，预计转化率可提升约 {top['potential_gain']:.2f}%。",
        })
        for w in weaknesses[1:3]:
            advice_items.append({
                "type": "secondary",
                "title": f"其次关注「{w['label']}」",
                "detail": f"你 {w['personal']}% vs 同组 {w['benchmark_mean']}%，差距 {abs(w['gap'])} 个百分点。",
            })
    
    if strengths:
        s = strengths[0]
        advice_items.append({
            "type": "strength",
            "title": f"「{s['label']}」是你的优势项",
            "detail": f"高出同组 {s['gap']} 个百分点，继续保持，也可以分享经验给团队。",
        })
    
    # 转化率对比
    pc = personal_data.get("最终转化率", 0)
    bc = others_stats.get("最终转化率", {}).get("mean", 0)
    cg = round(pc - bc, 2)
    if cg >= 1:
        conv_note = f"本期你的转化率 {pc}%，高于同组 {bc}%，表现优秀！"
    elif cg >= -0.5:
        conv_note = f"本期你的转化率 {pc}%，与同组 {bc}% 基本持平。"
    else:
        conv_note = f"本期你的转化率 {pc}%，低于同组 {bc}%，差距 {abs(cg)} 个百分点，建议重点关注上述短板。"
    
    if not weaknesses:
        summary = f"本期各项指标均不低于同组均值，整体表现均衡。转化率 {pc}%{'高于' if cg > 0 else '接近'}同组 {bc}%。继续保持！"
    elif len(weaknesses) <= 2:
        summary = f"本期有 {len(weaknesses)} 个环节略低于同组，优先补齐即可。转化率 {pc}%。"
    else:
        summary = f"本期有 {len(weaknesses)} 个环节明显落后于同组，建议从最影响转化率的环节逐个突破。转化率 {pc}%，同组 {bc}%。"
    
    return jsonify({
        "name": name,
        "period": str(period),
        "grade": grade or "全年级",
        "personal_data": personal_data,
        "others_stats": others_stats,
        "others_names": others_names,
        "others_count": len(others),
        "comparisons": comparisons,
        "y_comparisons": y_comparisons,
        "advice": advice_items,
        "conv_note": conv_note,
        "summary": summary,
    })


@app.route("/periods")
def get_periods():
    """返回数据库中所有期次（可按年级/学科筛选），按时间数值排序"""
    grade = request.args.get("grade", "")
    subject = request.args.get("subject", "")
    conn = get_db()
    where = []
    params = []
    if grade:
        where.append("grade = ?"); params.append(grade)
    if subject:
        where.append("subject = ?"); params.append(subject)
    wsql = ("WHERE " + " AND ".join(where)) if where else ""
    rows = conn.execute(
        "SELECT DISTINCT period FROM records " + wsql + " ORDER BY period", params
    ).fetchall()
    conn.close()

    # 数值排序，避免 "0717" < "608" 的字符串排序bug
    def _pk(p):
        try:
            return int(float(str(p["period"])))
        except (ValueError, TypeError):
            return 9999
    rows_sorted = sorted(rows, key=_pk)
    return jsonify({"periods": [r["period"] for r in rows_sorted]})


@app.route("/api/records")
def api_records():
    """返回当前筛选(学科/年级)下的全部记录，供前端查看与编辑。
    必须携带正确访问口令（query 的 passphrase 或 header X-Edit-Pass），
    否则拒绝——保证『不输密码就看不到数据』在后端也成立。"""
    p = request.args.get("passphrase", "") or request.headers.get("X-Edit-Pass", "")
    if not _pass_configured():
        return jsonify({"error": "尚未设置访问口令", "locked": True}), 400
    if not _check_pass(p):
        return jsonify({"error": "口令错误，无法查看数据", "locked": True}), 403
    grade = request.args.get("grade", "")
    subject = request.args.get("subject", "")
    conn = get_db()
    where, params = [], []
    if grade:
        where.append("grade = ?"); params.append(grade)
    if subject:
        where.append("subject = ?"); params.append(subject)
    wsql = ("WHERE " + " AND ".join(where)) if where else ""
    rows = conn.execute(
        "SELECT id, period, name, subject, grade, friend_rate, group_rate, app_rate, "
        "reply_rate, day1_rate, conversion_rate, day3_rate, source, uploaded_at "
        "FROM records " + wsql + " ORDER BY period, name", params
    ).fetchall()
    conn.close()

    def _pk(p):
        try:
            return int(float(str(p["period"])))
        except (ValueError, TypeError):
            return 9999
    rows = sorted(rows, key=lambda r: (_pk(r), str(r["name"])))
    out = [{
        "id": r["id"], "period": r["period"], "name": r["name"], "subject": r["subject"],
        "grade": r["grade"], "friend_rate": r["friend_rate"], "group_rate": r["group_rate"],
        "app_rate": r["app_rate"], "reply_rate": r["reply_rate"], "day1_rate": r["day1_rate"],
        "conversion_rate": r["conversion_rate"], "day3_rate": r["day3_rate"],
        "source": r["source"], "uploaded_at": r["uploaded_at"],
    } for r in rows]
    return jsonify({"records": out})


@app.route("/api/edit-lock-status")
def api_edit_lock_status():
    """返回修改口令是否已设置，前端据此决定显示『设置口令』还是『解锁』界面"""
    return jsonify({"configured": _pass_configured()})


@app.route("/api/edit-passphrase", methods=["POST"])
def api_edit_passphrase():
    """首次设置或修改访问口令。已设置时改口令需先通过 current 校验。"""
    data = request.get_json() or {}
    new = (data.get("new") or "").strip()
    if len(new) < 4:
        return jsonify({"error": "口令至少 4 位"}), 400
    cfg = _load_config()
    if cfg.get("edit_pass_hash"):
        cur = data.get("current", "")
        if not _check_pass(cur):
            return jsonify({"error": "当前口令不正确"}), 403
    cfg["edit_pass_hash"] = _hash_pass(new)
    _save_config(cfg)
    return jsonify({"ok": True})


@app.route("/api/edit-unlock", methods=["POST"])
def api_edit_unlock():
    """校验访问口令。前端解锁后用它发起修改/删除请求。"""
    data = request.get_json() or {}
    p = data.get("passphrase", "")
    if not _pass_configured():
        return jsonify({"error": "尚未设置口令，请先设置"}), 400
    if not _check_pass(p):
        return jsonify({"error": "口令错误"}), 403
    return jsonify({"ok": True})


@app.route("/api/record-update", methods=["POST"])
def api_record_update():
    """修改单条记录的数值指标与文本字段(姓名/学科/年级)，需通过访问口令校验。
    文本字段改动会按录入同款规矩做『唯一约束 + 年级唯一性』校验，
    避免把同人挂到不同年级或造出重复行。"""
    data = request.get_json() or {}
    if not _check_pass(data.get("passphrase", "")):
        return jsonify({"error": "口令校验失败，无法修改数据"}), 403
    rid = data.get("id")
    if not rid:
        return jsonify({"error": "缺少记录 id"}), 400

    conn = get_db()
    row = conn.execute(
        "SELECT id, period, name, subject, grade FROM records WHERE id = ?", (rid,)
    ).fetchone()
    if not row:
        conn.close()
        return jsonify({"error": "记录不存在或已被删除"}), 404

    cur_period = str(row["period"])
    cur_name = str(row["name"])
    cur_subject = str(row["subject"])
    cur_grade = str(row["grade"])

    # 文本字段（允许修改，年级/学科默认沿用原值）
    new_name = str(data.get("name", cur_name) or cur_name).strip()
    new_subject = str(data.get("subject", cur_subject) or cur_subject).strip() or cur_subject
    new_grade = str(data.get("grade", cur_grade) or cur_grade).strip()
    if not new_name:
        conn.close()
        return jsonify({"error": "姓名不能为空"}), 400

    # 数值指标
    metric_fields = ["friend_rate", "group_rate", "app_rate", "reply_rate",
                     "day1_rate", "conversion_rate", "day3_rate"]
    sets, params = [], []
    for f in metric_fields:
        if f in data:
            v = data.get(f)
            try:
                val = None if v in (None, "") else float(v)
            except (ValueError, TypeError):
                conn.close()
                return jsonify({"error": f"[{f}] 不是合法数字"}), 400
            sets.append(f + " = ?")
            params.append(val)

    # 文本字段有变化 → 做唯一约束 / 年级唯一性校验
    text_changed = (new_name != cur_name) or (new_subject != cur_subject) or (new_grade != cur_grade)
    if text_changed:
        # 1) 唯一约束 (subject, grade, period, name) 不可与既有记录(排除自己)撞车
        clash = conn.execute(
            "SELECT id FROM records WHERE subject=? AND grade=? AND period=? AND name=? AND id<>?",
            (new_subject, new_grade, cur_period, new_name, rid)
        ).fetchone()
        if clash:
            conn.close()
            return jsonify({"error": f"已存在记录 {new_subject}/{new_grade}/{cur_period}/{new_name}，修改会重复，请换姓名或年级"}), 400
        # 2) 年级唯一性硬约束：同(姓名,学科)全局只能一个年级
        ex = conn.execute(
            "SELECT DISTINCT grade FROM records WHERE name=? AND subject=? AND id<>?",
            (new_name, new_subject, rid)
        ).fetchall()
        ex_grades = set(str(r[0]) for r in ex if r[0])
        if ex_grades and new_grade not in ex_grades:
            grades_txt = "、".join(sorted(ex_grades))
            conn.close()
            return jsonify({"error": f"{new_name}（{new_subject}）已存在于 {grades_txt}，不能改成 {new_grade}——每人年级应唯一"}), 400
        sets.append("name = ?"); params.append(new_name)
        sets.append("subject = ?"); params.append(new_subject)
        sets.append("grade = ?"); params.append(new_grade)

    if not sets:
        conn.close()
        return jsonify({"error": "没有要更新的字段"}), 400

    params.append(rid)
    conn.execute("UPDATE records SET " + ", ".join(sets) + " WHERE id = ?", params)

    # 文本字段变化时同步 predictions 表，保证复盘/预测分析不串数据
    if text_changed:
        try:
            conn.execute(
                "UPDATE predictions SET name=?, subject=? WHERE period=? AND name=? AND subject=?",
                (new_name, new_subject, cur_period, cur_name, cur_subject)
            )
        except Exception:
            pass
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route("/api/record-delete", methods=["POST"])
def api_record_delete():
    """删除单条记录（需通过访问口令校验，仅平台主人可操作），并清理关联的预测数据"""
    data = request.get_json() or {}
    if not _check_pass(data.get("passphrase", "")):
        return jsonify({"error": "口令校验失败，无法删除数据"}), 403
    rid = data.get("id")
    if not rid:
        return jsonify({"error": "缺少记录 id"}), 400
    conn = get_db()
    row = conn.execute(
        "SELECT period, name, subject FROM records WHERE id = ?", (rid,)
    ).fetchone()
    if not row:
        conn.close()
        return jsonify({"error": "记录不存在或已被删除"}), 404
    conn.execute("DELETE FROM records WHERE id = ?", (rid,))
    # 同步清理该人该期同学科的孤立预测，避免后续分析串数据
    try:
        conn.execute(
            "DELETE FROM predictions WHERE period=? AND name=? AND subject=?",
            (row["period"], row["name"], row["subject"])
        )
    except Exception:
        pass
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


# ──────────────────────────────────────────────────────────────
# A方案：AI 生成复盘结论（一句话总结 + 下期3条优先级建议）
# 严格卡在「过程指标查漏补缺」scope，禁词命中或调用失败 → 返回 None（前端退回静态 advice）
# ──────────────────────────────────────────────────────────────
_AI_CFG_CACHE = None


def _ai_cfg():
    global _AI_CFG_CACHE
    if _AI_CFG_CACHE is None:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ai_config.json")
        try:
            with open(p, encoding="utf-8") as f:
                _AI_CFG_CACHE = json.load(f)
        except Exception:
            _AI_CFG_CACHE = {"enabled": False}
    return _AI_CFG_CACHE


# 命中即作废、退回静态兜底（防止 AI 越界写个人话术/高频催）
_AI_BANNED = ["主动加好友", "每天催三次", "高频触达", "多轮提醒", "多轮催",
              "直播间话术", "追单话术", "逼单", "强推", "话术模板"]


def _ai_call(messages, timeout=20):
    cfg = _ai_cfg()
    if not cfg.get("enabled"):
        return None
    key = cfg.get("deepseek_key", "")
    if not key:
        return None
    base = cfg.get("base_url", "https://api.deepseek.com").rstrip("/")
    url = base + "/chat/completions"
    try:
        resp = requests.post(url, json={
            "model": cfg.get("model", "deepseek-chat"),
            "messages": messages,
            "temperature": 0.85,
            "max_tokens": 1000,
            "stream": False,
        }, headers={"Authorization": "Bearer " + key}, timeout=timeout)
        if resp.status_code != 200:
            return None
        d = resp.json()
        return d.get("choices", [{}])[0].get("message", {}).get("content", "")
    except Exception:
        return None


def ai_review_conclusion(review_items, conv_summary, period, prev_period, name=""):
    """A方案：用 AI 生成复盘结论。喂入过程指标差距+转化小结+硬约束，
    输出一句话总结+3条下期优先级建议。禁词命中或失败返回 None。"""
    cfg = _ai_cfg()
    if not cfg.get("enabled"):
        return None
    # 只喂过程指标（结果指标不喂，避免 AI 染指转化话术）
    proc = [r for r in review_items if not r.get("_is_result")]
    lines = []
    for r in proc:
        arrow = "降" if r["gap"] < 0 else ("升" if r["gap"] > 0 else "平")
        seg = "- " + r["label"] + ": 本期 " + str(r["period_value"]) + "%, 上期(" + prev_period + ") " + str(r["baseline_value"]) + "%, " + arrow + " " + str(abs(r["gap"])) + " 个点"
        if r.get("coefficient"):
            seg += ", 约影响转化率 " + ("+" if r["conv_impact"] >= 0 else "") + str(round(r["conv_impact"], 2)) + "%"
        lines.append(seg)
    data_block = "\n".join(lines)
    who = ("「" + name + "」个人") if name else "团队"

    # ── 把团队手写方案库(ADVICE_LIB)作为背景喂给 AI，让它先学团队打法再延展 ──
    lib_lines = []
    for _var, _variants in ADVICE_LIB.items():
        for _i, _v in enumerate(_variants, 1):
            _seg = "【" + _var + "】方案" + str(_i) + "\n"
            _seg += "  定位: " + (_v.get("control", "") or "") + "\n"
            _seg += "  团队认知: " + (_v.get("note", "") or "") + "\n"
            _sop = _v.get("sop") or []
            if _sop:
                _seg += "  已有SOP: " + " / ".join(_sop) + "\n"
            _ai = _v.get("ai") or []
            if _ai:
                _seg += "  手写AI想法: " + " / ".join(_ai) + "\n"
            lib_lines.append(_seg)
    library_block = "\n".join(lib_lines)

    prompt = (
        "你是小学社群运营的数据分析助手。请先理解下方「团队已有手写方案（历史沉淀）」，再基于「这期 vs 上期」的过程指标数据生成复盘结论。\n\n"
        "【团队已有手写方案——请先学习我们的打法与红线，再在其基础上延展，不要与它矛盾】\n" + library_block + "\n\n"
        "【数据】" + who + "复盘，期次" + period + " vs 上期" + prev_period + "：\n" + data_block + "\n\n"
        "【转化率小结】" + conv_summary + "\n\n"
        "【硬约束——必须遵守，违反即作废】\n"
        "1. 你只做「过程指标查漏补缺」：指出好友率/入群率/APP下载率/开营回复率/Day1到课率 哪个环节这期比上期掉了、该怎么补。\n"
        "2. 绝不输出个人销售话术、直播间追单话术、转化技巧、D3/D4逼单话术——那些是个人业务，不归你写。\n"
        "3. 高频催/多轮提醒会被投诉封号，已封过号；简单1-2次有价值提醒可以，但不得建议高频触达。\n"
        "4. 好友率是上游BD引流质量决定、你完全不可控（家长加不加你决定不了，也不能主动加），不要把好友率列为可优化短板、不要给任何「提升好友率」的改善建议；它只作为数据展示。若总结或建议里提到好友率，须注明「好友率不可控、非你所能优化、无需动作」。\n"
        "5. 作业是AI批改、运营不逐个看，别建议「看作业记痛点」。\n\n"
        "【输出格式——严格按此格式，不要多余解释】\n"
        "总结：一句话，说清楚这期比上期最大问题是哪个环节、丢了多少。\n"
        "建议1：<环节> <沿用或改进团队已有打法的具体动作，不超30字>\n"
        "建议2：<环节> <具体动作，不超30字>\n"
        "建议3：<环节> <具体动作，不超30字>\n"
        "新方向1：在团队已有方案基础上延伸的1个新角度（仍限过程指标查漏补缺，不写销售话术/高频催），不超40字\n"
        "新方向2：同上，另一个新角度\n"
        "新方向3：同上，再一个不同角度"
    )
    raw = _ai_call([{"role": "user", "content": prompt}], timeout=cfg.get("timeout", 20))
    if not raw or not raw.strip():
        return None
    # 禁词过滤
    for w in _AI_BANNED:
        if w in raw:
            return None
    # 解析输出
    summary = ""
    actions = []

    def _split_kv(s):
        for sep in ("：", ":"):
            if sep in s:
                return s.split(sep, 1)[-1].strip()
        return s
    new_directions = []
    for ln in raw.strip().splitlines():
        s = ln.strip()
        if not s:
            continue
        if s.startswith("总结"):
            summary = _split_kv(s)
        elif s.startswith("建议"):
            actions.append(_split_kv(s))
        elif s.startswith("新方向"):
            new_directions.append(_split_kv(s))
    if not summary and not actions and not new_directions:
        return None
    return {"summary": summary, "actions": actions[:3], "new_directions": new_directions[:3]}


@app.route("/period-review", methods=["POST"])
def period_review():
    """
    本期复盘（团队视角）：
    - 杠杆模型 = 用全量历史数据(所有期/所有人)回归得到的各环节"真实影响力"系数
    - 复盘对象 = 选定某一期，自动找紧挨着的「上一期」，逐项对比这期 vs 上期，
      看这期比上期哪里掉了链子、下期该往哪调。
    二者用途不同：模型靠全量数据变准；复盘必须按每期单独看、且和上一期比。
    """
    data = request.get_json() or {}
    period = _normalize_period(data.get("period", ""))
    grade = data.get("grade", "")
    subject = data.get("subject", "")
    name = str(data.get("name", "")).strip()

    if not period:
        return jsonify({"error": "请选择期次"}), 400

    df = get_all_data(grade=grade if grade else None, subject=subject if subject else None)
    if df is None:
        return jsonify({"error": "数据库中暂无数据"}), 400

    period_df = df[df["期次"].astype(str) == period]
    if name:
        # 个人环比模式：只看这个人当期的单条数据
        person_df = period_df[period_df["姓名"].astype(str) == name]
        if len(person_df) == 0:
            return jsonify({"error": f"「{name}」在「{period}」期没有数据"}), 400
        period_df = person_df
    else:
        if len(period_df) == 0:
            return jsonify({"error": f"未找到「{period}」期的数据"}), 400

    # ── 杠杆模型：用全量数据回归（数据越多越准）──
    analysis = run_analysis(df)
    coef_map = {r["variable"]: r for r in analysis["final"]}

    x_vars = ["好友率", "入群率", "APP下载率", "开营回复率", "Day1到课率"]
    y_vars = ["最终转化率", "Day3出直播间转化率"]

    # ── 本期实际值（团队均值）──
    period_mean = {}
    for c in x_vars + y_vars:
        period_mean[c] = round(float(period_df[c].dropna().mean()), 2)

    # ── 参与本期复盘的团队名单 ──
    people = period_df["姓名"].astype(str).unique().tolist()

    # ── 基线：选中期的「上一期」均值（复盘 = 这期 vs 上一期）──
    # 用数值排序期次，避免 "0717" < "608" 的字符串排序bug
    def _period_key(p):
        try:
            return int(float(str(p)))
        except (ValueError, TypeError):
            return 9999
    all_periods = sorted(df["期次"].astype(str).unique().tolist(), key=_period_key)
    cur_idx = all_periods.index(period)
    if cur_idx <= 0:
        return jsonify({
            "error": f"「{period}」是最早的一期，没有上一期可对比，请选择较后的期次（如想看最早的期，可先补一期更早的数据）"
        }), 400
    prev_period = all_periods[cur_idx - 1]
    prev_df = df[df["期次"].astype(str) == prev_period]
    if len(prev_df) == 0:
        return jsonify({"error": f"上一期「{prev_period}」没有数据，无法对比"}), 400
    if name:
        prev_person = prev_df[prev_df["姓名"].astype(str) == name]
        if len(prev_person) == 0:
            return jsonify({"error": f"「{name}」在上一期「{prev_period}」没有数据，无法做个人环比"}), 400
        baseline_source = prev_person
    else:
        baseline_source = prev_df
    baseline_mean = {}
    for c in x_vars + y_vars:
        baseline_mean[c] = round(_safe_float(baseline_source[c].dropna().mean()), 2)

    # ── 逐项复盘：本期 vs 基线，用杠杆系数折算成转化率得失 ──
    # 建议使用全局 ADVICE_LIB（模块级结构化方案）
    label_map = {
        "Day1到课率": "Day1到课率",
        "APP下载率": "APP下载率",
        "开营回复率": "开营回复率",
        "入群率": "入群率",
        "好友率": "好友率",
    }

    review_items = []
    for x in x_vars:
        coef = abs(coef_map.get(x, {}).get("coefficient", 0))
        cur = period_mean[x]
        base = baseline_mean[x]
        gap = round(cur - base, 2)            # 本期相对常态的变化
        conv_impact = round(gap * coef, 3)    # 对转化率的影响（百分点，可正可负）
        review_items.append({
            "variable": x,
            "label": label_map.get(x, x),
            "period_value": cur,
            "baseline_value": base,
            "gap": gap,
            "conv_impact": conv_impact,
            "coefficient": round(coef, 4),
            "controllable": x not in UNCONTROLLABLE_VARS,
            "action": get_advice(x, period, name=name, grade=grade, subject=subject, cur_val=cur, base_val=base),
        })
    # 损失最大的排最前（conv_impact 越小越靠前）
    review_items.sort(key=lambda r: r["conv_impact"])

    # ── 追加结果指标（Day3直播间转化率、最终转化率）──
    result_tips = {
        "Day3出直播间转化率": "转化打法属于个人业务，平台不给话术。先检查前面的过程环节（下载/回复/到课）本期有没有掉，把过程漏点补上。",
        "最终转化率": "这是最终结果，由前面各过程环节共同决定。平台只负责帮你定位过程漏点，优先补齐上面红项。",
    }
    for y in y_vars:
        cur = period_mean.get(y)
        base = baseline_mean.get(y)
        if cur is None or base is None:
            continue
        gap = round(cur - base, 2)
        # 结果指标：gap 本身就是转化率变化
        review_items.append({
            "variable": y,
            "label": y,
            "period_value": cur,
            "baseline_value": base,
            "gap": gap,
            "conv_impact": gap,
            "coefficient": None,
            "action": result_tips.get(y, ""),
            "_is_result": True,  # 前端可用来区分样式
        })

    # ── 团队转化率 vs 基线 ──
    conv_gap = round(period_mean["最终转化率"] - baseline_mean["最终转化率"], 2)

    # ── 生成复盘结论 ──
    weak = [r for r in review_items if r["gap"] < -1 and r.get("controllable")]
    strong = [r for r in review_items if r["gap"] > 1]

    advice = []
    if weak:
        weak.sort(key=lambda r: r["conv_impact"])  # 损失最大优先
        top = weak[0]
        advice.append({
            "type": "priority",
            "title": f"本期最大短板：「{top['label']}」",
            "detail": f"本期仅 {top['period_value']}%，比上一期（{prev_period}）{top['baseline_value']}% 低 {abs(top['gap'])} 个百分点。"
                      f"按模型测算，这一项大约让本期转化率少了 {abs(top['conv_impact']):.2f}%。",
        })
        for w in weak[1:3]:
            advice.append({
                "type": "secondary",
                "title": f"其次关注：「{w['label']}」",
                "detail": f"本期 {w['period_value']}% vs 上期（{prev_period}）{w['baseline_value']}%，差 {abs(w['gap'])} 个点，约损失 {abs(w['conv_impact']):.2f}% 转化。",
            })
    else:
        # 本期没有明显下降的环节，仍给出"最弱一项"作为可优化点
        lowest = next((r for r in review_items if r.get("controllable")), review_items[0])  # 跳过不可控项
        if lowest["gap"] < 0:
            advice.append({
                "type": "secondary",
                "title": f"可优化点：「{lowest['label']}」",
                "detail": f"本期 {lowest['period_value']}% 略低于上一期（{prev_period}）{lowest['baseline_value']}% 约 {abs(lowest['gap'])} 个点，可作为下期微调目标。",
            })
        else:
            advice.append({
                "type": "secondary",
                "title": f"下期可继续突破：「{lowest['label']}」",
                "detail": f"本期 {lowest['period_value']}% 已是相对最弱项，上一期（{prev_period}）{lowest['baseline_value']}%，继续保持。",
            })
    if strong:
        s = strong[0]
        advice.append({
            "type": "strength",
            "title": f"本期亮点：「{s['label']}」",
            "detail": f"本期 {s['period_value']}%，高于上一期（{prev_period}）{s['baseline_value']}% 约 {s['gap']} 个点，继续保持。",
        })

    # ── 下期行动清单（具体可执行，按损失排序，最多列最该补的2项）──
    # 建议使用全局 ADVICE_LIB（模块级结构化方案）
    action_checklist = []
    regressed = [r for r in review_items if r["gap"] < -1 and r.get("controllable")]
    regressed.sort(key=lambda r: r["conv_impact"])
    for r in regressed[:2]:
        action_checklist.append({
            "variable": r["label"],
            "current": r["period_value"],
            "target": r["baseline_value"],
            "gap": r["gap"],
            "conv_loss": r["conv_impact"],
            "action": r["action"],
        })

    # 转化率小结（团队 / 个人自适应主语）
    subject = f"{name}本期" if name else "本期团队"
    if conv_gap <= -0.5:
        conv_summary = f"{subject}最终转化率 {period_mean['最终转化率']}%，比上一期（{prev_period}）{baseline_mean['最终转化率']}% 低 {abs(conv_gap)} 个百分点，需要复盘找原因。"
    elif conv_gap >= 0.5:
        conv_summary = f"{subject}最终转化率 {period_mean['最终转化率']}%，高于上一期（{prev_period}）{baseline_mean['最终转化率']}% 约 {conv_gap} 个百分点，表现不错，总结下经验。"
    else:
        conv_summary = f"{subject}最终转化率 {period_mean['最终转化率']}%，与上一期（{prev_period}）{baseline_mean['最终转化率']}% 基本持平。"

    # 杠杆模型（全量数据）排名，用于说明"数据越多越准"
    leverage = [{
        "variable": r["variable"],
        "coefficient": r["coefficient"],
        "significant": r["significant"],
    } for r in analysis["final"]]

    # ── A方案：AI 生成复盘结论（失败/禁词命中 → None，前端退回静态 advice）──
    ai_concl = ai_review_conclusion(review_items, conv_summary, period, prev_period, name)

    return jsonify({
        "period": period,
        "prev_period": prev_period,
        "grade": grade or "全年级",
        "is_personal": bool(name),
        "person_name": name or "",
        "team_size": len(period_df),
        "people": people,
        "period_mean": period_mean,
        "baseline_mean": baseline_mean,
        "conv_gap": conv_gap,
        "conv_summary": conv_summary,
        "action_checklist": action_checklist,
        "review_items": review_items,
        "advice": advice,
        "ai_conclusion": ai_concl,
        "leverage": leverage,
        "total_rows": len(df),
    })


def _get_personal_coef(df, var_name, grade):
    """获取personal系数：在大盘数据上跑回归得到该变量的系数"""
    if var_name == "Day1到课率":
        return 0  # Day1到课率自身就是直接的
    try:
        df["暑假"] = df["期次"].apply(
            lambda x: 1 if str(int(float(str(x)))).zfill(4) >= "0622" else 0
        )
        X = df[[var_name, "暑假"]].dropna()
        y = df.loc[X.index, "最终转化率"]
        X_sm = sm.add_constant(X)
        model = sm.OLS(y, X_sm).fit()
        return abs(_safe_float(model.params.get(var_name, 0)))
    except:
        return 0


@app.route("/sample-data")
def sample_data():
    sample_path = os.path.join(os.path.dirname(__file__), "..", "社群运营回归数据.csv")
    if os.path.exists(sample_path):
        with open(sample_path, "r", encoding="utf-8") as f:
            return f.read(), 200, {"Content-Type": "text/csv; charset=utf-8"}
    return "示例数据文件不存在", 404


@app.route("/download-template")
def download_template():
    """下载 Excel/CSV 导入模板：含「学科」列，支持整表按行标注学科批量导入"""
    header = ["期次", "姓名", "学科", "年级", "好友率", "入群率", "APP下载率",
              "开营回复率", "Day1到课率", "Day3出直播间转化率", "最终转化率"]
    examples = [
        ["717", "志豪", "数学", "五年级", "85", "70", "60", "75", "65", "1.0", "4.06"],
        ["717", "小李", "语文", "五年级", "80", "65", "55", "70", "60", "0.8", "3.50"],
        ["717", "小王", "英语", "六年级", "82", "68", "58", "72", "62", "0.9", "3.80"],
    ]
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(header)
    writer.writerows(examples)
    # utf-8-sig 带 BOM，Excel 打开中文不乱码
    csv_bytes = buf.getvalue().encode("utf-8-sig")
    fname = "社群运营数据导入模板.csv"
    # 文件名含中文，必须用 RFC5987 百分号编码，否则 HTTP 头非法导致连接重置
    disp = 'attachment; filename="import_template.csv"; filename*=UTF-8\'\'' + quote(fname)
    return Response(
        csv_bytes,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": disp},
    )


# ── 下期预测 & 真实对比 ────────────────────────────────────
@app.route("/latest-metrics")
def latest_metrics_route():
    """返回最新一期团队均值过程指标，用于预测表单预填"""
    grade = request.args.get("grade", "")
    subject = request.args.get("subject", "")
    data = latest_metrics(grade=grade if grade else None, subject=subject if subject else None)
    if not data:
        return jsonify({"error": "暂无数据，无法预填"}), 400
    return jsonify(data)


@app.route("/predict", methods=["POST"])
def predict_route():
    """根据输入的过程指标，预测下期 Day3 / 最终转化率"""
    data = request.get_json() or {}
    period = _normalize_period(data.get("period", ""))
    name = str(data.get("name", "")).strip()
    grade = data.get("grade", "")
    subject = data.get("subject", "")
    metrics = data.get("metrics", {})
    if not period:
        return jsonify({"error": "请填写目标期次"}), 400
    # 校验 5 个过程指标
    miss = [x for x in PREDICT_X if metrics.get(x) in (None, "")]
    if miss:
        return jsonify({"error": f"请填写完整的过程指标：{', '.join(miss)}"}), 400
    try:
        metrics = {x: float(metrics.get(x, 0) or 0) for x in PREDICT_X}
    except (ValueError, TypeError):
        return jsonify({"error": "过程指标必须是数字"}), 400

    pred = predict_conversion(period, metrics, grade=grade if grade else None, subject=subject if subject else None, name=name)
    if pred is None:
        if name:
            return jsonify({"error": f"「{name}」的个人历史数据不足（需≥5期有效记录），无法做个人预测。可清空姓名改用团队整体，或先多录入几期。"}), 400
        return jsonify({"error": "数据库暂无足够数据用于预测"}), 400
    return jsonify({
        "period": period,
        "name": name,
        "subject": subject,
        "metrics": metrics,
        "predicted_day3": pred["Day3出直播间转化率"],
        "predicted_final": pred["最终转化率"],
    })


@app.route("/save-prediction", methods=["POST"])
def save_prediction_route():
    """保存一次下期预测，待真实数据上传后做对比"""
    data = request.get_json() or {}
    period = _normalize_period(data.get("period", ""))
    name = str(data.get("name", "")).strip()
    subject = data.get("subject", "") or "数学"
    pred_day3 = data.get("predicted_day3")
    pred_final = data.get("predicted_final")
    metrics = data.get("metrics", {})
    if not period or pred_day3 is None or pred_final is None:
        return jsonify({"error": "请先预测并填写完整信息"}), 400
    conn = get_db()
    conn.execute("""
        INSERT INTO predictions (period, name, subject, predicted_day3, predicted_final, metrics_json, created_at)
        VALUES (?, ?, ?, ?, ?, ?, datetime('now','localtime'))
        ON CONFLICT(subject, period, name) DO UPDATE SET
            predicted_day3=excluded.predicted_day3,
            predicted_final=excluded.predicted_final,
            metrics_json=excluded.metrics_json,
            subject=excluded.subject,
            created_at=datetime('now','localtime')
    """, (period, name, subject, float(pred_day3), float(pred_final), json.dumps(metrics, ensure_ascii=False)))
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "period": period, "name": name})


@app.route("/prediction-compare", methods=["POST"])
def prediction_compare_route():
    """对比某期的预测值 vs 真实值"""
    data = request.get_json() or {}
    period = _normalize_period(data.get("period", ""))
    name = str(data.get("name", "")).strip()
    subject = data.get("subject", "")
    if not period:
        return jsonify({"error": "请选择期次"}), 400

    conn = get_db()
    if subject:
        row = conn.execute(
            "SELECT predicted_day3, predicted_final FROM predictions WHERE period=? AND name=? AND subject=?",
            (period, name, subject)
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT predicted_day3, predicted_final FROM predictions WHERE period=? AND name=?",
            (period, name)
        ).fetchone()
    if row is None:
        conn.close()
        return jsonify({"error": f"「{period}」{('（' + name + '）') if name else ''} 还没有保存过的预测，无法对比。请先做一次预测并保存。"}), 400

    # 取真实数据：个人模式取该人当期，团队模式取当期团队均值
    df = get_all_data(subject=subject if subject else None)
    if df is None or len(df) == 0:
        conn.close()
        return jsonify({"error": "暂无真实数据"}), 400
    period_df = df[df["期次"].astype(str) == period]
    if len(period_df) == 0:
        conn.close()
        return jsonify({"error": f"「{period}」还没有上传真实数据，无法对比"}), 400

    if name:
        actual_df = period_df[period_df["姓名"].astype(str) == name]
        scope = name
    else:
        actual_df = period_df
        scope = "团队"
    actual_day3 = round(float(actual_df["Day3出直播间转化率"].dropna().mean()), 2) if actual_df["Day3出直播间转化率"].notna().any() else None
    actual_final = round(float(actual_df["最终转化率"].dropna().mean()), 2) if actual_df["最终转化率"].notna().any() else None

    conn.close()
    return jsonify({
        "period": period,
        "name": name,
        "scope": scope,
        "predicted_day3": row["predicted_day3"],
        "predicted_final": row["predicted_final"],
        "actual_day3": actual_day3,
        "actual_final": actual_final,
        "gap_day3": round(actual_day3 - row["predicted_day3"], 2) if actual_day3 is not None else None,
        "gap_final": round(actual_final - row["predicted_final"], 2) if actual_final is not None else None,
    })


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  过程数据分析平台")
    print("  浏览器访问: http://localhost:5000")
    print("  支持: Excel/CSV上传 / 手动录入数据")
    print("=" * 60 + "\n")
    # 常驻模式：关闭调试重载器，更稳定，适合开机自启/长期运行
    app.run(host="0.0.0.0", port=5000, debug=False)
