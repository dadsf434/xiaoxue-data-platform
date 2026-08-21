# -*- coding: utf-8 -*-
"""
演示数据种子脚本（仅供 GitHub 开源版本地预览使用）。

说明：
- 本脚本插入的是**完全虚构**的演示数据（虚构姓名 + 虚构指标），不含任何真实业务数据。
- 首次克隆仓库后数据库为空，运行本脚本可快速填充若干期、若干人的演示记录，
  便于直接在页面看到「个人环比复盘 / 个人诊断 / 团队复盘 / 预测」等模块的效果。
- 重复运行安全：会先清除 source='demo' 的旧演示数据再插入。

用法：
    python seed_demo_data.py
"""

import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data.db")


def seed():
    # 确保数据表存在（首次运行仓库为空库时也能直接插入）
    try:
        from app import init_db
        init_db()
    except Exception as e:
        print("警告：无法调用 app.init_db()（%s），将尝试直接建表。" % e)
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            """CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period TEXT NOT NULL, name TEXT NOT NULL, subject TEXT DEFAULT '数学',
                friend_rate REAL, group_rate REAL, app_rate REAL, reply_rate REAL,
                day1_rate REAL, conversion_rate REAL, day3_rate REAL,
                source TEXT DEFAULT 'upload', uploaded_at TEXT,
                grade TEXT DEFAULT ''
            )"""
        )
        conn.commit()
        conn.close()

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # 清掉上一次的种子数据，保证可重复运行
    cur.execute("DELETE FROM records WHERE source = 'demo'")

    # 虚构人员（明显为假名，避免与真实数据混淆）
    people = ["演示甲", "演示乙", "演示丙", "演示丁", "演示戊"]
    periods = ["D01", "D02", "D03"]  # 三期，便于看环比

    # 为每个人生成三期过程指标（百分比数值，Day3 <= 最终）
    demo = []
    for name in people:
        base = 70 + (hash(name) % 15)  # 每人一个基线，制造差异
        for i, period in enumerate(periods):
            friend = base + i * 2
            group = friend - 3
            app = group - 8
            reply = app + 1
            day1 = reply - 5
            day3 = max(0, day1 - 12)  # Day3 出直播间转化率
            final = day3 + 4 if day3 + 4 <= day1 else day1  # 最终 >= Day3，且 <= Day1
            demo.append((
                period, name, "数学", "五年级",
                friend, group, app, reply, day1, final, day3, "demo"
            ))

    cur.executemany(
        """INSERT INTO records
           (period, name, subject, grade, friend_rate, group_rate, app_rate,
            reply_rate, day1_rate, conversion_rate, day3_rate, source)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        demo,
    )
    conn.commit()
    n = cur.execute("SELECT COUNT(*) FROM records WHERE source='demo'").fetchone()[0]
    conn.close()
    print(f"已写入演示数据 {n} 条（source='demo'，均为虚构数据）。")


if __name__ == "__main__":
    seed()
