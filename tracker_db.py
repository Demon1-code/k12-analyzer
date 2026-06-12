"""
高中网课销售追踪助手 - SQLite 数据层
管理客户档案与跟进记录的持久化存储
"""

import sqlite3
import json
from pathlib import Path
from datetime import date, datetime

DB_PATH = Path(__file__).parent / "sales_tracker.db"

STAGES = ["新线索", "已联系", "已试听", "谈单中", "已成交", "已流失"]
ACTIVE_STAGES = ["新线索", "已联系", "已试听", "谈单中"]

FOLLOWUP_METHODS = ["微信", "电话", "短信", "面谈/视频", "试听课", "朋友圈互动", "其他"]
FOLLOWUP_RESULTS = ["有进展", "正常维护", "客户犹豫", "暂无回应", "约定试听", "约定下次沟通", "成交", "明确拒绝"]


def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with _conn() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_name   TEXT NOT NULL,
            contact       TEXT DEFAULT '',
            student_name  TEXT DEFAULT '',
            grade         TEXT NOT NULL,
            subjects      TEXT DEFAULT '[]',
            budget        INTEGER DEFAULT 5000,
            performance   TEXT DEFAULT '中等',
            concerns      TEXT DEFAULT '[]',
            intent_signals TEXT DEFAULT '[]',
            source        TEXT DEFAULT '',
            stage         TEXT DEFAULT '新线索',
            intent_score  INTEGER DEFAULT 50,
            profile_type  TEXT DEFAULT '',
            notes         TEXT DEFAULT '',
            next_followup TEXT DEFAULT '',
            created_at    TEXT NOT NULL,
            updated_at    TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS followups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            fu_date     TEXT NOT NULL,
            method      TEXT DEFAULT '',
            content     TEXT DEFAULT '',
            result      TEXT DEFAULT '',
            next_date   TEXT DEFAULT '',
            created_at  TEXT NOT NULL,
            FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
        );
        """)


def _row_to_customer(row) -> dict:
    d = dict(row)
    for key in ("subjects", "concerns", "intent_signals"):
        try:
            d[key] = json.loads(d.get(key) or "[]")
        except Exception:
            d[key] = []
    return d


def add_customer(data: dict) -> int:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as conn:
        cur = conn.execute(
            """INSERT INTO customers
               (parent_name, contact, student_name, grade, subjects, budget,
                performance, concerns, intent_signals, source, stage,
                intent_score, profile_type, notes, next_followup,
                created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                data["parent_name"], data.get("contact", ""), data.get("student_name", ""),
                data["grade"], json.dumps(data.get("subjects", []), ensure_ascii=False),
                data.get("budget", 5000), data.get("performance", "中等"),
                json.dumps(data.get("concerns", []), ensure_ascii=False),
                json.dumps(data.get("intent_signals", []), ensure_ascii=False),
                data.get("source", ""), data.get("stage", "新线索"),
                data.get("intent_score", 50), data.get("profile_type", ""),
                data.get("notes", ""), data.get("next_followup", ""),
                now, now,
            ),
        )
        return cur.lastrowid


def update_customer(cid: int, data: dict):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as conn:
        conn.execute(
            """UPDATE customers SET
               parent_name=?, contact=?, student_name=?, grade=?, subjects=?,
               budget=?, performance=?, concerns=?, intent_signals=?, source=?,
               stage=?, intent_score=?, profile_type=?, notes=?, next_followup=?,
               updated_at=?
               WHERE id=?""",
            (
                data["parent_name"], data.get("contact", ""), data.get("student_name", ""),
                data["grade"], json.dumps(data.get("subjects", []), ensure_ascii=False),
                data.get("budget", 5000), data.get("performance", "中等"),
                json.dumps(data.get("concerns", []), ensure_ascii=False),
                json.dumps(data.get("intent_signals", []), ensure_ascii=False),
                data.get("source", ""), data.get("stage", "新线索"),
                data.get("intent_score", 50), data.get("profile_type", ""),
                data.get("notes", ""), data.get("next_followup", ""),
                now, cid,
            ),
        )


def set_stage(cid: int, stage: str):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as conn:
        conn.execute("UPDATE customers SET stage=?, updated_at=? WHERE id=?", (stage, now, cid))


def set_next_followup(cid: int, next_date: str):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as conn:
        conn.execute("UPDATE customers SET next_followup=?, updated_at=? WHERE id=?",
                     (next_date, now, cid))


def delete_customer(cid: int):
    with _conn() as conn:
        conn.execute("DELETE FROM followups WHERE customer_id=?", (cid,))
        conn.execute("DELETE FROM customers WHERE id=?", (cid,))


def get_customer(cid: int) -> dict | None:
    with _conn() as conn:
        row = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    return _row_to_customer(row) if row else None


def list_customers(stage: str = "", grade: str = "", keyword: str = "") -> list:
    sql, params = "SELECT * FROM customers WHERE 1=1", []
    if stage:
        sql += " AND stage=?"
        params.append(stage)
    if grade:
        sql += " AND grade=?"
        params.append(grade)
    if keyword:
        sql += " AND (parent_name LIKE ? OR student_name LIKE ? OR contact LIKE ?)"
        params += [f"%{keyword}%"] * 3
    sql += " ORDER BY intent_score DESC, updated_at DESC"
    with _conn() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [_row_to_customer(r) for r in rows]


def add_followup(cid: int, fu_date: str, method: str, content: str,
                 result: str, next_date: str = ""):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as conn:
        conn.execute(
            """INSERT INTO followups (customer_id, fu_date, method, content, result, next_date, created_at)
               VALUES (?,?,?,?,?,?,?)""",
            (cid, fu_date, method, content, result, next_date, now),
        )
        conn.execute("UPDATE customers SET next_followup=?, updated_at=? WHERE id=?",
                     (next_date, now, cid))


def list_followups(cid: int) -> list:
    with _conn() as conn:
        rows = conn.execute(
            "SELECT * FROM followups WHERE customer_id=? ORDER BY fu_date DESC, id DESC", (cid,)
        ).fetchall()
    return [dict(r) for r in rows]


def followup_count(cid: int) -> int:
    with _conn() as conn:
        row = conn.execute("SELECT COUNT(*) AS n FROM followups WHERE customer_id=?", (cid,)).fetchone()
    return row["n"]


def last_followup(cid: int) -> dict | None:
    with _conn() as conn:
        row = conn.execute(
            "SELECT * FROM followups WHERE customer_id=? ORDER BY fu_date DESC, id DESC LIMIT 1", (cid,)
        ).fetchone()
    return dict(row) if row else None


# ─── 工作台查询 ────────────────────────────────────────────────────────────────

def get_today_tasks() -> dict:
    """返回 overdue / today / upcoming(3天内) / stale(7天未跟进的活跃客户)"""
    today = date.today().isoformat()
    customers = list_customers()
    active = [c for c in customers if c["stage"] in ACTIVE_STAGES]

    overdue, today_due, upcoming, stale, no_plan = [], [], [], [], []
    for c in active:
        nf = c.get("next_followup") or ""
        if nf:
            if nf < today:
                overdue.append(c)
            elif nf == today:
                today_due.append(c)
            elif (datetime.fromisoformat(nf) - datetime.fromisoformat(today)).days <= 3:
                upcoming.append(c)
        else:
            lf = last_followup(c["id"])
            ref = lf["fu_date"] if lf else c["created_at"][:10]
            days = (datetime.fromisoformat(today) - datetime.fromisoformat(ref)).days
            if days >= 7:
                stale.append({**c, "_idle_days": days})
            else:
                no_plan.append(c)

    overdue.sort(key=lambda c: c.get("next_followup", ""))
    today_due.sort(key=lambda c: -c["intent_score"])
    upcoming.sort(key=lambda c: c.get("next_followup", ""))
    stale.sort(key=lambda c: -c.get("_idle_days", 0))
    return {"overdue": overdue, "today": today_due, "upcoming": upcoming,
            "stale": stale, "no_plan": no_plan}


def get_funnel_stats() -> dict:
    customers = list_customers()
    by_stage = {s: 0 for s in STAGES}
    for c in customers:
        by_stage[c.get("stage", "新线索")] = by_stage.get(c.get("stage", "新线索"), 0) + 1

    total = len(customers)
    won = by_stage.get("已成交", 0)
    lost = by_stage.get("已流失", 0)
    closed = won + lost
    return {
        "total": total,
        "by_stage": by_stage,
        "won": won,
        "lost": lost,
        "active": total - closed,
        "win_rate": round(won / closed * 100, 1) if closed else 0.0,
    }


def get_source_stats() -> dict:
    customers = list_customers()
    stats = {}
    for c in customers:
        src = c.get("source") or "未知"
        s = stats.setdefault(src, {"total": 0, "won": 0})
        s["total"] += 1
        if c["stage"] == "已成交":
            s["won"] += 1
    return stats
