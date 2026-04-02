import io
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st

DB_PATH = Path("bcs_tracker.db")
DEFAULT_IMPORT_PATH = Path("Last Blitzkrieg Step Loses.xlsx")

ALLIED_GROUP_STARTS = [1, 6, 11, 16, 21, 26, 31, 36]
GERMAN_GROUP_STARTS = [1, 6, 11, 16, 21, 26]
IGNORED_ALLIES_P55_BLOCK = (55, 1000, 16, 20)  # rows 55+, cols P:T (1-based) ignored for v1


@dataclass
class UnitRecord:
    side: str
    formation: str
    unit_name: str
    max_steps: int
    armor_flag: Optional[bool] = None


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS units (
    unit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    side TEXT NOT NULL,
    formation TEXT NOT NULL,
    unit_name TEXT NOT NULL,
    max_steps INTEGER NOT NULL,
    armor_flag INTEGER,
    status TEXT NOT NULL DEFAULT 'ACTIVE',
    notes TEXT DEFAULT '',
    UNIQUE(side, formation, unit_name)
);

CREATE TABLE IF NOT EXISTS unit_state (
    unit_id INTEGER PRIMARY KEY,
    current_losses INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (unit_id) REFERENCES units(unit_id)
);

CREATE TABLE IF NOT EXISTS game_meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS action_log (
    action_id INTEGER PRIMARY KEY AUTOINCREMENT,
    turn_number INTEGER NOT NULL,
    side TEXT NOT NULL,
    action_type TEXT NOT NULL,
    unit_id INTEGER NOT NULL,
    amount INTEGER NOT NULL,
    previous_losses INTEGER NOT NULL,
    new_losses INTEGER NOT NULL,
    note TEXT DEFAULT '',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (unit_id) REFERENCES units(unit_id)
);
"""


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA_SQL)
    conn.execute("INSERT OR IGNORE INTO game_meta (key, value) VALUES ('turn_number', '1')")

    existing_cols = [row[1] for row in conn.execute("PRAGMA table_info(units)").fetchall()]
    if "armor_flag" not in existing_cols and "is_available" in existing_cols:
        conn.execute("ALTER TABLE units RENAME COLUMN is_available TO armor_flag")
    elif "armor_flag" not in existing_cols:
        conn.execute("ALTER TABLE units ADD COLUMN armor_flag INTEGER")

    conn.commit()


def excel_bool_to_python(value) -> Optional[bool]:
    if pd.isna(value) or value is None:
        return None
    if isinstance(value, bool):
        return value
    value_str = str(value).strip().lower()
    if value_str in {"true", "yes", "1"}:
        return True
    if value_str in {"false", "no", "0"}:
        return False
    return None


def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def parse_group(df: pd.DataFrame, side: str, starts: List[int], ignore_block: Optional[Tuple[int, int, int, int]] = None) -> List[UnitRecord]:
    records: List[UnitRecord] = []

    for start_col in starts:
        formation_col = start_col - 1
        unit_col = start_col
        steps_col = start_col + 1
        av_col = start_col + 3

        for row_idx in range(1, len(df)):
            excel_row = row_idx + 1
            excel_col_1based = formation_col + 1

            if ignore_block:
                r1, r2, c1, c2 = ignore_block
                if r1 <= excel_row <= r2 and c1 <= excel_col_1based <= c2:
                    continue

            formation = clean_text(df.iat[row_idx, formation_col]) if formation_col < df.shape[1] else ""
            unit_name = clean_text(df.iat[row_idx, unit_col]) if unit_col < df.shape[1] else ""
            steps_val = df.iat[row_idx, steps_col] if steps_col < df.shape[1] else None
            av_val = df.iat[row_idx, av_col] if av_col < df.shape[1] else None

            if not formation and not unit_name:
                continue
            if not formation or not unit_name or pd.isna(steps_val):
                continue

            try:
                max_steps = int(float(steps_val))
            except (TypeError, ValueError):
                continue

            records.append(
                UnitRecord(
                    side=side,
                    formation=formation,
                    unit_name=unit_name,
                    max_steps=max_steps,
                    armor_flag=excel_bool_to_python(av_val),
                )
            )
    return records


def parse_workbook(uploaded_file) -> pd.DataFrame:
    xls = pd.ExcelFile(uploaded_file)
    allies_df = pd.read_excel(xls, sheet_name="Allies", header=None)
    german_df = pd.read_excel(xls, sheet_name="German", header=None)

    records = []
    records.extend(parse_group(allies_df, "Allies", ALLIED_GROUP_STARTS, ignore_block=IGNORED_ALLIES_P55_BLOCK))
    records.extend(parse_group(german_df, "German", GERMAN_GROUP_STARTS))

    roster = pd.DataFrame([r.__dict__ for r in records])
    roster = roster.drop_duplicates(subset=["side", "formation", "unit_name"]).reset_index(drop=True)
    roster["status"] = "ACTIVE"
    roster["notes"] = ""
    roster["starting_losses"] = 0
    return roster


def import_roster(conn: sqlite3.Connection, roster_df: pd.DataFrame, replace_existing: bool = False) -> None:
    if replace_existing:
        conn.execute("DELETE FROM action_log")
        conn.execute("DELETE FROM unit_state")
        conn.execute("DELETE FROM units")
        conn.execute("UPDATE game_meta SET value='1' WHERE key='turn_number'")
        conn.commit()

    inserted = 0
    for _, row in roster_df.iterrows():
        cursor = conn.execute(
            """
            INSERT OR IGNORE INTO units (side, formation, unit_name, max_steps, armor_flag, status, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row["side"],
                row["formation"],
                row["unit_name"],
                int(row["max_steps"]),
                None if pd.isna(row["armor_flag"]) else int(bool(row["armor_flag"])),
                row.get("status", "ACTIVE"),
                row.get("notes", ""),
            ),
        )
        if cursor.rowcount:
            unit_id = conn.execute(
                "SELECT unit_id FROM units WHERE side=? AND formation=? AND unit_name=?",
                (row["side"], row["formation"], row["unit_name"]),
            ).fetchone()[0]
            conn.execute(
                "INSERT OR IGNORE INTO unit_state (unit_id, current_losses) VALUES (?, ?)",
                (unit_id, int(row.get("starting_losses", 0))),
            )
            inserted += 1
    conn.commit()


def get_turn_number(conn: sqlite3.Connection) -> int:
    row = conn.execute("SELECT value FROM game_meta WHERE key='turn_number'").fetchone()
    return int(row[0]) if row else 1


def set_turn_number(conn: sqlite3.Connection, turn_number: int) -> None:
    conn.execute("UPDATE game_meta SET value=? WHERE key='turn_number'", (str(turn_number),))
    conn.commit()


def fetch_roster(conn: sqlite3.Connection, side: Optional[str] = None, damaged_only: bool = False) -> pd.DataFrame:
    query = """
    SELECT
        u.unit_id,
        u.side,
        u.formation,
        u.unit_name,
        u.max_steps,
        COALESCE(s.current_losses, 0) AS current_losses,
        (u.max_steps - COALESCE(s.current_losses, 0)) AS effective_steps,
        u.status,
        COALESCE(u.armor_flag, 0) AS armor_flag,
        u.notes
    FROM units u
    LEFT JOIN unit_state s ON s.unit_id = u.unit_id
    WHERE 1=1
    """
    params: List = []

    if side and side != "All":
        query += " AND u.side = ?"
        params.append(side)
    if damaged_only:
        query += " AND COALESCE(s.current_losses, 0) > 0"

    query += " ORDER BY u.side, u.formation, u.unit_name"
    return pd.read_sql_query(query, conn, params=params)


def fetch_formations(conn: sqlite3.Connection, side: str) -> List[str]:
    rows = conn.execute(
        "SELECT DISTINCT formation FROM units WHERE side=? ORDER BY formation", (side,)
    ).fetchall()
    return [r[0] for r in rows]


def fetch_units_for_side(
    conn: sqlite3.Connection,
    side: str,
    formation: Optional[str] = None,
    damaged_only: bool = False,
    armor_filter: str = "all",
) -> pd.DataFrame:
    query = """
    SELECT
        u.unit_id,
        u.formation,
        u.unit_name,
        u.max_steps,
        COALESCE(s.current_losses, 0) AS current_losses,
        (u.max_steps - COALESCE(s.current_losses, 0)) AS effective_steps,
        u.status,
        COALESCE(u.armor_flag, 0) AS armor_flag
    FROM units u
    LEFT JOIN unit_state s ON s.unit_id = u.unit_id
    WHERE u.side=?
    """
    params: List = [side]

    if formation and formation != "All":
        query += " AND u.formation=?"
        params.append(formation)

    if damaged_only:
        query += " AND COALESCE(s.current_losses, 0) > 0"

    if armor_filter == "av":
        query += " AND COALESCE(u.armor_flag, 0) = 1"
    elif armor_filter == "non_av":
        query += " AND COALESCE(u.armor_flag, 0) = 0"

    query += " ORDER BY u.formation, u.unit_name"
    return pd.read_sql_query(query, conn, params=params)


def record_loss(conn: sqlite3.Connection, unit_id: int, amount: int, turn_number: int, note: str = "") -> Tuple[bool, str]:
    row = conn.execute(
        """
        SELECT u.side, u.max_steps, COALESCE(s.current_losses, 0) AS current_losses
        FROM units u
        LEFT JOIN unit_state s ON s.unit_id = u.unit_id
        WHERE u.unit_id=?
        """,
        (unit_id,),
    ).fetchone()
    if not row:
        return False, "Unit not found."

    max_steps = row["max_steps"]
    previous_losses = row["current_losses"]
    new_losses = previous_losses + amount
    if new_losses > max_steps:
        return False, f"That would exceed max steps ({max_steps})."

    conn.execute(
        "UPDATE unit_state SET current_losses=? WHERE unit_id=?",
        (new_losses, unit_id),
    )
    conn.execute(
        """
        INSERT INTO action_log (turn_number, side, action_type, unit_id, amount, previous_losses, new_losses, note)
        VALUES (?, ?, 'loss', ?, ?, ?, ?, ?)
        """,
        (turn_number, row["side"], unit_id, amount, previous_losses, new_losses, note),
    )
    conn.commit()
    return True, "Loss recorded."


def apply_replacement(conn: sqlite3.Connection, unit_id: int, amount: int, turn_number: int, note: str = "") -> Tuple[bool, str]:
    row = conn.execute(
        """
        SELECT u.side, COALESCE(s.current_losses, 0) AS current_losses
        FROM units u
        LEFT JOIN unit_state s ON s.unit_id = u.unit_id
        WHERE u.unit_id=?
        """,
        (unit_id,),
    ).fetchone()
    if not row:
        return False, "Unit not found."

    previous_losses = row["current_losses"]
    if amount > previous_losses:
        return False, "Replacement amount cannot exceed current losses."

    new_losses = previous_losses - amount
    conn.execute(
        "UPDATE unit_state SET current_losses=? WHERE unit_id=?",
        (new_losses, unit_id),
    )
    conn.execute(
        """
        INSERT INTO action_log (turn_number, side, action_type, unit_id, amount, previous_losses, new_losses, note)
        VALUES (?, ?, 'replacement', ?, ?, ?, ?, ?)
        """,
        (turn_number, row["side"], unit_id, amount, previous_losses, new_losses, note),
    )
    conn.commit()
    return True, "Replacement applied."


def fetch_history(conn: sqlite3.Connection, limit: int = 200) -> pd.DataFrame:
    query = """
    SELECT
        a.action_id,
        a.turn_number,
        a.side,
        a.action_type,
        u.formation,
        u.unit_name,
        a.amount,
        a.previous_losses,
        a.new_losses,
        a.note,
        a.created_at
    FROM action_log a
    JOIN units u ON u.unit_id = a.unit_id
    ORDER BY a.action_id DESC
    LIMIT ?
    """
    return pd.read_sql_query(query, conn, params=[limit])


def add_unit(conn: sqlite3.Connection, side: str, formation: str, unit_name: str, max_steps: int, status: str, note: str = "", armor_flag: bool = False) -> Tuple[bool, str]:
    try:
        conn.execute(
            """
            INSERT INTO units (side, formation, unit_name, max_steps, armor_flag, status, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (side, formation.strip(), unit_name.strip(), max_steps, int(bool(armor_flag)), status, note.strip()),
        )
        unit_id = conn.execute(
            "SELECT unit_id FROM units WHERE side=? AND formation=? AND unit_name=?",
            (side, formation.strip(), unit_name.strip()),
        ).fetchone()[0]
        conn.execute("INSERT INTO unit_state (unit_id, current_losses) VALUES (?, 0)", (unit_id,))
        conn.commit()
        return True, "Unit added."
    except sqlite3.IntegrityError:
        return False, "That unit already exists for this side and formation."


def main() -> None:
    st.set_page_config(page_title="BCS Replacement Tracker", layout="wide")
    st.title("BCS Last Blitzkrieg Tracker")
    st.caption("Version 1 starter app: roster import, loss entry, damaged-unit view, replacements, history, and manual unit add.")

    conn = get_connection()
    init_db(conn)

    with st.sidebar:
        st.header("Game")
        turn_number = get_turn_number(conn)
        new_turn = st.number_input("Current turn", min_value=1, value=turn_number, step=1)
        if new_turn != turn_number:
            set_turn_number(conn, int(new_turn))
            turn_number = int(new_turn)

        st.divider()
        st.header("Import roster")
        uploaded = st.file_uploader("Upload roster workbook (.xlsx)", type=["xlsx"])
        replace_existing = st.checkbox("Replace existing data on import", value=False)
        if st.button("Import workbook", use_container_width=True):
            try:
                if uploaded is not None:
                    roster_df = parse_workbook(uploaded)
                elif DEFAULT_IMPORT_PATH.exists():
                    with open(DEFAULT_IMPORT_PATH, "rb") as f:
                        roster_df = parse_workbook(io.BytesIO(f.read()))
                else:
                    raise FileNotFoundError("No uploaded workbook and no default workbook found in app folder.")

                import_roster(conn, roster_df, replace_existing=replace_existing)
                st.success(f"Imported {len(roster_df)} roster rows.")
            except Exception as exc:
                st.error(f"Import failed: {exc}")

        st.divider()
        st.header("Quick summary")
        roster_all = fetch_roster(conn)
        damaged_all = roster_all[roster_all["current_losses"] > 0]
        st.metric("Total units", len(roster_all))
        st.metric("Damaged units", len(damaged_all))
        st.metric("Total step losses", int(damaged_all["current_losses"].sum()) if not damaged_all.empty else 0)

    tab_dashboard, tab_roster, tab_losses, tab_replacements, tab_add, tab_history = st.tabs(
        ["Dashboard", "Roster", "Record Losses", "Replacement Planner", "Add Unit", "History"]
    )

    with tab_dashboard:
        roster_all = fetch_roster(conn)
        allies = roster_all[roster_all["side"] == "Allies"]
        german = roster_all[roster_all["side"] == "German"]
        allies_damaged = allies[allies["current_losses"] > 0]
        german_damaged = german[german["current_losses"] > 0]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Allied damaged units", len(allies_damaged))
        c2.metric("German damaged units", len(german_damaged))
        c3.metric("Allied losses", int(allies_damaged["current_losses"].sum()) if not allies_damaged.empty else 0)
        c4.metric("German losses", int(german_damaged["current_losses"].sum()) if not german_damaged.empty else 0)

        st.subheader("Current damaged units")
        side_filter = st.segmented_control("Side", ["All", "Allies", "German"], default="All", key="dash_side")
        damaged_view = fetch_roster(conn, side=None if side_filter == "All" else side_filter, damaged_only=True)
        st.dataframe(damaged_view, use_container_width=True, hide_index=True)

    with tab_roster:
        st.subheader("Master roster")
        side_filter = st.selectbox("Side", ["All", "Allies", "German"], key="roster_side")
        damaged_only = st.checkbox("Show damaged units only", value=False)
        roster_view = fetch_roster(conn, side=None if side_filter == "All" else side_filter, damaged_only=damaged_only)
        st.dataframe(roster_view, use_container_width=True, hide_index=True)

    with tab_losses:
        st.subheader("Record step losses")
        side = st.radio("Side", ["Allies", "German"], horizontal=True)
        formations = ["All"] + fetch_formations(conn, side)
        formation = st.selectbox("Formation", formations)
        units_df = fetch_units_for_side(
    conn,
    side,
    formation if formation != "All" else None,
    damaged_only=False,
    armor_filter="all",
)
        if units_df.empty:
            st.info("No units found for that filter.")
        else:
            unit_labels = {
                f"{row['formation']} — {row['unit_name']} (losses: {row['current_losses']}/{row['max_steps']})": int(row["unit_id"])
                for _, row in units_df.iterrows()
            }
            selected_label = st.selectbox("Unit", list(unit_labels.keys()))
            selected_id = unit_labels[selected_label]
            amount = st.number_input("Step loss to add", min_value=1, max_value=10, value=1, step=1)
            note = st.text_input("Note", value="")
            if st.button("Record loss", type="primary"):
                ok, message = record_loss(conn, selected_id, int(amount), turn_number, note)
                if ok:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)

    with tab_replacements:
        st.subheader("Replacement planner")
        side = st.radio("Replacement side", ["Allies", "German"], horizontal=True)

        c1, c2 = st.columns(2)
        with c1:
            av_points = st.number_input("AV replacement points available", min_value=0, value=0, step=1)
        with c2:
            non_av_points = st.number_input("Non-AV replacement points available", min_value=0, value=0, step=1)

        repl_type = st.radio("Replacement type", ["Non-AV", "AV"], horizontal=True)

        armor_filter = "av" if repl_type == "AV" else "non_av"
        available_points = av_points if repl_type == "AV" else non_av_points

        damaged_df = fetch_units_for_side(
            conn,
            side,
            damaged_only=True,
            armor_filter=armor_filter,
        )

    if damaged_df.empty:
        st.info(f"No damaged {repl_type} units for this side.")
    else:
        display_df = damaged_df[
            ["formation", "unit_name", "max_steps", "current_losses", "effective_steps", "armor_flag"]
        ].copy()

        display_df["armor_flag"] = display_df["armor_flag"].map({1: "AV", 0: "Non-AV"})

        st.dataframe(display_df, use_container_width=True, hide_index=True)

        unit_labels = {
            f"{row['formation']} — {row['unit_name']} (losses: {row['current_losses']})": int(row["unit_id"])
            for _, row in damaged_df.iterrows()
        }

        selected_label = st.selectbox("Damaged unit", list(unit_labels.keys()), key="repl_unit")
        selected_id = unit_labels[selected_label]

        current_losses = int(
            damaged_df.loc[damaged_df["unit_id"] == selected_id, "current_losses"].iloc[0]
        )

        amount = st.number_input(
            "Replacement amount",
            min_value=1,
            max_value=current_losses,
            value=1,
            step=1,
            key="repl_amount",
        )

        note = st.text_input("Replacement note", value="", key="repl_note")

        remaining_if_applied = available_points - int(amount)
        st.caption(f"{repl_type} points remaining after apply: {remaining_if_applied}")

        if st.button("Apply replacement", type="primary"):
            if amount > available_points:
                st.error(f"Replacement amount exceeds available {repl_type} points.")
            else:
                tagged_note = f"{repl_type} replacement" if not note else f"{repl_type} replacement | {note}"
                ok, message = apply_replacement(conn, selected_id, int(amount), turn_number, tagged_note)
                if ok:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)

    with tab_add:
        st.subheader("Add new unit")
        with st.form("add_unit_form"):
            side = st.selectbox("Side", ["Allies", "German"])
            formation = st.text_input("Formation")
            unit_name = st.text_input("Unit name")
            max_steps = st.number_input("Max steps", min_value=1, max_value=20, value=1, step=1)
            status = st.selectbox("Status", ["ACTIVE", "REINFORCEMENT"])
            armor_flag = st.checkbox("Uses AV replacements")
            note = st.text_input("Note")
            submitted = st.form_submit_button("Add unit")
            if submitted:
                if not formation.strip() or not unit_name.strip():
                    st.error("Formation and unit name are required.")
                else:
                    ok, message = add_unit(conn, side, formation, unit_name, int(max_steps), status, note, armor_flag)
                    if ok:
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)

    with tab_history:
        st.subheader("Recent actions")
        history_df = fetch_history(conn)
        if history_df.empty:
            st.info("No actions recorded yet.")
        else:
            st.dataframe(history_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
