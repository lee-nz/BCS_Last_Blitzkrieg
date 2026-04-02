import io
import sqlite3
from pathlib import Path
from typing import List, Optional

import pandas as pd
import streamlit as st

DB_PATH = Path("bcs_tracker.db")
DEFAULT_IMPORT_PATH = Path("Last Blitzkrieg Step Loses.xlsx")

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS units (
    unit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    side TEXT NOT NULL,
    formation TEXT NOT NULL,
    unit_name TEXT NOT NULL,
    max_steps INTEGER NOT NULL,
    armor_flag INTEGER NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'ACTIVE',
    notes TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS unit_state (
    unit_id INTEGER PRIMARY KEY,
    current_losses INTEGER NOT NULL DEFAULT 0
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
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
"""


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(SCHEMA_SQL)
    conn.execute("INSERT OR IGNORE INTO game_meta (key, value) VALUES ('turn_number', '1')")
    conn.commit()
    ensure_unit_state_rows(conn)


def ensure_unit_state_rows(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        INSERT OR IGNORE INTO unit_state (unit_id, current_losses)
        SELECT unit_id, 0 FROM units
        """
    )
    conn.commit()


def parse_bool_flag(value) -> int:
    if value is None or pd.isna(value):
        return 0
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y"}:
        return 1
    try:
        return 1 if int(float(value)) != 0 else 0
    except Exception:
        return 0


def parse_unit_sheet(df: pd.DataFrame, side: str) -> pd.DataFrame:
    df = df.rename(
        columns={
            "Formation": "formation",
            "Unit": "unit_name",
            "Max. Steps": "max_steps",
            "AV Flag": "armor_flag",
            "Notes": "notes",
        }
    ).copy()

    required_cols = ["formation", "unit_name", "max_steps"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column '{col}' on {side} sheet.")

    if "armor_flag" not in df.columns:
        df["armor_flag"] = 0
    if "notes" not in df.columns:
        df["notes"] = ""

    df = df[["formation", "unit_name", "max_steps", "armor_flag", "notes"]].copy()

    df["formation"] = df["formation"].fillna("").astype(str).str.strip()
    df["unit_name"] = df["unit_name"].fillna("").astype(str).str.strip()
    df["notes"] = df["notes"].fillna("").astype(str).str.strip()

    df = df[(df["formation"] != "") & (df["unit_name"] != "")].copy()

    df["max_steps"] = pd.to_numeric(df["max_steps"], errors="coerce")
    df = df[df["max_steps"].notna()].copy()
    df["max_steps"] = df["max_steps"].astype(int)

    df["armor_flag"] = df["armor_flag"].apply(parse_bool_flag).astype(int)
    df["side"] = side
    df["status"] = "ACTIVE"

    return df[["side", "formation", "unit_name", "max_steps", "armor_flag", "status", "notes"]]


def parse_workbook(file_obj) -> pd.DataFrame:
    allied_df = pd.read_excel(file_obj, sheet_name="Allied Units")
    german_df = pd.read_excel(file_obj, sheet_name="German Units")

    allied_units = parse_unit_sheet(allied_df, "Allies")
    german_units = parse_unit_sheet(german_df, "German")

    roster_df = pd.concat([allied_units, german_units], ignore_index=True).reset_index(drop=True)
    return roster_df


def import_roster(conn: sqlite3.Connection, roster_df: pd.DataFrame, replace_existing: bool = False) -> int:
    if replace_existing:
        conn.execute("DELETE FROM action_log")
        conn.execute("DELETE FROM unit_state")
        conn.execute("DELETE FROM units")
        conn.execute("UPDATE game_meta SET value='1' WHERE key='turn_number'")
        conn.commit()

    for _, row in roster_df.iterrows():
        conn.execute(
            """
            INSERT INTO units (side, formation, unit_name, max_steps, armor_flag, status, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row["side"],
                row["formation"],
                row["unit_name"],
                int(row["max_steps"]),
                int(row.get("armor_flag", 0)),
                row.get("status", "ACTIVE"),
                row.get("notes", ""),
            ),
        )

    conn.commit()
    ensure_unit_state_rows(conn)
    count = conn.execute("SELECT COUNT(*) FROM units").fetchone()[0]
    return int(count)


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
        CASE WHEN COALESCE(u.armor_flag, 0) = 1 THEN 'AV' ELSE 'Non-AV' END AS replacement_type,
        u.status,
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

    query += " ORDER BY u.side, u.formation, u.unit_name, u.unit_id"
    return pd.read_sql_query(query, conn, params=params)


def fetch_formations(conn: sqlite3.Connection, side: str) -> List[str]:
    rows = conn.execute(
        "SELECT DISTINCT formation FROM units WHERE side=? ORDER BY formation",
        (side,),
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
        u.side,
        u.formation,
        u.unit_name,
        u.max_steps,
        COALESCE(s.current_losses, 0) AS current_losses,
        (u.max_steps - COALESCE(s.current_losses, 0)) AS effective_steps,
        COALESCE(u.armor_flag, 0) AS armor_flag,
        u.status,
        u.notes
    FROM units u
    LEFT JOIN unit_state s ON s.unit_id = u.unit_id
    WHERE u.side = ?
    """
    params: List = [side]

    if formation and formation != "All":
        query += " AND u.formation = ?"
        params.append(formation)

    if damaged_only:
        query += " AND COALESCE(s.current_losses, 0) > 0"

    if armor_filter == "av":
        query += " AND COALESCE(u.armor_flag, 0) = 1"
    elif armor_filter == "non_av":
        query += " AND COALESCE(u.armor_flag, 0) = 0"

    query += " ORDER BY u.formation, u.unit_name, u.unit_id"
    return pd.read_sql_query(query, conn, params=params)


def update_loss(conn: sqlite3.Connection, unit_id: int, delta: int, turn_number: int, note: str = "") -> tuple[bool, str]:
    conn.execute("INSERT OR IGNORE INTO unit_state (unit_id, current_losses) VALUES (?, 0)", (int(unit_id),))

    row = conn.execute(
        """
        SELECT u.side, u.max_steps, COALESCE(s.current_losses, 0) AS current_losses
        FROM units u
        LEFT JOIN unit_state s ON s.unit_id = u.unit_id
        WHERE u.unit_id = ?
        """,
        (int(unit_id),),
    ).fetchone()

    if row is None:
        return False, "Unit not found."

    previous_losses = int(row["current_losses"])
    new_losses = previous_losses + int(delta)

    if new_losses < 0:
        return False, "Losses cannot go below zero."
    if new_losses > int(row["max_steps"]):
        return False, f"That would exceed max steps ({row['max_steps']})."

    conn.execute(
        "UPDATE unit_state SET current_losses = ? WHERE unit_id = ?",
        (new_losses, int(unit_id)),
    )

    action_type = "loss" if delta >= 0 else "replacement"
    conn.execute(
        """
        INSERT INTO action_log (turn_number, side, action_type, unit_id, amount, previous_losses, new_losses, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            int(turn_number),
            row["side"],
            action_type,
            int(unit_id),
            abs(int(delta)),
            previous_losses,
            new_losses,
            note,
        ),
    )
    conn.commit()
    return True, "Updated."


def apply_replacement(conn: sqlite3.Connection, unit_id: int, amount: int, turn_number: int, note: str = "") -> tuple[bool, str]:
    return update_loss(conn, unit_id, -abs(int(amount)), turn_number, note)


def fetch_history(conn: sqlite3.Connection, limit: int = 200) -> pd.DataFrame:
    query = """
    SELECT
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


def add_unit(
    conn: sqlite3.Connection,
    side: str,
    formation: str,
    unit_name: str,
    max_steps: int,
    status: str,
    armor_flag: bool,
    note: str = "",
) -> tuple[bool, str]:
    conn.execute(
        """
        INSERT INTO units (side, formation, unit_name, max_steps, armor_flag, status, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            side,
            formation.strip(),
            unit_name.strip(),
            int(max_steps),
            1 if armor_flag else 0,
            status,
            note.strip(),
        ),
    )
    conn.commit()
    ensure_unit_state_rows(conn)
    return True, "Unit added."


def main() -> None:
    st.set_page_config(page_title="BCS Last Blitzkrieg Tracker", layout="wide")
    st.title("BCS Last Blitzkrieg Tracker")
    st.caption("Roster import, division-based loss entry, damaged-unit view, AV and Non-AV replacements, history, and manual unit add.")

    conn = get_connection()
    init_db(conn)

    with st.sidebar:
        st.header("Game")
        turn_number = get_turn_number(conn)
        turn_input = st.number_input("Current turn", min_value=1, value=turn_number, step=1)
        if int(turn_input) != turn_number:
            set_turn_number(conn, int(turn_input))
            turn_number = int(turn_input)

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

                total_units = import_roster(conn, roster_df, replace_existing=replace_existing)
                st.success(f"Import complete. Database now holds {total_units} units.")
            except Exception as exc:
                st.error(f"Import failed: {exc}")

        st.divider()
        roster_all = fetch_roster(conn)
        damaged_all = roster_all[roster_all["current_losses"] > 0]
        st.header("Quick summary")
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
        side_filter = st.selectbox("Side", ["All", "Allies", "German"], key="dash_side")
        damaged_view = fetch_roster(conn, None if side_filter == "All" else side_filter, damaged_only=True)
        st.dataframe(damaged_view, use_container_width=True, hide_index=True)

    with tab_roster:
        st.subheader("Master roster")
        side_filter = st.selectbox("Side", ["All", "Allies", "German"], key="roster_side")
        damaged_only = st.checkbox("Show damaged units only", value=False)
        roster_view = fetch_roster(conn, None if side_filter == "All" else side_filter, damaged_only=damaged_only)
        st.dataframe(roster_view, use_container_width=True, hide_index=True)

    with tab_losses:
        st.subheader("Record step losses")

        side = st.radio("Side", ["Allies", "German"], horizontal=True, key="loss_side")

        formations = fetch_formations(conn, side)
        if not formations:
            st.info("No units found. Import the workbook first.")
        else:
            formation = st.selectbox("Formation / Division", formations, key="loss_formation")

            units_df = fetch_units_for_side(
                conn,
                side,
                formation=formation,
                damaged_only=False,
                armor_filter="all",
            )

            if units_df.empty:
                st.info("No units found for that formation.")
            else:
                st.markdown("### Units")

                for _, row in units_df.iterrows():
                    unit_id = int(row["unit_id"])
                    unit_name = row["unit_name"]
                    losses = int(row["current_losses"])
                    max_steps = int(row["max_steps"])

                    col1, col2, col3, col4 = st.columns([3, 2, 1, 1])

                    with col1:
                        st.write(unit_name)

                    with col2:
                        st.write(f"{losses} / {max_steps}")

                    with col3:
                        if st.button("+1", key=f"add_{unit_id}"):
                            ok, message = update_loss(conn, unit_id, 1, turn_number, "Quick add")
                            if ok:
                                st.rerun()
                            else:
                                st.error(message)

                    with col4:
                        if st.button("-1", key=f"remove_{unit_id}"):
                            ok, message = update_loss(conn, unit_id, -1, turn_number, "Quick remove")
                            if ok:
                                st.rerun()
                            else:
                                st.error(message)

    with tab_replacements:
        st.subheader("Replacement planner")
        side = st.radio("Replacement side", ["Allies", "German"], horizontal=True, key="repl_side")

        c1, c2 = st.columns(2)
        with c1:
            av_points = st.number_input("AV replacement points available", min_value=0, value=0, step=1)
        with c2:
            non_av_points = st.number_input("Non-AV replacement points available", min_value=0, value=0, step=1)

        repl_type = st.radio("Replacement type", ["Non-AV", "AV"], horizontal=True)
        armor_filter = "av" if repl_type == "AV" else "non_av"
        available_points = av_points if repl_type == "AV" else non_av_points

        damaged_df = fetch_units_for_side(conn, side, damaged_only=True, armor_filter=armor_filter)
        if damaged_df.empty:
            st.info(f"No damaged {repl_type} units for this side.")
        else:
            display_df = damaged_df[["formation", "unit_name", "max_steps", "current_losses", "effective_steps", "armor_flag"]].copy()
            display_df["armor_flag"] = display_df["armor_flag"].map({1: "AV", 0: "Non-AV"})
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            unit_labels = {
                f"{row['formation']} — {row['unit_name']} (losses: {row['current_losses']})": int(row["unit_id"])
                for _, row in damaged_df.iterrows()
            }
            selected_label = st.selectbox("Damaged unit", list(unit_labels.keys()), key="repl_unit")
            selected_id = unit_labels[selected_label]
            current_losses = int(damaged_df.loc[damaged_df["unit_id"] == selected_id, "current_losses"].iloc[0])

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
                if int(amount) > int(available_points):
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
            formation = st.text_input("Formation / Division")
            unit_name = st.text_input("Unit name")
            max_steps = st.number_input("Max steps", min_value=1, max_value=20, value=1, step=1)
            status = st.selectbox("Status", ["ACTIVE", "REINFORCEMENT"])
            armor_flag = st.checkbox("Uses AV replacements")
            note = st.text_input("Note")
            submitted = st.form_submit_button("Add unit")
            if submitted:
                try:
                    ok, message = add_unit(conn, side, formation, unit_name, int(max_steps), status, armor_flag, note)
                    if ok:
                        st.success(message)
                        st.rerun()
                except Exception as exc:
                    st.error(str(exc))

    with tab_history:
        st.subheader("Recent actions")
        history_df = fetch_history(conn)
        if history_df.empty:
            st.info("No actions recorded yet.")
        else:
            st.dataframe(history_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
