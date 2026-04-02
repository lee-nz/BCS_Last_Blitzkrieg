# BCS Last Blitzkrieg Tracker

A lightweight web app for tracking unit losses and replacements in **Battalion Combat Series (BCS) – Last Blitzkrieg**.

Built using Streamlit, this tool replaces manual tracking and spreadsheets with a faster, structured workflow for managing step losses and replacements during gameplay.

---

## Features

- Built-in roster included (ready to use)
- Optional upload of custom roster files
- Track step losses by division and unit
- View all damaged units in one place
- Apply replacements (separate AV and Non-AV pools)
- Automatic calculation of effective steps
- Turn tracking
- Action history log

---

## Quick Start

### Option 1 — Use built-in roster (recommended)

1. Open the app  
2. Click **Import workbook**  
3. Start tracking immediately  

---

### Option 2 — Upload your own roster

1. Click **Upload roster workbook**  
2. Select your `.xlsx` file  
3. Click **Import workbook**

---

## Excel Format (for custom rosters)

Your Excel file must contain two sheets:

### `Allied Units`
### `German Units`

Each sheet must include the following columns:

| Column       | Description                  |
|-------------|------------------------------|
| Formation   | Division or formation name   |
| Unit        | Unit name                    |
| Max. Steps  | Maximum step value           |
| AV Flag     | 1 = AV unit, 0 = Non-AV      |
| Notes       | Optional                     |

⚠️ Column names must match exactly.

---

## Running Locally

Install dependencies:

```bash
pip install streamlit pandas openpyxl
