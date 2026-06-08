# KETKAT Hook-up & Installation Manager - V4 Internal Database

## What is new in V4?

This version does **not** ask users to upload the main BOQ database.

The app reads the database automatically from:

```txt
data/boq_data.xlsx
```

## Run locally

```bash
py -m pip install -r requirements.txt
py -m streamlit run streamlit_app.py
```

## Required database path

Create a folder called:

```txt
data
```

Then place your Excel database inside it:

```txt
data/boq_data.xlsx
```

## Required Excel columns

- Main Category
- Sub System
- Family
- Work Type
- Title
- No
- Description
- Unit
- Default Qty
- Unit Price

Recommended sheet name: `BOQ_Data`.

If `BOQ_Data` is not found, the app will read the first sheet.

## How to update the live website database

1. Edit `boq_data.xlsx` on your computer.
2. Upload the updated file to GitHub in:

```txt
data/boq_data.xlsx
```

3. Commit changes.
4. Wait for Streamlit to update, or open Manage app and click Reboot.

## Security note

If the GitHub repository is public, anyone can access the Excel database file from GitHub.

For real prices or private company data, use a private repository or keep only demo data in the public version.