# KETKAT Hook-up & Installation Manager - Streamlit MVP V3

## Run locally

```bash
py -m pip install -r requirements.txt
py -m streamlit run streamlit_app.py
```

## New in V3

- Light professional UI.
- Optional logo preview in sidebar.
- Export Center page.
- Export Database Summary.
- Export Last BOQ Selection.
- BOQ item selection and send to Rate Analysis.
- Rate Analysis resource table with Excel export.

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