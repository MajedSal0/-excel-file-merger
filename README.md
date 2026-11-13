# Excel & CSV File Merger

Upload multiple `.xlsx` and `.csv` files, align columns by header, keep the source filename for every row, preview the result, and download one formatted Excel file.

## Run

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Notes
- First worksheet only for Excel files.
- The first row is treated as the header.
- Different column names create separate columns; missing values stay blank.
- CSV decoding tries UTF-8 and then Windows Arabic (`cp1256`).
- Uploaded data is processed in the running app; no database is used.
