DATA CONSOLIDATION UTILITY - UPDATED

1. Extract this folder inside your EDI project root.
2. Put company dataset files inside input/, preferably in separate company folders.
3. Run consolidate_company_datasets.py using VS Code's Run Python File button.
4. The generated workbook will be saved in output/.

Updates in this version:
- All data cells use consistent left alignment. Numeric values remain numeric.
- Freeze panes set to F2: the header row and first five provenance columns stay visible while scrolling.
- Consolidation_Record_ID is now human-readable and avoids long relative file paths.
- Nested-array child IDs use the same readable parent ID and a related-item suffix.
- Existing worksheet filters remain; no structured Excel Tables are created.

No existing EDI source files or input datasets are modified.
