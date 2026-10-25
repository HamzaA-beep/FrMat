# FrMat components and workflow

FrMat is made up of a few essential parts that work together to create a local data review and cleanup workflow.

## Main components

### File import layer
This is the entry point for loading data files.

Supported types include:

- CSV
- XLS / XLSX
- TXT
- JSON
- XML
- PDF
- DOC / DOCX
- SQL / SQLite

### Table view
The imported data is displayed in a table so the user can inspect columns, values, and row counts before applying transformations.

### Filter engine
The filter system uses a column-based syntax such as:

```text
[age] >= 18 AND [region] = West
[city] ~ York
[status] != null
```

This helps users narrow the dataset without learning a full query language.

### Cleanup presets
FrMat includes simple operations such as:

- removing duplicate rows
- dropping rows with missing values
- removing IQR-based outliers

### Chart generation
Users can quickly produce basic charts such as:

- histogram
- bar chart
- line chart
- scatter plot

### Export layer
Processed datasets can be saved in formats such as:

- CSV
- XLSX
- JSON
- XML
- SQLite database
- DOCX
- PDF

## Why these pieces matter

Each part of FrMat is built to support a simple local-analysis pipeline:

1. import data
2. inspect and filter it
3. clean it
4. visualize patterns
5. export the output

## Summary

FrMat combines import, cleaning, filtering, visualization, and export in a single local tool so the user can work without a larger platform.
