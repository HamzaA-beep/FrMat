# FrMat troubleshooting

This page covers the most common issues users may encounter when running FrMat locally.

## App does not start

Check the following:

- Python is installed and available on PATH
- the environment is activated
- `requirements.txt` has been installed successfully
- the current project directory contains `data_workbench.py`

## Missing dependency errors

Reinstall the dependencies:

```powershell
python -m pip install -r requirements.txt
```

If a package fails to load, verify the Python version matches the supported range.

## Import fails for a file type

Common causes include:

- unsupported extension
- unreadable or corrupted file
- missing tool such as LibreOffice for `.doc` conversion
- insufficient file permissions

## Filter produces no rows

Check the syntax carefully. FrMat expects expressions in the form:

```text
[column] operator value
```

For example:

```text
[age] >= 18
[status] != null
```

## Chart does not render

This usually happens when:

- the selected column contains non-numeric data for numeric charting
- the dataset is empty
- the chart type doesn’t match the selected columns

## Export fails

Verify:

- the output path is writable
- the filename extension matches the export format
- the data is not too wide for the chosen PDF output

## Summary

Most FrMat issues are caused by dependency problems, incorrect file formats, or invalid filter expressions. Checking the file type and Python environment usually resolves the majority of problems quickly.
