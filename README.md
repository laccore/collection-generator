# CSDCO Collection Generator

This script takes information from the CSDCO SQLite database and converts the boreholes table into three variants (CSV, HTML, and KML) that are then published on our website. Each file has a different set of columns that are included.

It has one required input, the path to the sqlite database, and one optional flag (```-d```) which will export filenames without a \_YYYYMMDD suffix. Files are exported to the same directory the script is run from.

### Example usage
```python
python csdco-collection-generator.py /path/to/csdco_database.sqlite3 -d
```
