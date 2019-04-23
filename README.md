# CSDCO Collection Generator

This app takes information from the CSDCO SQLite database and converts the boreholes table into three variants (CSV, HTML, and KML) that are then published on our website. Each file has a different set of columns that are included.

It has one required input, the path to the sqlite database, and one optional flag (```-d```) which will export filenames with a \_YYYYMMDD suffix. Files are exported to the same directory the script is run from.

## Screenshot/usage
<img src="https://user-images.githubusercontent.com/6476269/56608831-49120a00-65d1-11e9-9c83-27ac284bf5ac.png" width="80%" alt="Screenshot of the CSDCO Collection Generator" />
