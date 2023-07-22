# csv-file-diff-checker

## Excel File Comparator

This Python script allows you to compare two Excel files and identify the differences between corresponding cells in the matching sheets. The script uses the `openpyxl` library to load and access the Excel files.

## Prerequisites

Before running the script, ensure you have the following installed:

- Python 3.x
- `openpyxl` library

You can install the required library using the following command:

```bash
pip install openpyxl
```

## How to Use

1. Place your old Excel file and new Excel file in the same directory as the script.
2. Provide the filenames of the old and new Excel files in the `old_file` and `new_file` variables, respectively.
3. Run the script, and it will compare the two files sheet by sheet, identifying the differences between corresponding cells.
4. The script will create a `Differences.json` file in the same directory containing the details of the differences found.

## Example

```python
import json
from openpyxl.styles import Font
import openpyxl
from openpyxl.utils import get_column_letter

# Set the filenames for the old and new Excel files
old_file = "old_file.xlsx"
new_file = "new_file.xlsx"

# Rest of the script...
```

Please ensure that you have a proper backup of your Excel files before running this script, as it may modify the data in the Excel files.

If you encounter any issues or have questions, feel free to reach out or refer to the documentation of the `openpyxl` library for more information. Happy comparing!