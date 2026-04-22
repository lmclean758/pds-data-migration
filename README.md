# PDS File Scanner

Scans a folder of PDS Excel files and populates a tracker spreadsheet (`PDS_TRACKER_FILE.xlsx`).

Each non-empty part number in a PDS file gets its own row in the tracker. Re-running the script is safe — existing rows are skipped.

## Requirements

- Python 3.10+
- [openpyxl](https://pypi.org/project/openpyxl/)

```
pip install openpyxl
```

## Usage

```
python scan_pds.py <PDS_FOLDER> <TRACKER_FILE>
```

| Argument | Description |
|---|---|
| `PDS_FOLDER` | Path to the folder containing your PDS `.xlsx` files |
| `TRACKER_FILE` | Path to the tracker `.xlsx` template |

**Example:**

```
python scan_pds.py "C:\Users\YourName\Desktop\PDS_FILES" "C:\Users\YourName\Desktop\PDS_TRACKER_FILE.xlsx"
```

The output file (`PDS_OUTPUT_<timestamp>.xlsx`) is saved to the same folder as the tracker.
