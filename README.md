# PDS Data Migration

Scans a folder of PDS Excel files and populates a tracker spreadsheet (`PDS_TRACKER_FILE.xlsx`).

Each non-empty part number in a PDS file gets its own row in the tracker. Re-running the script is safe — existing rows are skipped.

## Getting the Code

**Option A — Download ZIP** *(recommended for most users)*

1. Click the green **Code** button on the GitHub page
2. Select **Download ZIP**
3. Extract the folder anywhere on your machine (use C:\ drive for safety)

**Option B — Clone with Git**

```
git clone https://github.com/lmclean758/pds-data-migration.git
```

## Requirements

- Python 3.10+ — [download here](https://www.python.org/downloads/)
- openpyxl

Install openpyxl by opening **Command Prompt** or **PowerShell** (not the Python interpreter) and running:

```
pip install openpyxl
```

> **Note:** Make sure you see a regular terminal prompt (e.g. `C:\Users\YourName>`) and not the Python prompt (`>>>`) before running this command.

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

## License

MIT License — free to use, modify, and distribute.
