# Installation and Usage Guide for Spreadsheet Splitter

`spreadsheet_splitter` is a Python command-line tool for splitting large Excel (`.xls` or `.xlsx`) files into smaller parts with low memory usage. This guide explains how to install the tool via PyPI and use it on Linux, macOS, and Windows.

## Installation

### Prerequisites
- **Python**: Version 3.13 or higher. Download from [python.org](https://www.python.org/downloads/).
- An Excel file (`.xls` or `.xlsx`) to process.

### Install via PyPI
The recommended way to install `spreadsheet_splitter` is using `pip`:

```bash
pip install spreadsheet-splitter
```

This installs the tool and its dependencies (`openpyxl`, `pandas`, `psutil`, `pyexcelerate`, `xlrd`, `xlwt`).

### Verify Installation
Check that the tool is installed:
```bash
spreadsheet-splitter --help
```
This displays the CLI help with available arguments.

## Usage
The tool supports two modes:
- **Default Mode**: Splits an Excel file into two equal parts, preserving title rows.
- **Large Mode**: Splits large files into smaller chunks with a specified number of rows per iteration.

### Example Commands
1. **Split a File into Two Parts**
   ```bash
   spreadsheet-splitter /path/to/file.xlsx --output-dir split_files --title-rows 1
   ```
   - Splits `file.xlsx` into `split_files/file_part1.xlsx` and `split_files/file_part2.xlsx`.
   - Preserves 1 title row.
   - On Windows, use backslashes:
     ```cmd
     spreadsheet-splitter C:\path\to\file.xlsx --output-dir split_files --title-rows 1
     ```

2. **Split a Large File Iteratively**
   ```bash
   spreadsheet-splitter /path/to/very_large_file.xlsx --large --output-dir split_files --rows-per-iteration 500 --max-iterations 5 --columns 22
   ```
   - Processes `very_large_file.xlsx` in chunks of 500 rows, up to 5 iterations.
   - Reads only the first 22 columns.
   - Outputs files like `split_files/very_large_file_part001.xlsx`.
   - Logs progress to `split_log.json`.

3. **Resume Processing**
   ```bash
   spreadsheet-splitter /path/to/very_large_file.xlsx --large --resume --output-dir split_files --rows-per-iteration 500 --max-iterations 5 --columns 22
   ```
   - Resumes from the last row in `split_log.json`.

### Command-Line Arguments
| Argument | Description | Default |
|----------|-------------|---------|
| `input_file` | Path to the input `.xls` or `.xlsx` file | Required |
| `-o, --output-dir` | Directory to save output files | `output` |
| `-t, --title-rows` | Number of top rows to treat as title rows (ignored in `--large` mode) | 1 |
| `--large` | Process large files iteratively | False |
| `--rows-per-iteration` | Number of rows per iteration | 500 |
| `--max-iterations` | Maximum number of iterations | 10 |
| `--resume` | Resume from the last row | False |
| `--log-file` | Path to the log file | `split_log.json` |
| `--columns` | Number of columns to process | All columns |

### Log File
The tool generates `split_log.json` with details like:
- Input file and output directory.
- Processed row ranges (e.g., `start_row`, `end_row`).
- File profile (rows, columns, data types for `.xlsx`).
- Memory usage per chunk.

Example:
```json
{
  "input_path": "/path/to/very_large_file.xlsx",
  "output_dir": "split_files",
  "rows_per_iteration": 500,
  "processed_ranges": [
    {
      "file": "split_files/very_large_file_part001.xlsx",
      "start_row": 1,
      "end_row": 500,
      "iteration": 1
    }
  ],
  "last_row_processed": 499,
  "timestamp": "2025-05-27T11:12:46.172255",
  "columns": 22,
  "file_profile": {
    "total_rows": 1000,
    "total_columns": 22,
    "data_types": {
      "ZONE": "float64",
      "SITE": "object",
      ...
    }
  },
  "memory_usage": [
    {
      "iteration": 1,
      "memory_before_mb": 256.55,
      "memory_after_mb": 259.63
    }
  ]
}
```

## Platform-Specific Notes
- **Linux/macOS**:
  - Use forward slashes in file paths (e.g., `/path/to/file.xlsx`).
  - Install Python via package managers (`apt`, `brew`) or [python.org](https://www.python.org).
  - Ensure `pip` is updated: `pip install --upgrade pip`.

- **Windows**:
  - Use backslashes in file paths (e.g., `C:\path\to\file.xlsx`) or raw strings (`r"C:\path\to\file.xlsx"`).
  - Install Python from [python.org](https://www.python.org) or the Microsoft Store.
  - Run commands in Command Prompt, PowerShell, or Windows Terminal.
  - Ensure `pip` is in your PATH. If not, use:
    ```cmd
    python -m pip install spreadsheet-splitter
    ```

## Troubleshooting
- **Command Not Found**: Ensure `spreadsheet-splitter` is installed and `pip`’s `Scripts` directory is in your PATH:
  ```bash
  export PATH=$PATH:~/.local/bin  # Linux/macOS
  set PATH=%PATH%;%USERPROFILE%\AppData\Local\Programs\Python\Python313\Scripts  # Windows
  ```

- **Column Count Error**: Verify the column count:
  ```python
  import pandas as pd
  df = pd.read_excel("very_large_file.xlsx", nrows=1, engine='openpyxl')
  print(df.shape[1], "columns")
  ```

- **Memory Issues**: Reduce `--rows-per-iteration` (e.g., to 100). Check `split_log.json` for memory usage.

- **Slow Processing**: Use a fast drive (e.g., SSD). For `.xlsx` files with many strings, convert to `.csv`:
  ```python
  import pandas as pd
  df = pd.read_excel("very_large_file.xlsx", engine='openpyxl')
  df.to_csv("very_large_file.csv", index=False)
  ```

## Support
- Report issues at [GitHub Issues](https://github.com/parsaloi/spreadsheet_splitter/issues).
- See the [README](https://github.com/parsaloi/spreadsheet_splitter/blob/main/README.md) for building from source.
