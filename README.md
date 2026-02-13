## Excel District Splitter

This script splits an Excel file into multiple files based on unique values in a specified column.

### Configuration

The script is configured using the following variables:

- `BASE_DIR`: The base directory for the script.
- `TEMPLATE_FILE`: The path to the template Excel file.
- `SOURCE_DATA`: The path to the source Excel file.
- `OUTPUT_DIR`: The directory where the split files will be saved.
- `FILTER_COLUMN`: The column to use for filtering the data.

### Usage

To use the script, run it from the command line:

```bash
python main_v0.1/main.py

```

### Purpose

The script is designed to split an Excel file into multiple files based on unique values in a specified column.

### Output

The script will create a directory called `Split_Categories` in the same directory as the script. Each file in the directory will be named after the unique value in the specified column.