# Excel Cell Splitter
In some instances, it may be necessary to work with an Excel worksheet that contains cells with multiple values. These values can be separated by commas or newlines, which can make basic data analysis tasks like filtering and creating pivot tables difficult.

This Python script takes an Excel file as input and convert cells with comma or newline separated values into multiple rows. Additionally, the script ensures that any unaffected columns are preserved, maintaining the original data's integrity.


## Dependencies
Dependencies can be installed via `pip` or `pip3`.

```
pip3 install pandas openpyxl
```

1. pandas
2. openpyxl

## Usage
```
python splitter.py
```
