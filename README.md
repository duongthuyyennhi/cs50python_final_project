# Excel Chart Viewer

#### Video Demo: <URL HERE>

#### Description:

Excel Chart Viewer is a beginner-friendly desktop application for exploring
financial and other tabular data stored in an Excel workbook. It was designed
for exported market data, where a price may be stored as text such as
`60 +0,4 ( +0,01 %)` and large values may use dots as thousands separators.
The application lets a user inspect the original spreadsheet rows and turn
selected columns into a chart without copying data into another program.

After opening an `.xlsx` or `.xlsm` file, the application displays the first
ten rows as a preview. The preview intentionally shows the original cell
values, so the user can compare the chart with the source data. The first
worksheet is used because this keeps the workflow predictable for a beginner.
The workbook is opened in read-only mode, which reduces memory use when the
file contains many rows.

The user can choose `Scatter` or `Line` from the chart menu and independently
choose an `X column` and a `Y column`. A Line chart is useful for viewing a
value over time, while a Scatter chart is useful for comparing two variables,
such as price and trading volume. The program prevents choosing the same
column for both axes. Numeric text is converted internally for plotting, but
the source preview is not changed. Values such as `1.014.107.000` are read as
`1014107000`, while decimal commas are also supported.

Example input data:

| Time | Price | Trading volume | Cumulative volume | Weight |
| --- | --- | ---: | ---: | ---: |
| 09:15:59 | 60 +0,4 ( +0,01 %) | 20.000 | 824.973.000 | 0,0% |
| 09:15:59 | 60 +0,4 ( +0,01 %) | 8.400.000 | 833.373.000 | 0,01% |
| 09:15:59 | 60 +0,4 ( +0,01 %) | 20.000 | 824.973.000 | 0,0% |
| 09:15:59 | 60 +0,4 ( +0,01 %) | 80.000 | 829.773.000 | 0,0% |
| 09:15:59 | 60 +0,4 ( +0,01 %) | 1.200.000 | 830.973.000 | 0,0% |

The preview keeps values in this original form. When the user selects
`Price` and `Trading volume`, the application converts the numeric parts only
for plotting, while the table remains a faithful preview of the Excel cells.

When either selected column contains times, the `From` and `To` fields become
available. They accept `HH:MM` or `HH:MM:SS`; an empty field means that side of
the range is unlimited. The `Group` menu can keep every row or group time into
30-second, 1-minute, 5-minute, 15-minute, or 1-hour buckets. For each bucket,
the last observed Y value is retained. Filtering is performed before grouping,
which makes it possible to focus on a short trading period and then reduce a
dense chart.

The chart includes a mouse hover tooltip showing the X and Y values near a
data point. Matplotlib's navigation toolbar provides zoom, pan, home, back,
and forward controls. Line charts use a thin line without markers so large
datasets remain readable; Scatter charts keep individual points for examining
the relationship between columns.

`project.py` contains the GUI and application logic. `parse_number` handles
the fixed numeric format, `load_excel` reads the first worksheet,
`filter_by_time` limits rows, `aggregate_by_time` reduces time-series data,
and `create_chart` builds the Matplotlib figure. `test_project.py` contains
tests for numeric parsing, time parsing, filtering, grouping, workbook
loading, chart creation, axis formatting, and the hover tooltip. The tests
use a temporary workbook and Matplotlib's non-GUI backend. `requirements.txt`
lists the project's pip-installable dependencies: `matplotlib`, `openpyxl`,
and `pytest`.

The main design choice was to use Tkinter, openpyxl, and Matplotlib instead of
a larger data framework. Tkinter is included with Python, openpyxl is enough
to read the required Excel files, and Matplotlib provides chart navigation and
formatting. This keeps the code understandable for a new Python student while
still supporting a workbook with thousands of rows. The application currently
reads only the first worksheet and supports time values within one day; these
limitations keep the implementation focused and can be extended later.

To install dependencies and run the application from this folder:

```text
pip install -r requirements.txt
python project.py
```

To run the automated tests:

```text
pytest
```