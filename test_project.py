import matplotlib

matplotlib.use("Agg")

import openpyxl

from project import (aggregate_by_time, create_chart, format_axis_number,
                     filter_by_time, format_time_axis, is_time_column, load_excel,
                     parse_number, parse_time_value)


def test_parse_number_uses_fixed_excel_format():
    assert parse_number("60 +0,4 ( +0,01 %)") == 60.0
    assert parse_number("1.014.107.000") == 1014107000.0
    assert parse_number("60,4") == 60.4
    assert parse_number("VCB") is None


def test_chart_converts_unformatted_thousands_values():
    figure = create_chart(["60,0", "60,2"], ["292.200.000", "20.000"], "Scatter")
    offsets = figure.axes[0].collections[0].get_offsets()
    assert list(offsets[:, 1]) == [292200000.0, 20000.0]
    figure.clear()


def test_format_axis_number_shows_thousands():
    assert format_axis_number(1014107000, 0) == "1.014.107.000"
    assert format_axis_number(60.4, 0) == "60.4"


def test_time_values_are_kept_as_time_on_chart_axis():
    assert parse_time_value("09:15:00") == 33300
    assert format_time_axis(33300, 0) == "09:15:00"


def test_is_time_column_detects_only_time_data():
    rows = [["09:15:00", 60], ["09:16:00", 61]]
    assert is_time_column(rows, 0)
    assert not is_time_column(rows, 1)


def test_aggregate_by_time_keeps_last_value_in_each_bucket():
    x_values = ["09:15:01", "09:15:20", "09:15:31"]
    y_values = ["60,0", "60,2", "60,4"]
    grouped_x, grouped_y = aggregate_by_time(x_values, y_values, 30)
    assert grouped_x == ["09:15:00", "09:15:30"]
    assert grouped_y == ["60,2", "60,4"]


def test_filter_by_time_keeps_rows_inside_range():
    x_values = ["09:00:00", "09:15:00", "09:30:00"]
    y_values = [10, 20, 30]
    filtered_x, filtered_y = filter_by_time(x_values, y_values, "09:10", "09:20", True)
    assert filtered_x == ["09:15:00"]
    assert filtered_y == [20]


def test_load_excel_reads_headers_and_rows(tmp_path):
    path = tmp_path / "sample.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Time", "Price"])
    sheet.append(["09:00", "60 +0,4"])
    workbook.save(path)

    headers, rows = load_excel(path)
    assert headers == ["Time", "Price"]
    assert rows == [["09:00", "60 +0,4"]]


def test_create_chart_returns_figure():
    figure = create_chart([1, 2], [3, 4], "Scatter")
    assert len(figure.axes) == 1
    figure.clear()


def test_line_chart_has_no_markers_and_uses_thin_line():
    figure = create_chart([1, 2, 3], [3, 4, 2], "Line")
    line = figure.axes[0].lines[0]
    assert line.get_marker() in (None, "None", "")
    assert line.get_linewidth() == 0.8
    figure.clear()


def test_scatter_shows_each_numeric_x_value():
    figure = create_chart(["59,8", "59,9", "60,0"], [1, 2, 3], "Scatter")
    assert list(figure.axes[0].get_xticks()) == [59.8, 59.9, 60.0]
    figure.clear()


def test_chart_contains_hidden_hover_tooltip():
    figure = create_chart([1, 2], [3, 4], "Line")
    assert len(figure.axes[0].texts) == 1
    assert not figure.axes[0].texts[0].get_visible()
    figure.clear()