import pytest
from fillsheet.fs import FillSheet

# Assuming here that pytest is run from the root dir
file_name = 'tests/assets/Descr.xlsx'


def test_constructor_invalid_file_type_exception():
    with pytest.raises(TypeError):
        fs = FillSheet('SomeFile')


def test_constructor_nonexistent_file_exception():
    with pytest.raises(FileNotFoundError):
        fs = FillSheet('SomeFile.xlsx')


def test_constructor_file_ok():
    fs = FillSheet(file_name)
    assert fs.file == file_name


def test_constructor_automatic_fetching_of_rows():
    fs = FillSheet(file_name)
    assert fs.rows.start == 2 and fs.rows.end == 15


def test_constructor_manual_rows_input_ok_1():
    fs = FillSheet(file_name, '3:10')
    assert fs.rows.start == 3 and fs.rows.end == 10


def test_constructor_manual_rows_input_ok_2():
    fs = FillSheet(file_name, ':10')
    assert fs.rows.start == 2 and fs.rows.end == 10


def test_constructor_manual_rows_input_ok_3():
    fs = FillSheet(file_name, '3:')
    assert fs.rows.start == 3 and fs.rows.end == 15


def test_constructor_manual_rows_input_ok_4():
    fs = FillSheet(file_name, ':')
    assert fs.rows.start == 2 and fs.rows.end == 15


def test_constructor_manual_rows_input_start_gt_end_exception():
    with pytest.raises(ValueError):
        fs = FillSheet(file_name, '10:3')


def test_constructor_manual_rows_input_bad_format_exception():
    with pytest.raises(TypeError):
        fs = FillSheet(file_name, '3,10')


def test_constructor_seed_ok():
    fs = FillSheet(file_name)
    assert fs.seed == 4
