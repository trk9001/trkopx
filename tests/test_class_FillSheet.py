import pytest

from fillsheet.fs import FillSheet


def test_constructor_invalid_file_type():
    with pytest.raises(TypeError):
        fs = FillSheet('SomeFile')


def test_constructor_nonexistent_file():
    with pytest.raises(FileNotFoundError):
        fs = FillSheet('SomeFile.xlsx')
