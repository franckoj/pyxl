"""Pytest configuration and common fixtures."""

import pytest
import pyxl
import openpyxl


@pytest.fixture
def pyxl_wb():
    wb = pyxl.Workbook()
    yield wb
    wb.close()


@pytest.fixture
def openpyxl_wb():
    wb = openpyxl.Workbook()
    yield wb
    wb.close()
