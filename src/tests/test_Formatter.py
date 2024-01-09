from unittest import TestCase

from openpyxl import Workbook

from src.pyDicts.Formatter import Formatter


class TestFormatter(TestCase):
    def test_format(self):
        wb = Workbook()
        sht = wb.create_sheet('Demo')
        sht["A1"] = "Hello World"
        Formatter(sht).format("A1")
        wb.save('test2.xlsx')

