from OWGR import OWGR
import re
from openpyxl.styles import PatternFill
import pytest


def test_get_excel_sheet_data():
    # test that I can read the desired data from an excel file and return
    # a list
    results = OWGR.get_excel_sheet_data("docs/test.xlsx")
    # assert that they equal my expectations
    assert results == [(1, "Johnson, Dustin", "FF80FF80", "FFFF8080"),
                           (2, "Rahm, Jon", "FF80FF80", "FFFF8080")]


def test_process_lists():
    # test that I can compare two lists and return a third list
    from_excel = [
                  (1, "Johnson, Dustin", "FF80FF80", "FFFF8080"),
                  (2, "Rahm, Jon", "FF80FF80", "FFFF8080")
                  ]
    from_web = [
                (1, "Johnson, Dustin"),
                (2, "Kingsley, Hank"),
                (3, "Rahm, Jon"),
                (4, "Feeble, Hinkey")
                ]
    from_function = [(1, 'Johnson, Dustin',
                         PatternFill(start_color="FF80FF80",
                                     end_color="FF80FF80",
                                     fill_type="solid"),
                         PatternFill(start_color="FFFF8080",
                                     end_color="FFFF8080",
                                     fill_type="solid")),

                     (2, 'Kingsley, Hank',
                         PatternFill(start_color="FFFF8080",
                                     end_color="FFFF8080",
                                     fill_type="solid"),
                         PatternFill(start_color="FFFF8080",
                                     end_color="FFFF8080",
                                     fill_type="solid")),
                     (3, 'Rahm, Jon',
                         PatternFill(start_color="FF80FF80",
                                     end_color="FF80FF80",
                                     fill_type="solid"),
                         PatternFill(start_color="FFFF8080",
                                     end_color="FFFF8080",
                                     fill_type="solid")),
                     (4,
                      'Feeble, Hinkey',
                      PatternFill(start_color="FFFF8080",
                                  end_color="FFFF8080",
                                  fill_type="solid"),
                      PatternFill(start_color="FFFF8080",
                                  end_color="FFFF8080",
                                  fill_type="solid"))]
    results = OWGR.process_lists(from_excel, from_web)

    assert results == from_function
