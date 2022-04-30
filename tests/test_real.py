#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Xiang Wang <ramwin@qq.com>


import logging
import unittest
from pathlib import Path

from openpyxl import load_workbook

from excel_filter_paste.parse import parse_data, get_fengtong_file
from excel_filter_paste.paste import get_fengtong
from excel_filter_paste.paste import paste_convert


ROOT = Path.home().joinpath("test/ava")
logger = logging.getLogger("excel_filter_paste")
logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.INFO)


@unittest.skipIf(not ROOT.exists(), "文件夹不存在")
class Test(unittest.TestCase):

    def setUp(self):
        pass

    def test_parse(self):
        wb = load_workbook(ROOT.joinpath("input2.xlsx"))
        self.assertEqual(
            get_fengtong(wb.active, "42652-06D50"),
            {
                "row": 2260,
                "column": "I",
            }
        )

    def test(self):
        paste_convert(
            ROOT.joinpath("input2.xlsx"),
            ROOT.joinpath("output.xlsx"),
            ROOT,
        )
