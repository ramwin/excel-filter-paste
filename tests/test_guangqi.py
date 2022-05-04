#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Xiang Wang <ramwin@qq.com>


import unittest
from pathlib import Path

from openpyxl import load_workbook

from excel_filter_paste.paste import Guangqifengtian

ROOT = Path.home().joinpath("test/ava")


@unittest.skipIf(not ROOT.exists(), "文件夹不存在")
class Test(unittest.TestCase):

    def setUp(self):
        wb = load_workbook(ROOT.joinpath("input2.xlsx"))
        self.ws = wb.active

    def test1(self):
        converter = Guangqifengtian(sheet=self.ws, directory=ROOT)
        position = converter.get_position(
            "R6788"
        )
        result = converter.get_data()
