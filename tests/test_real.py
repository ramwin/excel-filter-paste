#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Xiang Wang <ramwin@qq.com>


import logging
import unittest
from pathlib import Path

from openpyxl import load_workbook

from excel_filter_paste.paste import paste_convert


ROOT = Path.home().joinpath("test/ava")
logger = logging.getLogger("excel_filter_paste")
handler = logging.StreamHandler()
logger.addHandler(handler)
logger.setLevel(logging.INFO)
handler.setFormatter(logging.Formatter("%(asctime)s %(pathname)s %(message)s"))


@unittest.skipIf(not ROOT.exists(), "文件夹不存在")
class Test(unittest.TestCase):

    def setUp(self):
        pass

    def test(self):
        paste_convert(
            ROOT.joinpath("input2.xlsx"),
            ROOT.joinpath("output.xlsx"),
            ROOT,
        )
