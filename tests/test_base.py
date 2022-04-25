#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Xiang Wang <ramwin@qq.com>


import logging
from pathlib import Path
import unittest


from excel_filter_paste.base import convert


logger = logging.getLogger("excel_filter_paste")
logger.setLevel(logging.DEBUG)
logger.addHandler(logging.StreamHandler())


class Test(unittest.TestCase):

    def setUp(self):
        self.input_path = Path("tests/input.xlsx")
        self.output_path = Path("tests/output.xlsx")
        if self.output_path.exists():
            self.output_path.unlink()

    def test(self):
        convert(self.input_path, self.output_path, "C", "B")
