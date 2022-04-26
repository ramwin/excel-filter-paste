#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Xiang Wang <ramwin@qq.com>


import datetime
import logging
import unittest

import pandas

from excel_filter_paste.paste import parse_data, paste_convert


logging.basicConfig(level=logging.INFO)


class Test(unittest.TestCase):

    def setUp(self):
        self.input = "tests/fengtong/丰通日报.xlsx"

    def test_fengtong(self):
        self.assertEqual(
            parse_data(self.input),
            {
                pandas.Timestamp("2022-04-01 00:00:00"): {
                    "D": 1,
                    "E": 2,
                    "F": 3,
                    "G": 4,
                },
                pandas.Timestamp("2022-04-02 00:00:00"): {
                    "D": 5,
                    "E": 6,
                    "F": 7,
                    "G": 8,
                }
            }
        )
        self.assertEqual(
            parse_data(
                self.input,
                end=str(datetime.date(2022, 4, 2)),
            ),
            {
                pandas.Timestamp("2022-04-01 00:00:00"): {
                    "D": 1,
                    "E": 2,
                    "F": 3,
                    "G": 4,
                },
            }
        )

    def test_paste(self):
        paste_convert(
            "tests/汇总表.xlsx",
            "tests/结果.xlsx",
            "tests/fengtong"
        )
