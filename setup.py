#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Xiang Wang <ramwin@qq.com>


from setuptools import setup, find_packages


setup(
    name="excel_filter_paste",
    version="0.0.3",
    install_requires=[
        "openpyxl",
        "click",
    ],
    packages=["excel_filter_paste"],
)
