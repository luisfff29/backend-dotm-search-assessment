#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "luisfff29"

import zipfile
import os

with zipfile.ZipFile('dotm_files/@999.dotm') as zf:
    print(zf.read('word/document.xml'))
print(os.getcwd)


def main():
    raise NotImplementedError("Your awesome code begins here!")


# if __name__ == '__main__':
#     main()
