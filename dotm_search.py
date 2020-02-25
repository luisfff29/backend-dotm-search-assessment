#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "luisfff29"

import argparse
import zipfile
import sys
import os


# with zipfile.ZipFile('dotm_files/@999.dotm') as zf:
#     print(zf.read('word/document.xml'))


def create_parser():
    parser = argparse.ArgumentParser(
        description="Print full path name of each file and a partial line of the dotm text that was found to contain the search text.")
    parser.add_argument('--dir', action="store",
                        help="OPTIONAL directory of .dotm files to scan")
    parser.add_argument('text', help="Text to search for")
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()

    if not args:
        parser.print_usage()
        sys.exit(1)

    my_dir = args.dir
    my_text = args.text
    # raise NotImplementedError("Your awesome code begins here!")


if __name__ == '__main__':
    main()
