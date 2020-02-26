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
    # my_dir = "./dotm_files"
    # my_text = "$"
    count_files_searched = 0
    count_files_matched = 0

    if my_dir:
        print("Searching directory {} for text '{}' ...".format(my_dir, my_text))
        [(dirpath, dirnames, filenames)] = list(os.walk(my_dir))
        for dotm in filenames:
            if dotm.endswith('.dotm'):
                count_files_searched += 1
                file_path = os.path.join(my_dir, dotm)
                with zipfile.ZipFile(file_path) as zf:
                    block40 = zf.read('word/document.xml')
                    i = block40.find(my_text)
                if i > 0:
                    count_files_matched += 1
                    print('Match found in file {}'.format(file_path))
                    print('   ...{}{}...'.format(
                        block40[i-40:i], block40[i:i+len(my_text)+39]))
        print('Total dotm files searched: {}'.format(count_files_searched))
        print('Total dotm files matched: {}'.format(count_files_matched))
    # raise NotImplementedError("Your awesome code begins here!")


if __name__ == '__main__':
    main()
