#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
import sys
import os
import zipfile
import re
import argparse
import glob


def search_dotm_for_string(dir, search_str):
    output = 'match found in file {} \n ...{}...'
    file_names = glob.glob(dir+'/*.dotm')
    matches = 0

    for file_name in file_names:
        doc = zipfile.ZipFile(file_name)
        content = doc.read('word/document.xml')

        if content.find(search_str) != -1:
            i = content.find(search_str)
            matches += 1
            print(output.format('.'+file_name, content[i-40:i+41]))

    print('Total dotm files searched: {}').format(len(file_names))
    print('Total dotm files matched: {}').format(matches)


def main():
    parser = argparse.ArgumentParser(
        description="choose a string to search for and a directory of .dotm files")
    parser.add_argument("--dir", type=str, action='store', default='.',
                        help="directory to scan .dotm files")
    parser.add_argument("search_str", type=str, action='store',
                        help="string to search for in each file")
    args = parser.parse_args()
    search_dotm_for_string(args.dir, args.search_str)


if __name__ == '__main__':
    main()
