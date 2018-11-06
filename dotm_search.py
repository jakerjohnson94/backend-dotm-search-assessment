#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
import sys
import os
import zipfile
import re
from xml.etree.ElementTree import XML

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_dir_files(dir):
    filenames = os.listdir(dir)
    return [os.path.join(dir, filename) for filename in filenames if filename.endswith('.dotm')]


def unzip_dotm_files():
    output = 'match found in file {} \n ...{}...'
    file_names = get_dir_files('dotm_files')
    for filename in file_names:
        doc = zipfile.ZipFile(filename)
        xml_content = doc.read('word/document.xml')
        if xml_content.find('$') != -1:
            i = xml_content.find('$')
            print(output.format(filename.split(
                '/')[1], xml_content[i-40:i+41]))



def main():
    unzip_dotm_files()


if __name__ == '__main__':
    main()
