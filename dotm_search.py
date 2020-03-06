#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "mike gabbard"

import zipfile
import argparse
import os
import sys

def create_parser():
    parser = argparse.ArgumentParser(description='search for text in dotm files')
    parser.add_argument('--dir', help='directory to search')
    parser.add_argument('text', help='test to find')
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()
    text = args.text
    path = args.dir
    files = os.listdir(path)
    file_type = 'word/document.xml'
    matched = 0
    searched = 0

    for file in files:
        if file.endswith('dotm'):
            searched += 1 
            full_path = os.path.join(path,file)
        if zipfile.is_zipfile(full_path):
            with zipfile.ZipFile(full_path) as f:
                names = f.namelist()
                if file_type in names:
                    with f.open(file_type) as doc:
                        data = doc.read()
                        location = data.find(text)
                        if location >-1:
                            matched +=1
                            print('Match found in file {}'.format(full_path))
                            print('...'+ data[location-40:location+40]+'...')
    print('total searched: ' + str(searched))
    print('total matched: ' + str(matched))


if __name__ == '__main__':
    main()
