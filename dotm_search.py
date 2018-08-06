#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!

import os
import sys
import glob
import zipfile

def dotm_search(text, directory):
    if directory: 
        os.chdir(directory)
    files = glob.glob('*.dotm')
    files_found = 0
    print "Searching directory {} for text '{}'".format(directory, text)
    for file in files:
        with zipfile.ZipFile(file) as z:
            with z.open('word/document.xml', 'r') as doc:
                for line in doc:
                    if text in line:
                        print "Match found in file {}/{}".format(directory, file)
                        print "..." + line[line.find(text) - 40:line.find(text) + 40] + "..."
                        files_found += 1
    print "Files searched: {}".format(len(files))
    print "Files matched: {}".format(files_found)

if __name__ == '__main__':
    if len(sys.argv) > 2:
        dotm_search(sys.argv[1], sys.argv[2])
    elif len(sys.argv) > 1:
        dotm_search(sys.argv[1], os.getcwd())
    else: 
        print "Please provide searchable text."