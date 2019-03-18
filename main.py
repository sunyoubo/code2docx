# coding:utf-8

import os
import argparse

from docx import Document


def main(directory, file_name):
    if not directory:
        return
    if not file_name:
        file_name = 'test'
    document = Document()
    for root, dirs, fs in os.walk(directory):
        for f in fs:
            fpath = os.path.join(root, f)
            fp = open(fpath, "r", encoding='utf-8')
            document.add_paragraph(fp.read())
            fp.close()

    document.save(file_name)


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--directory', default=None, type=str, help='code directory')
    parser.add_argument('-f', '--file', default=None, type=str,
                        help='file name of *.docx')

    arguments = parser.parse_args()
    main(arguments.directory, arguments.file)
