# coding:utf-8

import os
import argparse

from docx.api import Document


def main(directory, file_name):

    if not directory:
        return
    if not file_name:
        file_name = 'test'

    document = Document()
    total_lines = 0
    for root, dirs, fs in os.walk(directory):
        for f in fs:
            suffix = os.path.splitext(f)[1]
            if suffix not in [".java", ".py", ".js", ".go", ".css", ".html", ".cpp", ".h"]:
                continue

            fpath = os.path.join(root, f)
            with open(fpath, "r") as fp:
                try:
                    lines = fp.readlines()
                except Exception as e:
                    print(fpath)
                    continue

                for line in lines:
                    line = line.strip("\n").strip("\r\n")
                    if not line:
                        continue
                    if "__author__ " in line or "__datetime__" in line:
                        continue

                    if total_lines > 3500:
                        break
                    document.add_paragraph(line)
                    total_lines += 1

    document.save(file_name+".docx")


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--directory', default=None, type=str, help='code directory')
    parser.add_argument('-f', '--file', default=None, type=str,
                        help='file name of *.docx')

    arguments = parser.parse_args()
    main(arguments.directory, arguments.file)
