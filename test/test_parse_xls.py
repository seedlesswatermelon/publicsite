# coding:utf8

import logging
import os
from io import StringIO

import pandas as pd

from app import init_report_excel_data, parse_check_file, gen_doc_file

if __name__ == "__main__":
    file_path = r"C:\Users\chazzhuang\Documents\WeChat Files\zhi_chang\Files\汇总数据\01 收集材料"
    error_msg = StringIO()

    report_excel_data = init_report_excel_data()

    files = [filename for filename in os.listdir(file_path) if not filename.startswith("~$")]
    for filename in files:
        parse_check_file(error_msg, file_path + os.path.sep + filename, filename, report_excel_data, logging)
    pd.set_option('display.max_colwidth', -1)

    excel_html = "report_excel_data.html"
    report_excel_data.to_html(excel_html, index=False, index_names=False, header=True, na_rep="-")

    report_doc_data = gen_doc_file(report_excel_data, logging, error_msg)[1]
    doc_html = "report_doc_data.html"
    report_doc_data.to_html(doc_html, index=False, index_names=False, header=True, na_rep="-")

    print(error_msg.getvalue())
    # print(ss.getvalue().replace("\\n", "<br>"))

    # print(report_excel_data)
