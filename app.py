# coding:utf8

from io import StringIO
from math import isnan

import numpy as np
import pandas as pd
from flask import Flask
from flask import render_template, send_file
from flask import request

app = Flask(__name__)


@app.route('/')
def index():
    return send_file("static/index.html")


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("file")

    error_msg = StringIO()

    report_excel_data = init_report_excel_data()

    app.logger.info("遍历upload文件")
    for file in files:
        file_name = file.filename
        parse_check_file(error_msg, file, file_name, report_excel_data, app.logger)

    report_doc_data = gen_doc_file(report_excel_data, app.logger, error_msg)

    report_excel_data_html = StringIO()
    report_doc_data_1_html = StringIO()
    report_doc_data_2_html = StringIO()
    report_doc_data_3_html = StringIO()

    pd.set_option('display.max_colwidth', -1)

    report_excel_data.to_html(report_excel_data_html, escape=False, index=False, index_names=False, header=True,
                              na_rep="-")

    report_doc_data[0].to_html(report_doc_data_1_html, escape=False, index=False, index_names=False, header=True,
                               na_rep="-")
    report_doc_data[1].to_html(report_doc_data_2_html, escape=False, index=False, index_names=False, header=True,
                               na_rep="-")
    report_doc_data[2].to_html(report_doc_data_3_html, escape=False, index=False, index_names=False, header=True,
                               na_rep="-")

    error_msg = error_msg.getvalue().replace(";", "<br>")

    return render_template("index.html",
                           report_excel_data=report_excel_data_html.getvalue().replace("\\n", "<br>"),
                           report_doc_data1=report_doc_data_1_html.getvalue().replace("\\n", "<br>"),
                           report_doc_data2=report_doc_data_2_html.getvalue().replace("\\n", "<br>"),
                           report_doc_data3=report_doc_data_3_html.getvalue().replace("\\n", "<br>"),
                           error_msg=error_msg)

    # resp = Response({"code": 200})
    # resp.headers['Access-Control-Allow-Origin'] = '*'


def init_report_excel_data():
    src_file_name_cn = "来源文件"

    project_no_cn = "项目编号"
    company_name_cn = "合作单位"
    lost_score_item_no_cn = "扣分条款"
    lost_score_cn = "扣分值"
    lost_score_reason_cn = "扣分说明"
    project_owner_cn = "项目归属省工程科室"

    report_excel_data = pd.DataFrame(columns=[src_file_name_cn,
                                              project_no_cn, company_name_cn, lost_score_item_no_cn,
                                              lost_score_cn, lost_score_reason_cn, project_owner_cn])
    return report_excel_data


def do_strip_dict(data):
    data_tmp = {}
    for k, v in data.items():
        data_tmp[k.strip()] = v

    data.clear()
    data.update(data_tmp)


def do_strip(data):
    if data.columns[0].startswith("附件"):
        data.columns = data.iloc[0].values

    new_columns = []
    for column in data.columns:
        if type(column) == str:
            new_columns.append(column.strip())
        else:
            new_columns.append(column)

    data.columns = new_columns


def valid_check(series, expect_types):
    for item in series:
        try:
            if isnan(item):
                continue
        except Exception:
            pass

        if type(item) in expect_types:
            return True
    else:
        return False


def get_company_full_name(ex_data, sheet_name):
    return sheet_name


def gen_doc_file(df, logger, error_msg, states=("设计", "监理", "施工"), valid_columns=("来源文件", "合作单位", "扣分值", "扣分说明")):
    logger.debug("过滤NaN")
    doc_df = df[list(valid_columns)]
    valid_doc_df = doc_df.dropna()

    new_column_name = "阶段"
    logger.debug("添加 {} 列".format(new_column_name))
    new_column = []
    for file_name in valid_doc_df[valid_columns[0]]:
        new_column_item = None
        for state in states:
            if state in file_name:
                new_column_item = state
                break
        new_column.append(new_column_item)

    valid_doc_df[new_column_name] = new_column

    logger.debug("按照 {}->{} 排序".format(new_column_name, valid_columns[1]))
    valid_doc_df = valid_doc_df.sort_values(by=[new_column_name, valid_columns[1]])

    logger.debug("添加数据 到result")
    num_cn = "排名"
    company_cn = "合作单位"
    cost_score_cn = "总扣分"
    score_cn = "总成绩"
    cost_score_reason_cn = "扣分情况"

    res_dfs = []
    for _ in range(len(states)):
        res_dfs.append(pd.DataFrame(columns=[num_cn, company_cn, cost_score_cn, score_cn, cost_score_reason_cn]))

    last_state = None
    last_company = None
    cost_score_sum = 0
    cost_score_reason_summary = ""

    for _, sub_item in valid_doc_df.iterrows():
        curr_state = sub_item[new_column_name]
        curr_company = sub_item[valid_columns[1]]
        curr_score = sub_item[valid_columns[2]]
        curr_score_reason = sub_item[valid_columns[3]]

        if curr_state != last_state or curr_company != last_company:
            if last_state is None and last_company is None:
                pass

            elif curr_state in states:
                for ind in range(len(states)):
                    if curr_state == states[ind]:
                        res_dfs[ind].loc[res_dfs[ind].shape[0]] = [0,
                                                                   last_company, cost_score_sum,
                                                                   100 - cost_score_sum,
                                                                   cost_score_reason_summary]

                cost_score_sum = 0
                cost_score_reason_summary = ""

            else:
                error_msg_4th_state = "未知的分类"
                logger.error(error_msg_4th_state)
                error_msg.write("{};".format(error_msg_4th_state))
                continue

        last_state = curr_state
        last_company = curr_company
        cost_score_sum += curr_score
        cost_score_reason_summary = "\n".join([cost_score_reason_summary, curr_score_reason])
    else:
        for ind in range(len(states)):
            if last_state == states[ind]:
                res_dfs[ind].loc[res_dfs[ind].shape[0]] = [0,
                                                           last_company, cost_score_sum,
                                                           100 - cost_score_sum,
                                                           cost_score_reason_summary]

    for ind in range(len(res_dfs)):
        res_dfs[ind] = res_dfs[ind].sort_values(by=[score_cn], ascending=False)
        res_dfs[ind][num_cn] = range(len(res_dfs[ind][num_cn]))

    return res_dfs


def parse_check_file(error_msg, file, file_name, report_excel_data, logger):
    logger.info("处理文件:{}".format(file_name))
    ex_data = pd.read_excel(file, sheetname=None)
    do_strip_dict(ex_data)

    # total_sheet_name = "汇总"
    # logger.debug("整理sheet:{}".format(total_sheet_name))
    # total_sheet = ex_data[total_sheet_name]
    # do_strip(total_sheet)
    # logger.debug("筛选扣分单位")
    # lost_score_column_name = "总扣分"
    # lost_score_not_0 = total_sheet[total_sheet[lost_score_column_name] != 0]
    #
    # company_name_column_name = "合作单位"
    # for _, item in lost_score_not_0.iterrows():
    #     company_name = item[company_name_column_name]
    #     lost_score_sum = item[lost_score_column_name]
    #
    #     logger.info("寻找匹配:{}的sheet".format(company_name))
    #     for sheet_name in ex_data.keys():
    #         for sheet_name_word in sheet_name:
    #             if sheet_name_word not in company_name:
    #                 break
    #         else:
    #             dst_sheet_name = sheet_name
    #             logger.info("找到匹配:{}的sheet:{}".format(company_name, dst_sheet_name))
    #             break
    #     else:
    #         error_msg0 = "在 {} 找不到匹配 {} 的sheet".format(file_name, company_name)
    #         logger.error(error_msg0)
    #         error_msg += "; {}".format(error_msg0)
    #         continue
    skip_sheets = ["汇总", "汇总表"]
    for dst_sheet_name in ex_data.keys():
        logger.debug("解析sheet:{}".format(dst_sheet_name))

        if dst_sheet_name in skip_sheets:
            logger.debug("跳过sheet:{}".format(dst_sheet_name))
            continue

        sub_lost_score_column_name = "月度考核扣分"
        sub_lost_score_info_column_name = "月度考核意见"
        sub_project_owner_column_name = "项目归属省工程科室"
        sub_project_no_column_name = "项目编号"
        sub_lost_score_item_no_column_name = "序号"

        sub_sheet = ex_data[dst_sheet_name]
        do_strip(sub_sheet)

        logger.debug("检查数据有效性")
        to_check_columns = {sub_lost_score_item_no_column_name: [int, float],
                            sub_project_no_column_name: [str],
                            sub_lost_score_column_name: [int, float]}

        curr_sheet_error = False
        for to_check_column in to_check_columns.keys():
            if not valid_check(sub_sheet[to_check_column], to_check_columns[to_check_column]):
                error_msg_valid_check = "file:{}, sheet:{}, column:{} 数据类型错误".format(file_name, dst_sheet_name,
                                                                                     to_check_column)
                curr_sheet_error = True
                logger.error(error_msg_valid_check)
                error_msg.write("{};".format(error_msg_valid_check))

        if curr_sheet_error:
            logger.debug("数据异常，跳过当前页")
            continue

        logger.info("{} 特殊处理".format(sub_lost_score_item_no_column_name))
        sub_sheet[sub_lost_score_item_no_column_name] = sub_sheet[
            sub_lost_score_item_no_column_name].replace([np.nan], method="pad")

        logger.debug("{} 过滤".format(sub_lost_score_info_column_name))
        sub_items = sub_sheet[
            sub_sheet[sub_lost_score_info_column_name].apply(
                lambda x: True if type(x) == str and x != sub_lost_score_info_column_name else False)]
        # logger.debug("判断扣分是否计算正确")
        # if sub_items[sub_lost_score_column_name].sum() != lost_score_sum:
        #     error_msg1 = "sheet {} 扣分总数:{}, 汇总 扣分数:{}, 二者不一致".format(dst_sheet_name,
        #                                                              sub_items[sub_lost_score_column_name].sum(),
        #                                                              lost_score_sum)
        #     logger.error(error_msg1)
        #     error_msg += "; {}".format(error_msg1)
        #     continue

        logger.debug("添加扣分项到上报省公司")
        for _, sub_item in sub_items.iterrows():

            if type(sub_item[sub_project_no_column_name]) in (int, float) and isnan(
                    sub_item[sub_project_no_column_name]):
                error_msg_nan = "file:{}, sheet:{}, column:{} 存在空数据".format(file_name, dst_sheet_name,
                                                                            sub_project_no_column_name)
                logger.warn(error_msg_nan)
                error_msg.write("{};".format(error_msg_nan))

            if type(sub_item[sub_lost_score_column_name]) in (int, float) and isnan(
                    sub_item[sub_lost_score_column_name]):
                error_msg_nan = "file:{}, sheet:{}, column:{} 存在空数据".format(file_name, dst_sheet_name,
                                                                            sub_lost_score_column_name)
                error_msg.write("{};".format(error_msg_nan))

            report_excel_data.loc[report_excel_data.shape[0]] = (
                file_name,
                sub_item[sub_project_no_column_name], get_company_full_name(ex_data, dst_sheet_name),
                int(sub_item[sub_lost_score_item_no_column_name]), sub_item[sub_lost_score_column_name],
                sub_item[sub_lost_score_info_column_name], sub_item[sub_project_owner_column_name])

            # print("filename:{}, dst_sheet_name:{}, sub_project_no_column_name:{}, "
            #       "dst_sheet_name:{}, sub_lost_score_item_no_column_name:{}, sub_lost_score_column_name:{}"
            #       "sub_lost_score_info_column_name:{}, sub_project_owner_column_name:{}".format(
            #     file_name, dst_sheet_name, sub_item[sub_project_no_column_name],
            #     dst_sheet_name, sub_item[sub_lost_score_item_no_column_name], sub_item[sub_lost_score_column_name],
            #     sub_item[sub_lost_score_info_column_name], sub_item[sub_project_owner_column_name]))


if __name__ == '__main__':
    app.run()
