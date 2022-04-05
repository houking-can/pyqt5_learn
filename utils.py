from win32com.client import Dispatch
import xlrd
import re
import time
import traceback
from xlutils.copy import copy
import shutil


name2col_id = {
    "同比": "H",
    "环比": "I",
    "累积": "O"
}

def fun_switch(self):
    if self.fun_name == "三费-智慧城市":
        try:
            shutil.copy(self.file, self.savePath)
            old_workbook = xlrd.open_workbook(self.savePath)
            sheet_names = old_workbook.sheet_names()
            new_workbook = copy(old_workbook)
            sale_raw, manage_raw, technology_res = None, None, None
            for i, name in enumerate(sheet_names):
                name = re.sub(r'[\s\u3000]+', '', name)
                if "销售费用底稿" == name:
                    sale_raw = old_workbook.sheet_by_index(i)
                elif "管理费用底稿" == name:
                    manage_raw = old_workbook.sheet_by_index(i)
                elif "研发费用底稿" == name:
                    technology_raw = old_workbook.sheet_by_index(i)

                elif "销售费用" == name:
                    sale_res = new_workbook.get_sheet(i)
                    sale_res_tmp = old_workbook.sheet_by_index(i)
                elif "管理费用" == name:
                    manage_res = new_workbook.get_sheet(i)
                    manage_res_tmp = old_workbook.sheet_by_index(i)
                elif "研发费用" == name:
                    technology_res = new_workbook.get_sheet(i)
                    technology_res_tmp = old_workbook.sheet_by_index(i)

            close_excel_file(self.savePath)
            combine_sheet_names = []
            if sale_raw and sale_res:
                write_split(sale_raw, sale_res_tmp, sale_res)
                combine_sheet_names.append("销售费用")
            if manage_raw and manage_res:
                write_split(manage_raw, manage_res_tmp, manage_res)
                combine_sheet_names.append("管理费用")
            if technology_raw and technology_res:
                write_split(technology_raw, technology_res_tmp, technology_res)
                combine_sheet_names.append("研发费用")

            new_workbook.save(self.savePath)
            do_combine(self.savePath, self.fun_name)

            self.success = True

        except:
            tb = traceback.format_exc()
            self.success = "error"
            self.ui.LogText.appendPlainText(tb)


def close_excel_file(file):
    xlApp = Dispatch('Excel.Application')
    xlApp.DisplayAlerts = False  # 设置不显示警告和消息框
    # xlBook = xlApp.Workbooks.Open(file)
    workbooks_n = xlApp.Workbooks.Count
    if workbooks_n < 0:
        return
    file_tmp = file.replace("\\", "").replace("/", "")
    for i in range(1, workbooks_n + 1):  # 工作簿索引从1开始
        this_path = xlApp.Workbooks(i).Path
        this_name = xlApp.Workbooks(i).Name
        this_file = this_path + "\\" + this_name
        this_file_tmp = this_file.replace("\\", "").replace("/", "")
        if this_file_tmp == file_tmp:
            xlApp.Workbooks(i).Activate()
            xlApp.Workbooks(i).SaveAs(this_file)
            xlApp.Workbooks(i).Close()
            break
    del xlApp


def get_html(content, color):
    template = """<html><head><style type="text/css">p {color: #color#} </style></head><body><p>#content#</p></body></html>"""
    html = template.replace("#content#", content).replace("#color#", color)
    return html


def timer(self):
    time_str = str(time.strftime("[INFO] %Y-%m-%d %H:%M:%S", time.localtime()))
    html = get_html(time_str, "blue")
    self.ui.LogText.appendHtml(html)


def write_split(src, des_tmp, des):
    def do_write(key, sort_key):
        output = ""
        unit = 10000.0
        increase = [(each[0], each[sort_key]) for each in project if each[sort_key] > 0]
        decrease = [(each[0], each[sort_key]) for each in project if each[sort_key] < 0]
        increase.sort(key=lambda k: k[1], reverse=True)
        decrease.sort(key=lambda k: k[1])
        this_table = summary[sort_key - 1]
        if this_table[i] > 0:
            total = this_table[i] / unit
            tmp = f"{total:.2f}"
            if tmp!="0.00":
                output += f"{key}增加{tmp}万：\n其中"
                for each in increase:
                    this_num = each[1] / unit
                    tmp = f"{this_num:.2f}"
                    if tmp!="0.00":
                        output += f"{each[0]}增加{tmp}万，"
                output = output.strip("，") + "；"
                for each in decrease:
                    this_num = abs(each[1]) / unit
                    tmp = f"{this_num:.2f}"
                    if tmp!="0.00":
                        output += f"{each[0]}减少{tmp}万，"
        elif this_table[i] < 0:
            total = abs(this_table[i]) / unit
            tmp = f"{total:.2f}"
            if tmp!="0.00":
                output += f"{key}减少{tmp}万：\n其中"
                for each in decrease:
                    this_num = abs(each[1]) / unit
                    tmp = f"{this_num:.2f}"
                    if tmp!="0.00":
                        output += f"{each[0]}减少{tmp}万，"
                output = output.strip("，") + "；"
                for each in increase:
                    this_num = each[1] / unit
                    tmp = f"{this_num:.2f}"
                    if tmp!="0.00":
                        output += f"{each[0]}增加{tmp}万，"
        output = output.strip("，").strip("；") + "。"
        try:
            index = des_titles.index(title)
            des.write(index, ord(name2col_id[key]) - ord("A"), output)
        except:
            if sort_key == 1:
                print(f"Error: 确保有sheet “销售费用” 有 “{title}” 这个一栏")
            elif sort_key == 2:
                print(f"Error: 确保有sheet “管理费用” 有 “{title}” 这个一栏")
            elif sort_key == 3:
                print(f"Error: 确保有sheet “研发费用” 有 “{title}” 这个一栏")

    same_cmp = src.col_values(ord("G") - ord("A"))
    ring_cmp = src.col_values(ord("H") - ord("A"))
    accumulate_cmp = src.col_values(ord("K") - ord("A"))
    summary = [same_cmp, ring_cmp, accumulate_cmp]
    start_flag = False
    des_titles = des_tmp.col_values(0)
    for i, (title, flag, subject) in enumerate(zip(src.col_values(0), src.col_values(1), src.col_values(2))):
        title = re.sub(r'[\s\u3000]+', '', title)
        flag = re.sub(r'[\s\u3000]+', '', flag)
        subject = re.sub(r'[\s\u3000]+', '', subject)
        if title == "分部报告科目":
            start_flag = True
            project = []
            continue
        if start_flag and flag == "汇总":
            do_write("同比", 1)
            do_write("环比", 2)
            do_write("累积", 3)
            start_flag = False
            project = []
            continue
        if start_flag and subject != "":
            project.append((subject, same_cmp[i], ring_cmp[i], accumulate_cmp[i]))


def do_combine(file, fun_name):
    old_workbook = xlrd.open_workbook(file)
    sheet_names = old_workbook.sheet_names()
    new_workbook = copy(old_workbook)
    process_sheets = []
    process_sheet_names = []
    for i, name in enumerate(sheet_names):
        name = re.sub(r'[\s\u3000]+', '', name)
        if "销售费用" == name:
            process_sheets.append(old_workbook.sheet_by_index(i))
            process_sheet_names.append("销售")
        elif "管理费用" == name:
            process_sheets.append(old_workbook.sheet_by_index(i))
            process_sheet_names.append("管理")
        elif "研发费用" == name:
            process_sheets.append(old_workbook.sheet_by_index(i))
            process_sheet_names.append("研发")
        elif fun_name == name:
            combine_res = new_workbook.get_sheet(i)

    for key_id, key in enumerate(name2col_id.keys()):
        this_key_cols = [sheet.col_values(ord(name2col_id[key]) - ord("A")) for sheet in process_sheets]
        write_index = ord("P") - ord("A") + key_id
        for row_id in range(len(this_key_cols[0])):
            output = []
            if row_id==0:
                combine_res.write(row_id, write_index, f"{key}分析")
                continue
            for sheet_id, sheet_name in enumerate(process_sheet_names):
                content = this_key_cols[sheet_id][row_id]
                if content.startswith(key) and len(content[2:]) > 0:
                    output.append(f"{sheet_name}{content[2:]}")
            if len(output) > 0:
                output_str = ""
                for i, content in enumerate(output):
                    output_str += f"（{i + 1}）{content}\n"
                output_str = output_str.strip()
                combine_res.write(row_id, write_index, output_str)

    new_workbook.save(file)
