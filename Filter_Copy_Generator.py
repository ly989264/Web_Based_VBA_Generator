# VBA Filter Copy Generator
# By Yu Liu

# This script can be used to generate multi-job filter-copy-paste VBA code

# This script will be used in the backend of a web server using the JSON posted by the frontend

# Note:
# Must ensure that the information and format are all correct in JSON file
# Must use at least one filter for each job
# Must ensure that the A column of the target worksheet has the longest column (filled and longest)

import json


def operate(json_dict):
    job_list = json_dict["job"]
    job_num = len(job_list)
    result_str = ""

    job_name_list = []

    job_cnt = 0

    for each_job in json_dict["job"]:
        # default value
        current_workbook = "ActiveWorkbook"
        target_workbook = "ActiveWorkbook"

        job_cnt += 1
        job_name_list.append(f"operation_{job_cnt}")

        # VBA Dim
        result = f"""Sub operation_{job_cnt}()
    Dim current_workbook As Workbook
    Dim target_workbook As Workbook
    Dim current_worksheet As Worksheet
    Dim target_worksheet As Worksheet
    Dim row_num As Integer
    Dim target_row_num As Integer
    Dim arr_length As Integer
    Dim current_arr_index As Integer
    Dim i As Integer
    Dim current_row_num As Integer
"""
        fix2col = {}
        if "fix2col" in each_job["copy"]:
            for each in each_job["copy"]["fix2col"]:
                fix2col[each] = "\"" + each_job["copy"]["fix2col"][each] + "\""
        col2col = {}
        col2col_reverse = {}
        if "col2col" in each_job["copy"]:
            for each_key in each_job["copy"]["col2col"]:
                col2col[each_key] = each_job["copy"]["col2col"][each_key]
                if each_job["copy"]["col2col"][each_key] in col2col_reverse.keys():
                    col2col_reverse[each_job["copy"]["col2col"][each_key]].append(each_key)
                else:
                    col2col_reverse[each_job["copy"]["col2col"][each_key]] = []
                    col2col_reverse[each_job["copy"]["col2col"][each_key]].append(each_key)
        key_list = col2col.keys()
        key_list = sorted(key_list)
        for i in key_list:
            result += f"    Dim {i.lower()}_arr() As String\n"
        result += "\n"
        current_workbook_writable = current_workbook if each_job["worksheet_configuration"]["current_workbook"] is None else each_job["worksheet_configuration"]["current_workbook"]
        result += f"    Set current_workbook = {current_workbook_writable}\n"
        target_workbook_writable = target_workbook if each_job["worksheet_configuration"]["target_workbook"] is None else each_job["worksheet_configuration"]["target_workbook"]
        result += f"    Set target_workbook = {target_workbook_writable}\n"
        result += f"    Set current_worksheet = current_workbook.Worksheets(\"{each_job['worksheet_configuration']['current_worksheet']}\")\n"
        result += f"    Set target_worksheet = target_workbook.Worksheets(\"{each_job['worksheet_configuration']['target_worksheet']}\")\n"
        result += f"    row_num = current_worksheet.Range(\"{each_job['filter']['filters'][0]['column_str']}\" & Rows.Count).End(xlUp).Row\n"
        for i in key_list:
            result += f"    Redim {i.lower()}_arr(row_num) As String\n"  # even if not use date or some data types else, still fine
        result += """    arr_length = 0
    current_arr_index = 0
    target_row_num = target_worksheet.Range("A" & Rows.Count).End(xlUp).Row\n
"""
        result += f"    For i = 2 To row_num\n"
        for each_filter in each_job['filter']['filters']:
            cnt = 0
            result += f"        If"
            for each_criteria in each_filter['criteria']:
                if cnt != 0:
                    result += " And" if each_filter['operator'] == 'AND' else " Or"
                cnt +=1
                if each_criteria.startswith("<>"):
                    result += f" current_worksheet.Cells(i, {ord(each_filter['column_str'].lower()) - ord('a') + 1}).Value <> \"{each_criteria[2:]}\""
                else:
                    result += f" current_worksheet.Cells(i, {ord(each_filter['column_str'].lower()) - ord('a') + 1}).Value = \"{each_criteria}\""
            result += " Then\n"
        cnt = 0
        for i in key_list:
            result += f"            {i.lower()}_arr(current_arr_index) = current_worksheet.Cells(i, {ord(col2col[i].lower()) - ord('a') + 1}).Value\n"
        result += f"            current_arr_index = current_arr_index + 1\n"
        for i in range(len(each_job['filter']['filters'])):
            result += f"        End If\n"
        result += f"    Next i\n\n"

        # delete row
        if each_job['options']['append'].lower() == 'false':
            result += f'    target_worksheet.Rows("2:" & target_row_num).EntireRow.Delete\n'
            result += f'    current_row_num = 2\n'
        else:
            result += f'    current_row_num = target_row_num + 1\n'
        if each_job["options"]["disable_screen_updating"].lower() == "true":
            result += f'    Application.ScreenUpdating = False\n'
        result += f"\n    For i = 0 To current_arr_index - 1\n"
        total_key_arr = sorted(list(set(fix2col.keys()).union(set(col2col.keys()))))
        # print(total_key_arr)
        for each_key in total_key_arr:
            result += f"        target_worksheet.Cells(current_row_num, {ord(each_key.lower()) - ord('a') + 1}).Value = {(each_key.lower() + '_arr(i)') if each_key in col2col.keys() else fix2col[each_key]}\n"
        result += f"        current_row_num = current_row_num + 1\n"
        result += f"    Next i\n"
        if each_job["options"]["disable_screen_updating"].lower() == "true":
            result += f'    Application.ScreenUpdating = True\n'
        result += "End Sub\n"
        # print(result, file=f)
        result_str += result

    # wrapper
    result = "Sub wrapper()\n"
    for each in job_name_list:
        result += f"    Call {each}\n"
    result += "End Sub\n"
    # print(result, file=f)
    result_str += result
    return result_str
