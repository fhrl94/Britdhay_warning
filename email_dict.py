import os
import xlrd


#  人员范围已处理,详见 loading 中 _create_data() 限定了部门
def to_send_email(file_name):
    # file_name = '主管及以上名单.xlsx'
    assert os.path.exists(file_name), "{file}文件不存在".format(file=file_name)
    send_email = {}
    workbook = xlrd.open_workbook(file_name)
    for sheet, one in enumerate(workbook.sheet_names()):
        for i in range(1, workbook.sheet_by_name(one).nrows):
            if workbook.sheet_by_name(one).cell_value(i, 2) != "":
                # print(workbook.sheet_by_name(one).cell_value(i, 2))
                # print(workbook.sheet_by_name(one).cell_value(i, 6))
                send_email[workbook.sheet_by_name(one).cell_value(i, 2)] = workbook.sheet_by_name(one).cell_value(i, 6)
    # print(send_email)
    # pprint.pprint(send_email)
    return send_email

# if __name__ == '__main__':
#     to_send_email()

