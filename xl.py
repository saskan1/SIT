
import xlrd

def make_json_from_data(column_names, row_data):
    """
    take column names and row info and merge into a single json object.
    :param data:
    :param json:
    :return:
    """
    row_list = []
    for item in row_data:
        json_obj = {}
        for i in range(0, column_names.__len__()):
            json_obj[column_names[i]] = item[i]
            row_list.append(json_obj)



    return row_list

def xls_to_dict(workbook_url):
    """
    Convert the read xls file into JSON.
    :param workbook_url: Fully Qualified URL of the xls file to be read.
    :return: json representation of the workbook.
    """
    workbook_dict = {}
    book = xlrd.open_workbook(workbook_url)
    sheets = book.sheets()
    for sheet in sheets:
        if sheet.name == 'Input':
            continue
        workbook_dict[sheet.name] = {}
        columns = sheet.row_values(0)
        rows = []
        for row_index in range(1, sheet.nrows):
            row = sheet.row_values(row_index)
            rows.append(row)
        sheet_data = make_json_from_data(columns, rows)
        workbook_dict[sheet.name] = sheet_data
    return workbook_dict

d={}
d = xls_to_dict('Input_param_ram.xlsx')

with open('test.csv', 'w') as f:
    for key in d.keys():
        f.write("%s,%s\n"%(key,d[key]))