import xlrd
import xlsxwriter
import os


def get_data_from_table(file_name):
    file = xlrd.open_workbook(file_name)
    select_sheet = file.sheet_by_name("暂无")

    row_list = []
    # 获取总共的行数
    rows_num = select_sheet.nrows
    print(rows_num)

    for row in range(rows_num - 2):
        if select_sheet.cell(row + 2, 7).value != '/':
            row_list.append(select_sheet.cell(row + 2, 4).value)
    return row_list


def get_file_list(video_dir):
    video_list = []
    for root, dirs, files in os.walk(video_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.mp3' or os.path.splitext(file)[1] == '.mp4':
                video_list.append(os.path.splitext(file)[0])
    return video_list


def find_file(q_list, video_list):
    not_find_list = []
    for q_value in q_list:
        if not video_list.count(q_value):
            not_find_list.append(q_value)
    return not_find_list


def write_excel(file_list):
    wb = xlsxwriter.Workbook('E:\\work\\video\\not_find_q.xlsx')
    sheet = wb.add_worksheet('sheet1')
    headings = ['Number', 'FileName']
    sheet.write_row('A1', headings)
    i = 0
    for file in file_list:
        data = [i + 1, file]
        sheet.write_row('A' + str(i + 2), data)
        i = i + 1
        data.clear()
    wb.close()
    return


if __name__ == '__main__':
    list_file = find_file(get_data_from_table('E:\\work\\doc\\灯显音效定义7月27 变更.xlsx')
                          , get_file_list('E:\\work\\video'))
    write_excel(list_file)

