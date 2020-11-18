import openpyxl
import sys

def collect_eid(worksheet) -> list:
    result = set()

    for i in range(start_row, end_row + 1):
        result.add(worksheet.cell(row=i, column=eid_column).value.strip())

    result = list(result)
    result.sort()
    return result


def getEIDData(eid_set, worksheet) -> dict:
    result = {}

    for each_eid in eid_set:
        result[each_eid] = {}

    # 对于每一个EID,组装一个字典包含所有的列数和对应的数据
    for each_eid in eid_set:
        # 先扫描到EID对应的行
        for i in range(start_row, end_row + 1):
            if ws.cell(row=i, column=eid_column).value.strip() == each_eid:
                # 组装一个字典
                for j in range(5, worksheet.max_column + 1):
                    if ws.cell(row=i, column=j).value:
                        result[each_eid][j] = ws.cell(row=i, column=j).value

    return result


def writeResult(EIDData, worksheet, fileName='result.xlsx'):

    # 删除第二行开始的数据
    worksheet.delete_rows(2, worksheet.max_row + 1)

    # 从第二行开始写入
    row_start = 2

    for each_eid in EIDData:
        worksheet.cell(row=row_start, column=2).value = each_eid

        for each_column, each_value in EIDData[each_eid].items():
            worksheet.cell(row=row_start, column=each_column).value = each_value

        row_start = row_start + 1

    worksheet.parent.save(fileName)


if __name__ == "__main__":

    filename = 'result.xlsx'

    command_length = len(sys.argv)

    if command_length!=2 and command_length !=3:
        print("参数错误，第一个参数为要处理的文件，第二个参数为要生成的文件名称，可以忽略第二个参数")
        sys.exit(0)

    wb = openpyxl.open(sys.argv[1])

    ws = wb.active

    print("读取表格中....")
    # 数据开始的行数
    start_row = 2
    print("默认数据从第 {} 行开始处理".format(start_row))

    # 数据结束的行数
    end_row = ws.max_row
    print("表格有效行数为：{} 行".format(end_row))


    # EID所在列数
    eid_column = 2
    print("默认EID所在列数：{}".format(eid_column))

    # 最大列数
    print("表格最大列数为：{}".format(ws.max_column))

    print("上述基础参数错误请联系开发者, 按回车键开始处理数据。。。")
    input()

    # 生成EID列表
    print("正在生成不重复的EID清单")
    eid_list = collect_eid(ws)
    print("不重复的EID是：{}".format(eid_list))

    print("正在生成不重复的EID与对应汇总数据")
    eid_data = getEIDData(eid_list, ws)
    print("EID汇总数据为：{}".format(eid_data))

    print("准备写入数据")
    if command_length==2:
        print("未输入文件名，默认为result.xlsx")

    if command_length==3:
        filename = sys.argv[2]
        print("文件名为 {}".format(filename))

    try:
        writeResult(eid_data,ws, filename)

        print("文件已经成功写入到：{}".format(filename))

    except:
        print("写入文件发生错误，请检查文件名是否有效或目标文件已经打开")



