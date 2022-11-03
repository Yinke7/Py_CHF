# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import config
import table


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.

# VERSION = r'Ver: 1.1'

# VERSION = r'CHF666 Ver: 1.2'   # 2022.04.20.15.11
# 修复bug：此bug曾导致表格单元个中无内容时，会终止程序的情况
# 现将此情况表格赋值为空字符
# 新增提示消息框

VERSION = r'CHF666 Ver: 1.3'   # 2022.11.03.16.21
# 更新，将config.txt的格式改为json格式，并使用json包解析文件
# 新增，可设置[连续个数]
# 修复，文件最后一行无法查询得问题

if __name__ == '__main__':
    # print_hi('PyCharm')
    print('%s' % VERSION)
    para = config.ReadConfig()
    if para['evaluation'] == 1:
        res = input("评价陈海峰：")
        if res != "帅得一匹":
            input("验证不通过！按回车键结束...")
            exit()
    print('read table...')
    table.read_excel()
    print('write new table...')
    table.write_excel()
    print('check new table...')
    table.check()
    print('end')



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
