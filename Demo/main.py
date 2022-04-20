# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import table


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.

# VERSION = r'Ver: 1.1'

VERSION = r'CHF666 Ver: 1.2'   # 2022.04.20.15.11
# 修复bug：此bug曾导致表格单元个中无内容时，会终止程序的情况
# 现将此情况表格赋值为空字符
# 新增提示消息框


if __name__ == '__main__':
    # print_hi('PyCharm')
    print('%s' % VERSION)
    print('read table...')
    table.read_excel()
    print('write new table...')
    table.write_excel()
    print('check new table...')
    table.check()
    print('end')



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
