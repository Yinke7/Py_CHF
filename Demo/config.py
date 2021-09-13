import os
import re

file_conf = 'config.txt'

def ReadConfig():
    dict_conf = {'row': 0, 'col': 0}
    if os.path.exists(file_conf):
        f = open(file_conf)
        list_line = f.readlines()
        line = list_line[0]
        res = re.search(r'row:(.\d*),col:(.\d*)', line)
        dict_conf['row'] = int(res.group(1))
        dict_conf['col'] = int(res.group(2))
        print(dict_conf)
        f.close()
    else:
        print("no config file found")
        return
    return dict_conf

if __name__ == '__main__':
    ReadConfig()
