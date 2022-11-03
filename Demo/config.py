import os
import re
import json

file_conf = 'config.txt'

def ReadConfig():
    if os.path.exists(file_conf):
        f = open(file_conf)
        para = json.load(f)
        # print(para)
        f.close()
    else:
        # print("no config file found")
        return None
    return para

if __name__ == '__main__':
    ReadConfig()
