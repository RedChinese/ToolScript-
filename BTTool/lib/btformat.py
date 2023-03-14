# -*- coding: utf-8 -*-
# @Time : 2022/10/10 11:18
# @Author : zhenyu
# @Software: PyCharm
import configparser
from pathlib import Path
import os


class btformat:
    def __init__(self):
        self.homedir = Path.home()
        self.configdir = '{0}/btconfig'.format(self.homedir)

        self.configfilepath = "{0}/btconfig.ini".format(self.configdir)
        if (not os.path.exists(self.configdir)):
            os.system("mkdir " + self.configdir)
            pass
        if (not os.path.exists(self.configfilepath)):
            self.writedefaultformatconfigfile()

    # 读配置文件
    def readformatconfigfile(self):
        print(self.configfilepath)
        cofigidic = {}
        config = configparser.ConfigParser()
        config.sections()
        config.read(self.configfilepath)
        m = config.sections()

        for key in m:
            print(key)
            cofigidic[key] = {}
            for k in config[key]:
                if (k == "bold"):
                    cofigidic[key][k] = (config[key][k] == str(True))
                elif (k == "border" or k == "font_size"):
                    cofigidic[key][k] = int(config[key][k])
                elif (key == "TableHeader"):
                    cofigidic[key][k] = float(config[key][k])
                elif (key == 'DFU' or key == 'FCT' or key == 'PFCT'):
                    cofigidic[key][k] = int(config[key][k])
                else:
                    cofigidic[key][k] = config[key][k]
        pass
        return cofigidic

    # 写默认配置文件
    def writedefaultformatconfigfile(self):
        DEFAULT_REPORT_DIR = '{0}/BTReport/'.format(self.homedir)
        with open(self.configfilepath, 'w') as configfile:
            config = configparser.ConfigParser()
            config['default'] = {}
            default_config = config["default"]
            default_config['report_dic'] = DEFAULT_REPORT_DIR
            default_config['productimagepath'] = ""
            config['TableHeader'] = {}
            table_handle = config['TableHeader']
            table_handle['image_column_witdh'] = str(42.3)
            table_handle['chanel_column_witdh'] = str(13)
            table_handle['fxitureid_column_witdh'] = str(24)
            table_handle['resuilt_witdh'] = str(13)
            table_handle['header_row_hight'] = str(43)
            table_handle['image_row_hight'] = str(200)
            config['header_format'] = {}
            hander_format = config['header_format']
            hander_format['bold'] = str(True)
            hander_format['border'] = str(2)
            hander_format['align'] = 'center'
            hander_format['valign'] = 'vcenter'
            hander_format['font_size'] = str(18)
            hander_format['font_name'] = "Helvetica"
            hander_format['bg_color'] = '#84A6CE'
            config['merge_format'] = {}
            merge_format = config['merge_format']
            merge_format['bold'] = str(False)
            merge_format['border'] = str(2)
            merge_format['font_size'] = str(18)
            merge_format['align'] = 'center'
            merge_format['valign'] = 'vcenter'
            config['row_format'] = {}
            rowformat = config['row_format']
            rowformat['bold'] = str(True)
            rowformat['border'] = str(2)
            rowformat['align'] = 'center'
            rowformat['valign'] = 'vcenter'
            rowformat['font_size'] = str(18)
            rowformat['font_name'] = "SongTi"
            config['resuilt_format'] = {}
            resuilt_format = config['resuilt_format']
            resuilt_format['bold'] = str(True)
            resuilt_format['border'] = str(2)
            resuilt_format['align'] = 'center'
            resuilt_format['valign'] = 'vcenter'
            resuilt_format['font_size'] = str(18)
            resuilt_format['font_name'] = "SongTi"
            resuilt_format['color'] = "green"
            config['image_format'] = {}
            dfu_config = config['image_format']
            dfu_config['bold'] = str(False)
            dfu_config['column_number'] = str(6)
            config['DFU'] = {}
            dfu_config = config['DFU']
            dfu_config['row'] = str(12)
            dfu_config['column'] = str(6)
            config['FCT'] = {}
            dfu_config = config['FCT']
            dfu_config['row'] = str(8)
            dfu_config['column'] = str(12)
            config['PFCT'] = {}
            dfu_config = config['PFCT']
            dfu_config['row'] = str(4)
            dfu_config['column'] = str(12)
            config.write(configfile)
        pass
