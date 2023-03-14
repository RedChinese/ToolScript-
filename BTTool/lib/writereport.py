# -*- coding: utf-8 -*-
# @Time : 2022/10/10 13:16
# @Author : zhenyu
# @Software: PyCharm
import os
from pathlib import Path
from lib.btformat import btformat
import xlsxwriter


class writereport:
    REPORT_TYPE_LIST = ["DFU", "BI", "FCT", "PFCT"]
    HOME_PATH = Path.home()
    REPORT_DIR = '{0}/BTReport/'.format(HOME_PATH)  # 报告存放路径
    CONFIG_DIR = "{0}/BTConfig/".format(HOME_PATH)
    CONFIG_Path = "{0}config.ini".format(CONFIG_DIR)

    def __init__(self, fixturename, reporttype):
        print("")
        print(fixturename, reporttype)
        # self.reporttype="DFU";
        # self.column_number=0;
        # self.row_number=0;
        self.fixturename = fixturename;
        self.productimagespath = ""
        self.btformat = btformat()
        self.reporttype = reporttype
        self.config = self.btformat.readformatconfigfile()
        if reporttype in writereport.REPORT_TYPE_LIST:
            self.column_number = self.config[reporttype]["column"];
            self.row_number = self.config[reporttype]["row"];
            print("column", self.column_number)
            print("row", self.row_number)
            print("tyep:", self.reporttype)
        if (not os.path.exists(writereport.REPORT_DIR)):
            os.system("mkdir " + writereport.REPORT_DIR)
        self.file_path = writereport.REPORT_DIR + "{0}.xlsx".format(fixturename)
        print("生成蓝膜报告：", self.file_path)
        self.workbook = xlsxwriter.Workbook(self.file_path)
        self.worksheet = self.workbook.add_worksheet()

        pass

    # 设置报告类型

    def add_tableHandle(self):
        print("添加表头")
        format = self.workbook.add_format(self.config["header_format"])
        tableHead = ['Fixture Sn', 'CH No.']
        for i in range(1, self.column_number + 1):
            tableHead.append('plcture {0}#'.format(i))
        tableHead.append('result')
        self.worksheet.write_row(1, 1, tableHead, format)
        pass

    def setMerge_format(self):
        print("合并单元格")
        format = self.workbook.add_format(self.config["merge_format"])
        self.worksheet.merge_range(2, 1, self.row_number + 1, 1, self.fixturename, format)

    def setColumn_format(self):
        print("设置列宽")
        self.worksheet.set_column(0, 0, 10)
        self.worksheet.set_column(1, 1, self.config['TableHeader']["fxitureid_column_witdh"])  # 设置Fixtureid列宽
        self.worksheet.set_column(2, 2, self.config['TableHeader']['chanel_column_witdh'])  # 设置ChanelID列宽
        for i in range(self.column_number):
            self.worksheet.set_column(i + 3, i + 3, self.config['TableHeader']['image_column_witdh'])  # 设置图片列的宽度
            # print(i)
        self.worksheet.set_column(self.column_number + 3, self.column_number + 3,
                                  self.config['TableHeader']['resuilt_witdh'])  # 设置列的宽度
        pass

    def setRow_format(self):
        print('设置行高及填入数据')
        format = self.workbook.add_format(self.config['row_format'])
        self.worksheet.set_row(0, 40)
        self.worksheet.set_row(1, self.config['TableHeader']['header_row_hight'])
        for index in range(0, self.row_number):
            self.worksheet.set_row(index + 2, self.config['TableHeader']['image_row_hight'])
            self.worksheet.write(index + 2, 2, 'CH{0}'.format(index + 1), format)
            for column_index in range(0, self.column_number + 1):
                self.worksheet.write(index + 2, column_index + 3, '', format)

    # 向文档插入一张产品的图片
    def addProductImages(self):
        print('向文档插入一张产品的图片')
        imageconfig = {'x_offset': 30,
                       'y_offset': 20,
                       'x_scale': 1,
                       'y_scale': 1,
                       'object_position': 0,
                       'decorative': True,
                       'positioning': 2
                       }
        productimagespath = self.config['default']['productimagepath']
        if (len(productimagespath) > 0 and os.path.exists(productimagespath)):
            self.worksheet.insert_image(1, self.column_number + 4, productimagespath, imageconfig)
        pass

    # 插入蓝膜照片
    def addProductBlueFilmImage(self, BlueFilmImageDic):
        print('向文档插入产品的蓝膜图片')
        imageconfig = {'x_offset': 10, 'y_offset': 10, 'x_scale': 6.8, 'y_scale': 7.3}
        for column_index in range(0, self.column_number):
            for row_index in range(0, self.row_number):
                imagefilepath = BlueFilmImageDic + '/{0}.{1}.bmp'.format(row_index + 1, column_index + 1)
                if os.path.exists(imagefilepath):
                    self.worksheet.insert_image(row_index + 2, column_index + 3, imagefilepath, imageconfig)
        pass

    # 设置报告格式
    def resetReportFormat(self):
        pass

    def addPriductBlueFilemImageResult(self):
        print('写入结果')
        format = self.workbook.add_format(self.config['resuilt_format'])
        value = ["PASS"]
        for row_index in range(0, self.row_number):
            self.worksheet.write_row(row_index + 2, self.column_number + 3, value, format)
        pass

    def closefile(self):
        print('报告生成完毕')
        self.workbook.close()
        os.system('open {0}'.format(writereport.REPORT_DIR))
        pass
