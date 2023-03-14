# -*- coding: utf-8 -*-
import os
# from btreport import btformat
from lib.writereport import writereport
import argparse


def getfixtureimagedir(allfixtureimgaedir):
    dirconfig = {}
    for fixtureName in os.listdir(allfixtureimgaedir):
        if (fixtureName != ".DS_Store"):
            fixtureimagedir = os.path.join(allfixtureimgaedir, fixtureName)
            dirconfig[fixtureName] = fixtureimagedir
    return dirconfig


def writebtreport(imagedir, reprottype):
    print(imagedir, reprottype)
    dirconfig = getfixtureimagedir(imagedir)
    for fixturename, imagedir in dirconfig.items():
        print(fixturename)
        print(imagedir)
        report = writereport(fixturename, reprottype)
        # report.setReportTyep(reprottype)
        report.add_tableHandle()
        report.setMerge_format()
        report.setColumn_format()
        report.setRow_format()
        report.addProductImages()
        report.addProductBlueFilmImage(imagedir)
        report.addPriductBlueFilemImageResult()
        report.closefile()
    pass


def main():
    parser = argparse.ArgumentParser(description="example: \
        python3 PS_Analysis_main.py -ch 0")
    parser.add_argument("-type", type=str, default='DFU', help='报告类型 FCT DFU 或者PDFUT', required=True)
    parser.add_argument('-d', help='蓝膜图片目录', type=str, default="./GK_image")
    agrs = parser.parse_args()
    Image_dir_Path = agrs.d
    Type_str = agrs.type
    print('报告类型：{0}'.format(Type_str))
    writebtreport(Image_dir_Path, Type_str)
    #
    # os.system('open {0}'.format(REPORT_DIR))

    pass


if __name__ == '__main__':
    main()
