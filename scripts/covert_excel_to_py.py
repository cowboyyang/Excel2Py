#coding=utf-8

import xlrd
import os
import sys
from xml.dom import minidom
import optparse
import json

type_convert_map = {}
type_convert_map["int"] = long
type_convert_map["long"] = long
type_convert_map["string"] = str
type_convert_map["float"] = float

class Excel2PythonDataConverter:
    def __init__(self, excelname, excelsheet, outdir, targetfilename, messagemeta, xmlfie):
        self.excel = excelname
        self.sheet = excelsheet
        self.outdir = outdir
        self.targetfile = targetfilename
        self.metaname = messagemeta
        self.xmlfile = xmlfie
        self.metadict = {}

    def build_xml_dict(self):
        '''构造一个元数据字典'''
        domtree = minidom.parse(self.xmlfile)
        value = domtree.documentElement
        for node in value.childNodes:
            if node.nodeName == "struct":
                structname = node.getAttribute("name")
                self.metadict[structname] = {}
                for child in node.childNodes:
                    if child.nodeName == "entry":
                        cname = child.getAttribute("cname")
                        self.metadict[structname][cname] = {}
                        self.metadict[structname][cname]["name"]=child.getAttribute("name")
                        self.metadict[structname][cname]["type"]=child.getAttribute("type")
                        self.metadict[structname][cname]["option"]=child.getAttribute("option")

    def set_raw_filed(self, sheet, row, meta, key, itemmsg, leftkey="", rightkey=""):
        '''设置好一个属性'''
        keytype = meta[key].get("type")
        keyname = meta[key].get("name")
        pType = type_convert_map.get(keytype)
        properkey= leftkey + key + rightkey
        bFound = False
        bSetValue = False
        bStr = False
        for col in xrange(0, sheet.ncols):
            cname=sheet.cell_value(0, col)
            if cname == properkey:
                bFound = True
                value = sheet.cell_value(row, col)
                vlen = 0
                bStr = False
                if keytype == "string":
                    #value = value.encode('utf-8')
                    vlen = len(value)
                    bStr = True
                else:
                    if str(value) == "":
                        vlen = 0
                    else:
                        vlen = len(str(pType(value)))

                # 无数据，不写入字典
                if vlen > 0 and isinstance(itemmsg, dict):
                    if bStr:
                        itemmsg[keyname] = value
                    else:
                        itemmsg[keyname] = pType(value)
                    bSetValue = True
                elif vlen > 0 and isinstance(itemmsg, list):
                    itemmsg.append(pType(value))
                    bSetValue = True
                break

        if bFound is False:
            # 说明没有找到对应key字段
            return -1
        elif bSetValue is False:
            # 说明对应key数据为空
            return 1
        else:
            # 说明写入了对应key的数据
            return 0

    def gen_one_row_data(self, sheet, row):
        '''解析sheet中的第row行，生成python字典数据'''
        metadata = self.metadict.get(self.metaname)
        onerowdata = {}
        for key in metadata:
            keytype = metadata[key].get("type")
            keyname = metadata[key].get("name")
            option = metadata[key].get("option")

            if option == "repeated":
                # 说明是数组类型
                print "found repeated %s " % key
                array = []
                structmeta = self.metadict.get(keytype)
                if type_convert_map.get(keytype):
                    seq = 1
                    while True:
                        ret = self.set_raw_filed(sheet, row, metadata, key, array, rightkey=str(seq))
                        if ret < 0:
                            break
                        seq += 1
                else:
                    # 复合结构的数组类型
                    seq = 1
                    while True:
                        structitem = {}
                        for structkey in structmeta:
                            ret = self.set_raw_filed(sheet, row, structmeta,  structkey, structitem, leftkey=key+str(seq))
                            if ret < 0 :
                                break
                        if structitem:
                            array.append(structitem)
                        else:
                            # 一个值都没有设置上，终止
                            break
                        seq += 1
                if array:
                    onerowdata[keyname] = array
            else:
                # 非数组类型
                # 原始类型
                if type_convert_map.get(keytype):
                    self.set_raw_filed(sheet, row, metadata, key, onerowdata)
                else:
                    # 结构体类型
                    structmeta = self.metadict.get(keytype)
                    structitem = {}
                    for structkey in structmeta:
                        self.set_raw_filed(sheet, row, structmeta,  structkey, structitem, leftkey=key)
                    if structitem:
                        onerowdata[keyname] = structitem
        return onerowdata

    def convert_excel_to_python(self):
        '''将excel转换成python格式'''
        # 首先构建一个元数据字典
        self.build_xml_dict()
        # 读取excel
        workbook = xlrd.open_workbook(self.excel)
        sheet = workbook.sheet_by_name(self.sheet)

        # 生成python字典
        row_array_msg = []
        for row in xrange(1, sheet.nrows):
            onerow = self.gen_one_row_data(sheet, row)
            if onerow:
                row_array_msg.append(onerow)

        self.write_to_file(row_array_msg)

    def write_to_file(self, msg):
        str1 = json.dumps(msg, indent=2, ensure_ascii=False).encode('utf-8')
        visualfile = self.targetfile.split('.')[0] + ".py"
        targetvisualfile = visualfile
        if os.path.exists(targetvisualfile):
            # 如果有旧文件，先删除
            os.remove(targetvisualfile)
        print targetvisualfile
        handle = open(targetvisualfile, 'w')
        # 写入编码格式
        handle.writelines("#coding=utf-8\n")
        dictname = "configdata_" + self.metaname + " = \\" + "\n"
        handle.writelines(dictname)
        handle.writelines(str1)
        handle.flush()
        handle.close()

if __name__ == "__main__":
    # cmdline config info
    parser = optparse.OptionParser()
    parser.add_option("--xmlfile",  dest="xmlfile", help="process target xml files")
    parser.add_option("--outdir", dest="outdir", help="target file store dir")
    parser.add_option("--excelfile", dest="excelfile", help="excel file name")
    parser.add_option("--sheetname", dest="sheetname", help="excel sheet name")
    parser.add_option("--messagemeta", dest="messagemeta", help="message meta data")
    parser.add_option("--dataname", dest="dataname", help="convert protobuf data name")

    (options, args) = parser.parse_args()
    procxmlfilelist = []
    if options.xmlfile is None:
        print "no input xml file"
        parser.print_help()
        exit(1)
    else:
        procxmlfilelist = options.xmlfile.split(" ")

    if options.outdir is None:
        print "need store target dir"
        parser.print_help()
        exit(1)

    outdir = os.path.abspath(options.outdir)
    excelfile = str(options.excelfile).strip()
    excelsheetname = str(options.sheetname).strip().decode("utf-8")
    targetfilename = str(options.dataname).strip().decode("utf-8")
    messagemeta = str(options.messagemeta).strip()
    msgxmlfile = procxmlfilelist[0]
    excelconvert = Excel2PythonDataConverter(excelfile,
                                             excelsheetname,
                                             outdir,
                                             targetfilename,
                                             messagemeta,
                                             msgxmlfile)
    excelconvert.convert_excel_to_python()