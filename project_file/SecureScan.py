import openpyxl as xml
import logging


global conf_xml
global scan_xml
global full_xml


def init_scan():
    global conf_xml
    global scan_xml
    global full_xml
    global logger
    # 设置日志级别， 低于某界别则不打印， 方便调试
    logging.basicConfig(level=logging.DEBUG)
    # 加载所有表格， conf为配置表格， scan被扫描表格， full_xml为全集表格
    conf_xml = xml.load_workbook("../start_dir/conf.xlsx").worksheets[0]
    scan_xml = xml.load_workbook(conf_xml.cell(2, 2).value)[conf_xml.cell(2, 3).value]
    full_xml = xml.load_workbook(conf_xml.cell(3, 2).value)[conf_xml.cell(3, 3).value]
    logging.debug(f"config xml:{conf_xml} scan xml: {scan_xml} full xml: {full_xml}")
    return


def scan():
    return

if __name__ == "__main__":
    init_scan()






