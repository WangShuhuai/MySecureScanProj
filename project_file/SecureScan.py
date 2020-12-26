import openpyxl as xml
import logging


global conf_xml
global scan_xml
global full_xml


def init_scan():
    global conf_xml
    global scan_xml
    global full_xml
    # 设置日志级别， 低于某界别则不打印， 方便调试
    logging.basicConfig(level=logging.DEBUG)
    # 加载所有表格， conf为配置表格， scan被扫描表格， full_xml为全集表格
    conf_xml = xml.load_workbook("../start_dir/conf.xlsx").worksheets[0]
    scan_xml = xml.load_workbook(conf_xml.cell(2, 2).value)[conf_xml.cell(2, 3).value]
    full_xml = xml.load_workbook(conf_xml.cell(3, 2).value)[conf_xml.cell(3, 3).value]
    logging.info(f"config xml:{conf_xml} scan xml: {scan_xml} full xml: {full_xml}")
    return


def scan():
    global conf_xml
    global scan_xml
    global full_xml
    logging.debug(conf_xml)
    logging.debug(scan_xml)
    logging.debug(full_xml)
    scan_match_col = conf_xml.cell(2, 4).value
    scan_fill_col = conf_xml.cell(2, 5).value
    full_match_col = conf_xml.cell(3, 4).value
    full_fill_col = conf_xml.cell(3, 5).value
    logging.info(f"scan_match_col: {scan_match_col}")
    logging.info(f"scan_fill_col: {scan_fill_col}")
    logging.info(f"full_match_col: {full_match_col}")
    logging.info(f"full_fill_col: {full_fill_col}")
    scan_match_col_list = scan_match_col.split(",")

    # 字典存储需要扫描的列
    scan_dic = {}
    for i in scan_match_col_list:
        scan_dic.setdefault(i, "")
        logging.debug(scan_dic)

    # Todo：获取全集字典列表
    
    # 获取逐行读取填到字典中，然后比对
    # 如果最大行配置为0， 则获取最大行， 否则获取配置值
    if int(conf_xml.cell(3, 6).value) == 0:
        scan_xml_max_row = scan_xml.max_row
    else:
        scan_xml_max_row = int(conf_xml.cell(3, 6).value)
    logging.info(scan_xml_max_row)

    # 起始行获取与最大行检查
    scan_start_row = int(conf_xml.cell(2, 6).value)
    if scan_start_row >= scan_xml_max_row:
        logging.error("最大行配置错误， 请重新配置！")
    for n in range(scan_start_row, scan_xml_max_row):
        for i in scan_match_col_list:
            scan_dic[i] = scan_xml.cell(n, int(i)).value
        logging.debug(scan_dic)
    return


if __name__ == "__main__":
    init_scan()
    scan()

