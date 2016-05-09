#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os, sys
import traceback
# sys.path.append("python_package/xlrd/lib/python")
# sys.path.append("python_package/xlwt/lib")
sys.path.append(os.path.join('python_package', 'xlrd', 'lib', 'python'))
sys.path.append(os.path.join('python_package', 'xlwt', 'lib'))
sys.path.append(os.path.join('python_package'))

import xlrd
import xlwt
from xlrd import open_workbook

import ntpath
import json

# create logger
#----------------------------------------------------------------------
import logging

# numeric_level = getattr(logging, loglevel.upper(), None)
# if not isinstance(numeric_level, int):
#       raise ValueError('Invalid log level: %s' % loglevel)

# logging.basicConfig(level=logging.INFO)

log = logging.getLogger('python_logger')
log.setLevel(logging.DEBUG)

fh = logging.FileHandler('out.log', 'w')
fh.setLevel(logging.DEBUG)
# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
# create formatter and add it to the handlers
# formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# 2015-08-28 17:01:57,662 - simple_example - ERROR - error message

# formatter = logging.Formatter('%(asctime)s %(levelname)-8s %(filename)s:%(lineno)-4d: %(message)s')
formatter = logging.Formatter('%(asctime)s %(levelname)-2s %(lineno)-4d: %(message)s')

fh.setFormatter(formatter)
ch.setFormatter(formatter)
# add the handlers to the logger
log.addHandler(fh)
log.addHandler(ch)
#----------------------------------------------------------------------

# 'application' code
# logger.debug('debug message')
# logger.info('info message')
# logger.warn('warn message')
# logger.error('error message')
# logger.critical('critical message')
#----------------------------------------------------------------------
# create logger end

MAX_ROW_NUM = 60000


def TestImport():
  log.info("test import success")

class Order(object):
    def __init__(self, order_ID, dest_name, dest_province, dest_city, dest_country, dest_addr, dest_phone_num, src_name, note):
        self.order_ID = order_ID
        self.dest_name = dest_name
        self.dest_province = dest_province
        self.dest_city = dest_city
        self.dest_country = dest_country
        self.dest_addr = dest_addr
        self.dest_phone_num = dest_phone_num
        self.src_name = src_name
        self.note = note
        self.date = unicode(int(note[4:6])) + u'月' + unicode(int(note[6:8])) + u'日'
        self.ID = note[0:4]

    def __str__(self):
        return("Order object:\n"
               "  order_ID = {0}\n"
               "  dest_name = {1}\n"
               "  dest_province = {2}\n"
               "  dest_city = {3}\n"
               "  dest_country = {4}\n"
               "  dest_addr = {5}\n"
               "  dest_phone_num = {6}\n"
               "  src_name = {7}\n"
               "  note = {8}\n"
               "  date = {9}\n"
               "  ID = {10}\n"
               .format(self.order_ID.encode('UTF-8'), self.dest_name.encode('UTF-8'),
                 self.dest_province.encode('UTF-8'), self.dest_city.encode('UTF-8'),
                 self.dest_country.encode('UTF-8'), self.dest_addr.encode('UTF-8'),
                 unicode(self.dest_phone_num)[:-2].encode('UTF-8'), self.src_name.encode('UTF-8'),
                 self.note, self.date.encode('UTF-8'), self.ID
                 ))

def WriteXls(path, element_list, element_info_list):
  workbook = xlwt.Workbook()
  sheet = workbook.add_sheet("right address")

  head_list = [u"订单号",u"商品名称",u"单位",u"数量",u"单价",u"重量",u"SKU编码", u"SKU名称", u"客户名称*", u"备注",u"收件人姓名",u"收件人省",u"收件人市",u"收件人区",u"收件人地址",u"收件人邮编",u"收件人电话",u"收件人手机",u"收件人邮箱",u"发件人姓名",u"发件人省",u"发件人市",u"发件人区",u"发件人地址",u"发件人邮编",u"发件人电话",u"发件人手机",u"发件人邮箱",u"扩展单号",u"批次号",u"大头笔",u"面单号",u"代收货款",u"到付款",u"网点ID"]
  for index, item in enumerate(head_list) :
    sheet.write(0, index, item)

  for row in range(1, len(element_list) + 1):
    sheet.write(row, head_list.index(u"数量"), 1)
    sheet.write(row, head_list.index(u"代收货款"), element_info_list[1])
    sheet.write(row, head_list.index(u"客户名称*"), element_info_list[2].decode('GBK'))
    sheet.write(row, head_list.index(u"备注"), element_info_list[0] + element_info_list[3] + "%04d"%row)
    sheet.write(row, head_list.index(u"收件人姓名"), element_list[row - 1 ].dest_name)
    sheet.write(row, head_list.index(u"收件人地址"), element_list[row - 1].addr)
    sheet.write(row, head_list.index(u"收件人电话"), unicode(element_list[row - 1].phone_num)[:-2])
    sheet.write(row, head_list.index(u"发件人姓名"), element_list[row - 1].src_name)
  workbook.save(path)
  log.info("finish WriteXls:{}, row numbers:{}".format(path, len(element_list)))


def WriteOrderToSheet(work_book, sheet_name, element_list) :
  sheet = work_book.add_sheet(sheet_name)

  head_list = [u'发件日期', u"快递单号", u"收件人姓名", u"收件人省", u"收件人市", u"收件人区", u"收件人地址", u"收件人电话", u"客服", u"客户编码"]

  for index, item in enumerate(head_list) :
    sheet.write(0, index, item)

  for row in range(1, len(element_list) + 1):
    sheet.write(row, 0, element_list[row - 1].date)
    sheet.write(row, 1, element_list[row - 1].order_ID)
    sheet.write(row, 2, element_list[row - 1].dest_name)
    sheet.write(row, 3, element_list[row - 1].dest_province)
    sheet.write(row, 4, element_list[row - 1].dest_city)
    sheet.write(row, 5, element_list[row - 1].dest_country)
    sheet.write(row, 6, element_list[row - 1].dest_addr)
    sheet.write(row, 7, element_list[row - 1].dest_phone_num)
    sheet.write(row, 8, element_list[row - 1].src_name)
    sheet.write(row, 9, element_list[row - 1].note)

  return work_book


def WriteListToXls(path, sheet_name, element_list):
  workbook = xlwt.Workbook()
  sheet = workbook.add_sheet(sheet_name)
  col = 0
  for index, item in enumerate(element_list) :
    if (index % (MAX_ROW_NUM -1) == 0 and index > 0) :
      col = col + 1
    sheet.write(index % (MAX_ROW_NUM - 1), col, item)

  workbook.save(path)
  log.info("finish WriteXls: %s, element numbers is: %s", path, len(element_list))

def GetFileList(path):
  from os import listdir
  from os.path import isfile, join
  onlyfiles = [ join(path, f) for f in listdir(path) if isfile(join(path, f))]
  return onlyfiles

def PathLeaf(path):
  head, tail = ntpath.split(path)
  return tail or ntpath.basename(head)

# read the xls files

def ParseSheetToList(sheet) :
  ret_list = []
  number_of_rows = sheet.nrows
  number_of_columns = sheet.ncols
  if (number_of_rows <= 0 or number_of_columns <= 0):
    log.error('can not get elem from sheet')
    return ret_list

  try:
    for col in range(number_of_columns):
      for row in range(number_of_rows):
        value = sheet.cell(row, col).value

        if (type(value) is float and value > 0):
          ret_list.append("%.0f"%value)
        elif (type(value) is unicode and len(value) > 0):
          ret_list.append(value)


  except Exception:
    log.error("Got exception on ParseSheetToList:%s", traceback.format_exc() )
  return ret_list

def ReadFinancialFile(path) :
  order_receive_list = []
  order_reject_list = []

  ret_dict = {"receive_list" : order_receive_list, "reject_list" : order_reject_list}

  file_list = GetFileList(path)
  if len(file_list) == 0:
    log.error("can not find file in path:%s", path)
    return ret_dict, len(file_list)

  for file in file_list:
    log.info("open file: %s", file)

    wb = open_workbook(file)
    for sheet in wb.sheets():

      if (sheet.name == u"签收"):
        order_receive_list.extend(ParseSheetToList(sheet))
        log.debug("order_receive_list.size:%d", len(order_receive_list))
      elif (sheet.name == u"退回"):
        order_reject_list.extend(ParseSheetToList(sheet))
        log.debug("order_reject_list.size:%d", len(order_reject_list))

  ret_dict["receive_list"] = order_receive_list
  ret_dict["reject_list"] = order_reject_list
  return ret_dict, len(file_list)

def GetWholeOrderSet(path):
  log.debug("GetWholeOrderSet, path:%s", path)
  file_list = GetFileList(path)
  whole_receive_order_list = []
  whole_reject_order_list = []

  if len(file_list) == 0:
    log.error("can not find file in path:%s", path)
    return {}
  for file in file_list:
    log.info("open file: %s", file)

    if u'总签收' in file:
      wb = open_workbook(file)
      for sheet in wb.sheets():
        whole_receive_order_list.extend(ParseSheetToList(sheet))
    if u'总退回' in file:
      wb = open_workbook(file)
      for sheet in wb.sheets():
        whole_reject_order_list.extend(ParseSheetToList(sheet))

  ret_dict = {"whole_receive_set" : set(whole_receive_order_list), "whole_reject_set" : set(whole_reject_order_list)}
  log.info("whole_receive_order_list.size:%d", len(set(whole_receive_order_list)))
  log.info("whole_reject_order_list.size:%d", len(set(whole_reject_order_list)))
  return ret_dict

def WriteWholeReceiveOrderListToFile(path, receive_order_list) :
  log.info("GetWholeRejectOrderList, path:%s, receive_order_list.size:%d", path, len(receive_order_list))
  WriteListToXls(path, u'总签收', receive_order_list)

def WriteWholeRejectOrderListToFile(path, reject_order_list) :
  log.info("GetWholeRejectOrderList, path:%s, reject_order_list.size:%d", path, len(reject_order_list))
  WriteListToXls(path, u'总退回', reject_order_list)

def AddFinancialOrderToWholeOrderList(path = u"单日签收退回from财务"):
  finanical_order_dict, file_count = ReadFinancialFile(path)

  if (len(finanical_order_dict["receive_list"]) == 0 and len(finanical_order_dict["reject_list"]) == 0) :
    log.error("can not get order_list from financial path:%s", path)
    return 0, file_count

  whole_order_dict = GetWholeOrderSet(u'签收退回总订单号')
  whole_order_dict["whole_receive_set"] = whole_order_dict["whole_receive_set"] | set(finanical_order_dict["receive_list"])
  whole_order_dict["whole_reject_set"] = whole_order_dict["whole_reject_set"] | set(finanical_order_dict["reject_list"])

  WriteListToXls(os.path.join(u'签收退回总订单号', u'总签收.xls'), u'总签收',
      sorted(list(whole_order_dict["whole_receive_set"])))
  WriteListToXls(os.path.join(u'签收退回总订单号', u'总退回.xls'), u'总退回',
      sorted(list(whole_order_dict["whole_reject_set"])))

  return len(finanical_order_dict["receive_list"]) + len(finanical_order_dict["reject_list"]), file_count

def ProcessRowBackOrderToBackOrder(path = u'待处理回单'):
  file_list = GetFileList(path)
  if len(file_list) == 0:
    log.error("can not find file in path:%s", path)
    return 0

  for file in file_list:
    log.info("open file: %s", file)

    rd_workbook = open_workbook(file)
    wt_workbook = xlwt.Workbook()
    head_list = [u"快递单号", u"收件人姓名", u"收件人省", u"收件人市", u"收件人区", u"收件人地址", u"收件人电话", u"发件人姓名", u"备注"]

    for rd_sheet in rd_workbook.sheets():

      number_of_rows = rd_sheet.nrows
      number_of_columns = rd_sheet.ncols

      rd_head_list = []

      if (number_of_rows <= 0 or number_of_columns <= 0):
        continue

      wt_sheet = wt_workbook.add_sheet(rd_sheet.name)

      for column in range(number_of_columns) :
        rd_head_list.append(rd_sheet.cell(0, column).value)

      for index, item in enumerate(head_list) :
        wt_sheet.write(0, index, item)

      for row in range(1, number_of_rows) :
        wt_sheet.write(row, head_list.index(u"快递单号"), rd_sheet.cell(row, rd_head_list.index(u"快递单号")).value)
        wt_sheet.write(row, head_list.index(u"收件人姓名"), rd_sheet.cell(row, rd_head_list.index(u"收件人姓名")).value)
        wt_sheet.write(row, head_list.index(u"收件人省"), rd_sheet.cell(row, rd_head_list.index(u"收件人省")).value)
        wt_sheet.write(row, head_list.index(u"收件人市"), rd_sheet.cell(row, rd_head_list.index(u"收件人市")).value)
        wt_sheet.write(row, head_list.index(u"收件人区"), rd_sheet.cell(row, rd_head_list.index(u"收件人区")).value)
        wt_sheet.write(row, head_list.index(u"收件人地址"), rd_sheet.cell(row, rd_head_list.index(u"收件人地址")).value)
        wt_sheet.write(row, head_list.index(u"收件人电话"), rd_sheet.cell(row, rd_head_list.index(u"收件人电话")).value)
        wt_sheet.write(row, head_list.index(u"发件人姓名"), rd_sheet.cell(row, rd_head_list.index(u"发件人姓名")).value)
        wt_sheet.write(row, head_list.index(u"备注"), rd_sheet.cell(row, rd_head_list.index(u"备注")).value)

    path = os.path.join(u'已处理回单', u'已处理回单_' + PathLeaf(file))
    wt_workbook.save(path)
    log.debug("write file: %s", path)
  return len(file_list)



def ProcessRowBackOrderToRecordOrder(path = u'待处理回单'):
  file_list = GetFileList(path)
  if len(file_list) == 0:
    log.error("can not find file in path:%s", path)
    return 0

  elem_list = [u'小李', '17090152657', u'河南', u'信阳', u'浉河区', u'河南信阳市浉河区平西涵洞', u'张海霞', '15293143927', u'甘肃省', u'兰州市', u'榆中县', u'详情见面单', u'物品', 0, 0]

  for file in file_list:
    log.info("open file: %s", file)

    wt_workbook = xlwt.Workbook()
    wt_sheet = wt_workbook.add_sheet("sheet1")

    rd_workbook = open_workbook(file)
    # head_list = [u"快递单号", u"收件人姓名", u"收件人省", u"收件人市", u"收件人区", u"收件人地址", u"收件人电话", u"发件人姓名", u"备注"]

    for rd_sheet in rd_workbook.sheets():

      number_of_rows = rd_sheet.nrows
      number_of_columns = rd_sheet.ncols

      if (number_of_rows <= 0 or number_of_columns <= 0):
        continue

      rd_head_list = []
      for column in range(number_of_columns) :
        rd_head_list.append(rd_sheet.cell(0, column).value)

      # for index, item in enumerate(head_list) :
      #   wt_sheet.write(0, index, item)

      for row in range(1, number_of_rows) :
        wt_sheet.write(row, 0, rd_sheet.cell(row, rd_head_list.index(u"快递单号")).value)

        for idx, elem in enumerate(elem_list):
          wt_sheet.write(row, idx + 2, elem)

        wt_sheet.write(row, 17, rd_sheet.cell(row, rd_head_list.index(u"代收货款")).value)

    path = os.path.join(u'录单', u'录单_' + PathLeaf(file))
    wt_workbook.save(path)
    log.info("finish WriteXls: %s, element numbers is: %s", path, number_of_rows)
  return len(file_list)

def ParseBackOrderToBill(path = u'已处理回单') :
  log.info("ParseBackOrderToBill, path:%s", path)

  file_list = GetFileList(path)
  if len(file_list) == 0:
    log.error("can not find file in path:%s", path)

  whole_order_dict = GetWholeOrderSet(u'签收退回总订单号')

  for file in file_list:
    log.info("open file: %s", file)

    # get order book list
    order_list = []
    order_receive_list = []
    order_reject_list = []
    order_unknown_list = []
    rd_workbook = open_workbook(file)
    for sheet in rd_workbook.sheets():
      number_of_rows = sheet.nrows
      number_of_columns = sheet.ncols

      if (number_of_rows <= 0 or number_of_columns <= 0):
        continue

      for row in range(1, number_of_rows) :
        values = []

        for col in range(number_of_columns) :
          value = sheet.cell(row,col).value
          if (type(value) is unicode or type(value) is str):
            value = value.replace(u'\xa0', "").replace(u" ", "").replace(u'\uff0c',"").replace(u",","")
          values.append(value)

        order = Order(*values)
        order_list.append(order)

    if len(order_list) <= 0 :
      continue

    log.debug("get order_list number:%d", len(order_list))

    for order in order_list:
      log.debug(order)
      break

    # check order status
    for order in order_list:
      if order.order_ID in whole_order_dict["whole_receive_set"]:
        order_receive_list.append(order)
      elif order.order_ID in whole_order_dict["whole_reject_set"]:
        order_reject_list.append(order)
      else:
        order_unknown_list.append(order)

    # write order to bill xls
    wt_workbook = xlwt.Workbook()

    sheet = wt_workbook.add_sheet(u'总汇')
    for index, item in enumerate([u'客户编码', u'发货日期', u'总发货量', u'签收件', u'退回件', u'无状态']) :
      sheet.write(0, index, item)

    sheet.write(1, 0, order_list[0].ID)
    sheet.write(1, 1, order_list[0].date)
    sheet.write(1, 2, len(order_list))
    sheet.write(1, 3, len(order_receive_list))
    sheet.write(1, 4, len(order_reject_list))
    sheet.write(1, 5, len(order_unknown_list))

    wt_workbook = WriteOrderToSheet(wt_workbook, u'签收件', order_receive_list)
    wt_workbook = WriteOrderToSheet(wt_workbook, u'退回件', order_reject_list)
    wt_workbook = WriteOrderToSheet(wt_workbook, u'无状态', order_unknown_list)

    # write bill file
    path = os.path.join(u'账单', u'账单_' + PathLeaf(file))
    wt_workbook.save(path)
    log.info("finish WriteXls: %s, row numbers: %d", path, len(order_list))
  return len(file_list)

def TestWriteXLS(end_num):
  wt_workbook = xlwt.Workbook()
  wt_sheet = wt_workbook.add_sheet("test1")

  for i, elem in enumerate(range(1, end_num)):
    wt_sheet.write(i, 1, elem)

  wt_workbook.save("test.xls")


def Start() :
  while (True):
    log.info("\n\n\n")
    log.info("*******************************")
    log.info(u"1. 将财务订单号导入总订单号")
    log.info(u"2. 将系统导出回单处理成回单")
    log.info(u"3. 将回单处理成账单")
    log.info(u"4. 将回单处理成录单")
    log.info(u"5. 1->2->3->4 依次完成")
    log.info(u"6. 退出")
    log.info("*******************************")
    log.info(u"输入你想选择的功能(1 ~ 6):")

    try :
      choose_num = int(raw_input())
      log.info("\n\n\n")

      log.info(u'你的选择是: %d', choose_num)

      if (choose_num == 1):
        log.info(u"1. 将财务订单号导入总订单号 开始.")
        input_num, file_num = AddFinancialOrderToWholeOrderList()
        log.info(u"1. 将财务订单号导入总订单号 完成, 导入数量: %d, 导入文件数: %d", input_num, file_num)

      elif (choose_num == 2):
        log.info(u"2. 将系统导出回单处理成回单 开始.")
        file_num = ProcessRowBackOrderToBackOrder()
        log.info(u"2. 将系统导出回单处理成回单 完成, 处理文件数:%d", file_num)

      elif (choose_num == 3):
        log.info(u"3. 将回单处理成账单 开始.")
        file_num = ParseBackOrderToBill()
        log.info(u"3. 将回单处理成账单 完成, 处理文件数:%d", file_num)

      elif (choose_num == 4):
        log.info(u"4. 将回单处理成录单 开始.")
        file_num = ProcessRowBackOrderToRecordOrder()
        log.info(u"4. 将回单处理成录单 完成, 处理文件数:%d", file_num)

      elif (choose_num == 5):
        log.info(u"1. 将财务订单号导入总订单号 开始.")
        input_num, file_num = AddFinancialOrderToWholeOrderList()
        log.info(u"1. 将财务订单号导入总订单号 完成, 导入数量: %d, 导入文件数: %d", input_num, file_num)
        log.info("\n\n")

        log.info(u"2. 将系统导出回单处理成回单 开始.")
        file_num = ProcessRowBackOrderToBackOrder()
        log.info(u"2. 将系统导出回单处理成回单 完成, 处理文件数:%d", file_num)
        log.info("\n\n")

        log.info(u"3. 将回单处理成账单 开始.")
        file_num = ParseBackOrderToBill()
        log.info(u"3. 将回单处理成账单 完成, 处理文件数:%d", file_num)
        log.info("\n\n")

        log.info(u"4. 将回单处理成录单 开始.")
        file_num = ProcessRowBackOrderToRecordOrder()
        log.info(u"4. 将回单处理成录单 完成, 处理文件数:%d", file_num)

      elif (choose_num == 6):
        break
      else:
        log.info(u'输入非法, 请输入数字: 1~6')
        continue
    except Exception:
      log.error("Got exception on ParseSheetToList:%s", traceback.format_exc() )
      continue


if __name__ == "__main__":
  Start()



