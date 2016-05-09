#!/usr/bin/python
# -*- coding: UTF-8 -*-

from order_analysis import *

try:
  # TestImport()
  # ReadFinancialFile(u"单日签收退回from财务")
  GetWholeOrderSet(u'签收退回总订单号')
  # TestWriteXLS(10)
except Exception:
  log.error("Got exception on ParseSheetToList:%s", traceback.format_exc() )
  raw_input("press Enter to exit")

def changeNUm(num):
  num += 1
  print "changeNUm", num

def test():
  num = 5
  changeNUm(num)
  print "test", num

test()
