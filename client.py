#!/usr/bin/python3
# coding=utf-8

import codecs
import os
import sys
import time
import traceback
import win32con
import win32evtlog
import win32evtlogutil
import winerror
from datetime import datetime


def getAllEvents(companyId, pickUplogTypes, basePath):
    """
    """
    if not companyId:
        serverName = "未列出主机"
    else:
        serverName = companyId
    for logtype in pickUplogTypes:
        path = os.path.join(basePath, "%s_%s_日志.log" % (serverName, logtype))
        print('输出日志位置为：', path)
        getEventLogs(companyId, logtype, path)


# ----------------------------------------------------------------------
def getEventLogs(companyId, logtype, logPath):
    """
    Get the event logs from the specified machine according to the
    logtype (Example: Application) and save it to the appropriately
    named log file
    """
    print("载入%s事件" % logtype)
    log = codecs.open(logPath, encoding='utf-8', mode='w')
    line_break = '-' * 80 + '\n'
    log.write("服务器：%s 日志事件类型：%s " % (companyId, logtype))
    log.write("创建时间: %s \n" % datetime.now().strftime("%Y-%m-%d %H:%S:%M"))
    log.write(line_break)
    # 读取本机的,system系统日志
    hand = win32evtlog.OpenEventLog(None, logtype)
    # 获取system日志的总行数
    total = win32evtlog.GetNumberOfEventLogRecords(hand)
    print("%s总日志数量为%s" % (logtype, total))
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    events = win32evtlog.ReadEventLog(hand, flags, 0)
    # 错误级别类型
    # evt_dict = {win32con.EVENTLOG_AUDIT_FAILURE: 'EVENTLOG_AUDIT_FAILURE',
    #             win32con.EVENTLOG_AUDIT_SUCCESS: 'EVENTLOG_AUDIT_SUCCESS',
    #             win32con.EVENTLOG_INFORMATION_TYPE: 'EVENTLOG_INFORMATION_TYPE',
    #             win32con.EVENTLOG_WARNING_TYPE: 'EVENTLOG_WARNING_TYPE',
    #             win32con.EVENTLOG_ERROR_TYPE: 'EVENTLOG_ERROR_TYPE'}

    evt_dict = {win32con.EVENTLOG_WARNING_TYPE: '警告',
                win32con.EVENTLOG_ERROR_TYPE: '错误',
                win32con.EVENTLOG_AUDIT_FAILURE: '审核失败',
                win32con.EVENTLOG_AUDIT_SUCCESS: '审核成功',
                win32con.EVENTLOG_INFORMATION_TYPE: '信息'
                }

    try:
        events = 1
        while events:
            events = win32evtlog.ReadEventLog(hand, flags, 0)

            for ev_obj in events:
                the_time = ev_obj.TimeGenerated.Format()  # '12/23/99 15:54:09'
                time_obj = datetime.strptime(the_time, '%c')
                evt_id = int(winerror.HRESULT_CODE(ev_obj.EventID))
                computer = str(ev_obj.ComputerName)
                cat = ev_obj.EventCategory
                # seconds=date2sec(the_time)
                record = ev_obj.RecordNumber
                msg = win32evtlogutil.SafeFormatMessage(ev_obj, logtype)
                source = str(ev_obj.SourceName)
                nowtime = datetime.today()
                daySpace = nowtime.__sub__(time_obj).days
                if not ev_obj.EventType in evt_dict.keys():
                    evt_type = "未知"
                else:
                    evt_type = str(evt_dict[ev_obj.EventType])
                if evt_type in pickUpEventTypes and evt_id in pickUpEventIDs and daySpace <= needDaySpace:
                    log.write("计算机名：%s\n" % computer)
                    log.write("记录编码：%s\n" % record)
                    log.write("事件时间: %s\n" % time_obj)
                    log.write("事件ID：%s | 事件类别: %s\n" % (evt_id, evt_type))
                    log.write("来源: %s\n" % source)
                    log.write(msg + '\n')
                    log.write(line_break)
    except:
        print(traceback.print_exc(sys.exc_info()))

    print("日志文件创建完成：%s" % logPath)


if __name__ == "__main__":
    companyId = 1204  # None = local machine
    needDaySpace = 7  # 需要几天前的事件日志
    pickUplogTypes = ["System"]  # 目前只获取日志名称为：系统的日志
    # EventTypes = ['win32con.EVENTLOG_WARNING_TYPE', 'win32con.EVENTLOG_ERROR_TYPE',
    # 'win32con.EVENTLOG_INFORMATION_TYPE', 'win32con.EVENTLOG_AUDIT_FAILURE', 'win32con.EVENTLOG_AUDIT_SUCCESS']
    pickUpEventTypes = ['错误']  # 需要提取的类别列表
    pickUpEventIDs = [7, 8, 9, 14, 6008]  # 需要提取的事件ID列表
    # pickUplogTypes = ["System", "Application", "Security"]
    getAllEvents(companyId, pickUplogTypes, ".\\")  # todo需要修改输出位置
