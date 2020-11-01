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
import configparser
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime


def execlog():
    detailIniPath = 'detail.ini'
    if os.path.exists(detailIniPath):
        getdetaiconfig = configparser.ConfigParser()
        getdetaiconfig.read("detail.ini", encoding="utf-8")
        ts = set()
        for i in getdetaiconfig:
            if i == 'DEFAULT':
                continue
            getdetail = eval(getdetaiconfig[i]['detail'])
            print(len(getdetail))
            for k in (getdetail):
                t = getdetail[k]
                if t not in ts:
                    ts.add({k: t})
            getdetaiconfig.set(str(i), 'sentstatus', str(ts))
        with open(detailIniPath, 'w', encoding='utf-8') as detailIniPath:
            getdetaiconfig.write(detailIniPath)
    else:
        print("detail.ini不存在，退出!")
        sys.exit()


def getAllEvents(companyId, logType, eventIDs, logPath, needDaySpace):
    """
    """
    if not companyId:
        serverName = "未列出主机"
    else:
        serverName = companyId
    for logtype in logType:
        path = os.path.join(logPath, "%s-%s-日志.log" % (serverName, logtype))
        print('输出日志位置为：', path)
        getEventLogs(companyId, logtype, eventIDs, path, needDaySpace)

def getEventLogs(companyId, logtype, eventIDs, path, needDaySpace):
    """
    Get the event logs from the specified machine according to the
    logtype (Example: Application) and save it to the appropriately
    named log file
    """
    print("载入%s事件" % logtype)
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
        log = codecs.open(path, encoding='utf-8', mode='w')
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
                try:
                    msg = win32evtlogutil.SafeFormatMessage(ev_obj, logtype)
                except:
                    msg = ''
                source = str(ev_obj.SourceName)
                nowtime = datetime.today()
                daySpace = nowtime.__sub__(time_obj).days
                if not ev_obj.EventType in evt_dict.keys():
                    evt_type = "未知"
                else:
                    evt_type = str(evt_dict[ev_obj.EventType])
                if evt_type in eventType and str(evt_id) in eventIDs and int(daySpace) <= int(needDaySpace):
                    detailIniPath = 'detail.ini'
                    if not os.path.exists(detailIniPath):
                        detaiconfig = configparser.ConfigParser()
                        detaiconfig['{}'.format(evt_id)] = {'evt_id': evt_id,
                                                            'evt_type': evt_type,
                                                            'source': source,
                                                            'msg': msg.strip(),
                                                            'detail': {time_obj.strftime("%Y-%m-%d %H:%M"): record},
                                                            "sentstatus": ''
                                                            }
                        with open(detailIniPath, 'w', encoding='utf-8') as detailIniPath:
                            detaiconfig.write(detailIniPath)
                    else:
                        detaiconfig = configparser.ConfigParser()
                        detaiconfig.read("detail.ini", encoding="utf-8")
                        if detaiconfig.has_section(str(evt_id)):
                            detail = eval(detaiconfig[str(evt_id)]['detail'])
                            if time_obj.strftime("%Y-%m-%d %H:%M") not in detail.keys():
                                detail[time_obj.strftime("%Y-%m-%d %H:%M")] = record
                                detaiconfig.set(str(evt_id), 'detail', str(detail))
                                with open(detailIniPath, 'w', encoding='utf-8') as detailIniPath:
                                    detaiconfig.write(detailIniPath)
                        else:
                            detaiconfig.add_section(str(evt_id))
                            detaiconfig.set(str(evt_id), 'evt_id', str(evt_id))
                            detaiconfig.set(str(evt_id), 'evt_type', evt_type)
                            detaiconfig.set(str(evt_id), 'source', source)
                            detaiconfig.set(str(evt_id), 'msg', msg.strip())
                            detaiconfig.set(str(evt_id), 'detail', str({time_obj.strftime("%Y-%m-%d %H:%M"): record}))
                            detaiconfig.set(str(evt_id), 'sentstatus', '')
                            with open(detailIniPath, 'w', encoding='utf-8') as detailIniPath:
                                detaiconfig.write(detailIniPath)
                    log.write("计算机名：%s" % computer)
                    log.write("记录编码：%s" % record)
                    log.write("事件时间: %s" % time_obj)
                    log.write("事件ID：%s | 事件类别: %s" % (evt_id, evt_type))
                    log.write("来源: %s" % source)
                    log.write(msg + '')
    except:
        print(traceback.print_exc(sys.exc_info()))

    print("日志文件创建完成：%s" % path)
    print("")


def sendmail(companyId, mail_host, port, mail_user, mail_pass, sender, receivers, From, To, getdetaiconfig):
    # 第三方 SMTP 服务
    # mail_host = ""  # 设置服务器
    # mail_user = ""  # 用户名
    # mail_pass = ""  # 口令
    #
    # sender = ''
    # receivers = ['']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
    idCount = len(getdetaiconfig.sections())
    sended = set()
    mail_msg = '''
                <h1>预警邮件来自---{}</h1>
        <ul>
           '''.format(companyId)
    ncount = 1
    for section in getdetaiconfig.sections():
        eevt_id = getdetaiconfig[section]['evt_id']
        eevt_type = getdetaiconfig[section]['evt_type']
        esource = getdetaiconfig[section]['source']
        emsg = getdetaiconfig[section]['msg']
        edetail = eval(getdetaiconfig[section]['detail'])
        esentstatus = getdetaiconfig[section]['sentstatus']
        count = len(dict(edetail))
        mail_msg = mail_msg + """
                     <h2>事件类型合计：{} 个 当前序列 {} / {} 的事件ID = {}
                </h2>
                <li>
                    级别 = {}
                </li>
                <li>
                    来源 = {}
                </li>
                <li>消息 = {}</li>
                <li>合计： {} 条</li>
                <li>记录编码与事件
        """.format(idCount, ncount, idCount, eevt_id, eevt_type, esource, str(emsg), count)
        for dt in edetail:
            if not esentstatus:
                mail_msg += '''
                <p>记录编码：{} 记录时间：{}</p>
                '''.format(edetail[dt], dt)
                aa = str(section) + '-' + str(dt)
                sended.add(aa)
                getdetaiconfig.set(str(section), 'sentstatus', str(sended))
            else:
                aa = str(section) + '-' + str(dt)
                if str(aa) not in esentstatus:
                    mail_msg += '''
                    <p>记录编码：{} 记录时间：{}</p>
                    '''.format(dt, edetail[dt])
                    sended.add(aa)
                    getdetaiconfig.set(str(section), 'sentstatus', str(sended))
        mail_msg += '''</li></ul>'''
        ncount += 1
    detailIniPath = 'detail.ini'
    with open(detailIniPath, 'w', encoding='utf-8') as detailIniPath:
        getdetaiconfig.write(detailIniPath)
    if sended:
        print("检测到有数据,开始发送邮件......")
        message = MIMEText(mail_msg, 'html', 'utf-8')
        message['From'] = Header(From, 'utf-8')
        subject = '预警邮件来自---{}'.format(companyId)
        message['Subject'] = Header(subject, 'utf-8')
        if isinstance(receivers, str):
            message["To"] = receivers
        elif isinstance(receivers, list):
            message['To'] = ";".join(receivers)
        try:
            if int(port) == 465:
                smtpSSLObj = smtplib.SMTP_SSL(mail_host, int(port))
                smtpSSLObj.login(mail_user, mail_pass)
                smtpSSLObj.sendmail(sender, receivers.split(","), message.as_string())
                print("邮件发送成功")
            elif int(port) == 25:
                smtpObj = smtplib.SMTP()
                smtpObj.connect(mail_host, int(port))
                smtpObj.login(mail_user, mail_pass)
                smtpObj.sendmail(sender, receivers.split(","), message.as_string())
                print("邮件发送成功")
        except smtplib.SMTPException as e:
            print("Error: 无法发送邮件")
            print(e)
    else:
        print("检测到没有数据,不发送邮件,跳过......")


if __name__ == "__main__":
    while True:
        configIniPath = 'config.ini'
        if not os.path.exists(configIniPath):

            # DEFAULT
            IP, barName, serverNotes, loopTime = '', '', '', ''
            while not IP:
                IP = input("请输入IP地址,按回车继续 *必填项*：\n")
            while not barName:
                barName = input("请输入网吧名称,按回车继续 *必填项*：\n")
            while not serverNotes:
                serverNotes = input("请输入服务器备注,按回车继续 *必填项*：\n")
            while not loopTime:
                loopTime = input("请输入脚本循环时间(分钟),按回车继续(直接回车默认值=5)：\n")

            # DETAIL
            needDaySpace, logType, eventType, eventIDs, logPath = '', '', '', '', ''
            while not needDaySpace:
                needDaySpace = input("请输入匹配至今多少天内的日志,按回车继续(直接回车默认值=7)：\n")
            while not logType:
                logType = input("请输入事件日志(多个值用,隔开),按回车继续(直接回车默认值=System)：\n")
            while not eventType:
                eventType = input("请输入事件级别(多个值用,隔开),按回车继续(直接回车默认值=错误)：\n")
            while not eventIDs:
                eventIDs = input("请输入事件ID列表(多个值用,隔开),按回车继续(直接回车默认值=7,8,9,11,14)：\n")
            while not logPath:
                logPath = input("请输入日志路径（\\必须为\\\\）,按回车继续(直接回车默认值=D:\\错误日志文件)：\n")

            # MAIL
            mail_host, port, mail_user, mail_pass, sender, receivers, From, To, = '', '', '', '', '', '', '', ''
            while not mail_host:
                mail_host = input("请输入smtp地址,按回车继续 *必填项* ：\n")
            while not port:
                port = input("请输入port地址,按回车继续 *必填项* ：\n")
            while not mail_user:
                mail_user = input("请输入邮箱用户名,按回车继续 *必填项* ：\n")
            while not mail_pass:
                mail_pass = input("请输入邮箱密码或授权码,按回车继续 *必填项* ：\n")
            while not sender:
                sender = input("请输入发件人邮箱地址,按回车继续 *必填项* ：\n")
            while not receivers:
                receivers = input("请输入收人邮箱地址（多个值用,隔开）,按回车继续 *必填项* ：\n")
            while not From:
                From = input("请输入发件人,按回车继续 *必填项* ：\n")
            while not To:
                To = input("请输入收件人(多个值用,隔开),按回车继续 *必填项* ：\n")

            config = configparser.ConfigParser()
            config["DEFAULT"] = {'; 备注：': 'IP=IP地址, barName=网吧名称,serverNotes=服务器备注, loopTime=脚本循环时间(分钟) ',
                                 'IP': IP,
                                 'barName': barName,
                                 'serverNotes': serverNotes,
                                 'loopTime': loopTime}
            config["DETAIL"] = {'; 备注：': 'needDaySpace=需要匹配至今多少天内的日志, logType=事件日志, eventType=事件级别, event'
                                         'IDs=事件ID列表，'
                                         'logPath＝日志路径 ',
                                'needDaySpace': needDaySpace,
                                'logType': logType,
                                'eventType': eventType,
                                'eventIDs': eventIDs,
                                'logPath': logPath
                                }
            config["MAIL"] = {'; 备注：': 'mail_host＝smtp地址, port=smtp地址(默认25), mail_user=邮箱用户名, mail_pass=邮箱密码'
                                       'sender=发送人邮件地址, receivers=收件人邮件地址, from=发送人姓名, to=接收人姓名',
                              'mail_host': mail_host,
                              'port': port,
                              "mail_user": mail_user,
                              "mail_pass": mail_pass,
                              "sender": sender,
                              "receivers": receivers,
                              "from": From,
                              "to": To
                              }
            with open(configIniPath, 'w', encoding='utf-8') as configfile:
                config.write(configfile)

            print("配置完成,已生成默认配置文件：{}，请关闭程序后再次打开！".format(configIniPath))
            out = input("按回车退出\n")
            sys.exit()
        else:
            print("存在配置文件")
            # 获取配置文件信息
            config = configparser.ConfigParser()
            config.read("config.ini", encoding="utf-8")
            IP = config["DEFAULT"]['IP']
            barName = config["DEFAULT"]['barName']
            serverNotes = config["DEFAULT"]['serverNotes']
            needDaySpace = int(config["DETAIL"]["needDaySpace"])
            logType = config["DETAIL"]["logType"].split(',')
            eventType = config["DETAIL"]["eventType"].split(',')
            eventIDs = config["DETAIL"]["eventIDs"].split(',')
            logPath = config["DETAIL"]["logPath"]
            companyId = barName + '-' + serverNotes + '-' + IP
            mail_host = config["MAIL"]["mail_host"]
            port = config["MAIL"]["port"]
            mail_user = config["MAIL"]["mail_user"]
            mail_pass = config["MAIL"]["mail_pass"]
            sender = config["MAIL"]["sender"]
            receivers = config["MAIL"]["receivers"]
            From = config["MAIL"]["from"]
            To = config["MAIL"]["to"].split(',')
            if len(To) == 1:
                To = To[0]
            looptime = config["DEFAULT"]['looptime']
            looptime = int(looptime) * 60
            if not os.path.isdir(logPath):
                os.makedirs(logPath)
            # 输出核对
            print("----------------------------------基础配置----------------------------------\n")
            print("IP地址(IP):{} \n网吧名称(barName):{} \n服务器备注(serverNotes):{}\n".format(IP, barName, serverNotes))
            print("----------------------------------详细配置----------------------------------\n")
            print('需要几天前的事件日志(needDaySpace):{}\n 日志类型(logType):{}\n事件类别(eventType):{}\n'
                  '事件ID列表(EventIDs):{}\n 输出日志路径(logPath):{}\n'.format(needDaySpace, logType, eventType,
                                                                      eventIDs, logPath))
            print("----------------------------------邮件配置----------------------------------\n")
            print("开始获取系统日志......")
            # getAllEvents() 获取系统日志函数
            getAllEvents(companyId, logType, eventIDs, logPath, needDaySpace)  # todo需要修改输出位置
            print("开始处理系统日志......")
            # execlog() 处理系统日志函数
            execlog()
            print("开始处理邮件......")
            detailIniPath = 'detail.ini'
            if os.path.exists(detailIniPath):
                getdetaiconfig = configparser.ConfigParser()
                getdetaiconfig.read("detail.ini", encoding="utf-8")
                # sendmail() 发送邮件函数
                sendmail(companyId, mail_host, port, mail_user, mail_pass, sender, receivers, From, To, getdetaiconfig)
            else:
                print("detail.ini不存在，退出!")
                sys.exit()

            print('暂停 {} 分钟'.format(int(looptime / 60)))
            time.sleep(looptime)

