#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: zhouw
# @Date:   2016-02-17 11:41:11
# @Last Modified by:   zhouw
# @Last Modified time: 2016-02-17 18:16:25

import os
import sys
import time
import socket
import urllib2
import ConfigParser
import smtplib
import poplib
import threading
import subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from email.utils import parsedate_tz
from email.utils import mktime_tz
from PIL import ImageGrab
from Tkinter import *

def get_ip_info():
    url = "http://1212.ip138.com/ic.asp"
    try:
        page = urllib2.urlopen(url)
    except:
        pass
    else:
        text = page.read()
        if "<center>" in text and "</center>" in text:
            text = text[text.index("<center>") + 8:text.index("</center>")]
            info_list = text.split(" ")
            ip = info_list[0]
            ip = ip[ip.index("[")+1:]
            ip = ip[:ip.index("]")]
            location = info_list[1]
            location = location[6:]
            location = location.decode("gb2312").encode("utf8")

            return (ip,location)

def get_info(msg, indent = 0):
    subject = ""
    addr = ""
    content = ""
    datetime = 0
    if indent == 0:
        for header in ['From', 'Subject', 'Date']:
            value = msg.get(header, '')
            if value:
                if header=='Subject':
                    subject = decode_str(value)

                if header=='From':
                    hdr, addr = parseaddr(value)
                if header=='Date':
                    datetime = mktime_tz(parsedate_tz(value))
    if msg.is_multipart():
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            content_value = get_info(part, indent + 1)[2]
            if content_value != "":
                content = content_value
    else:
        content_type = msg.get_content_type()
        if content_type=='text/plain':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset:
                content = content.decode(charset)
        else:
            content = ""
    return (subject, addr, content, datetime)

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def showMessage(msg):
    #show reminder message window
    root = Tk()  #建立根窗口
    #root.minsize(500, 200)   #定义窗口的大小
    #root.maxsize(1000, 400)  #不过定义窗口这个功能我没有使用
    root.withdraw()  #hide window
    #获取屏幕的宽度和高度，并且在高度上考虑到底部的任务栏，为了是弹出的窗口在屏幕中间
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight() - 100
    root.resizable(False,False)
    #添加组件
    root.title("Warning!!")
    frame = Frame(root, relief=RIDGE, borderwidth=3)
    frame.pack(fill=BOTH, expand=1) #pack() 放置组件若没有则组件不会显示
    #窗口显示的文字、并设置字体、字号
    label = Label(frame, text=msg, font="Monotype\ Corsiva -20 bold")
    label.pack(fill=BOTH, expand=1)
    #按钮的设置
    button = Button(frame, text="OK", font="Cooper -25 bold", fg="red", command=root.destroy)
    button.pack(side=BOTTOM)

    root.update_idletasks()
    root.deiconify() #now the window size was calculated
    root.withdraw() #hide the window again 防止窗口出现被拖动的感觉 具体原理未知？
    root.geometry('%sx%s+%s+%s' % (root.winfo_width() + 10, root.winfo_height() + 10,
        (screenwidth - root.winfo_width())/2, (screenheight - root.winfo_height())/2))
    root.deiconify()
    root.mainloop()

def send(touser, title, msg = None, file_name = None):
    if smtpserver == '' or smtpport == '':
        raise Exception("Error config smtp")

    #加邮件头
    msgRoot = MIMEMultipart('related')
    msgRoot['From'] = "Email My PC"
    msgRoot['Subject'] = title

    #构建消息文本
    if msg != None:
        msgText = MIMEText('%s'%msg,'html','utf-8')
        msgRoot.attach(msgText)

    #构造附件
    if file_name != None:
        att = MIMEText(open('%s'%file_name, 'rb').read(), 'base64', 'utf-8')
        att["Content-Type"] = 'application/octet-stream'
        att["Content-Disposition"] = 'attachment; filename="%s"'%file_name
        msgRoot.attach(att)

    #发送邮件
    try_num = 0
    try_max = 10
    while 1:
        try_num += 1
        try:
            smtp.sendmail(user, touser, msgRoot.as_string())
            break
        except:
            try:
                if smtpstype == 'ssl':
                    smtp = smtplib.SMTP_SSL(smtpserver, smtpport)
                else:
                    smtp = smtplib.SMTP(smtpserver, smtpport)

                    if smtpstype == 'tls':
                        smtp.ehlo()
                        smtp.starttls()
                        smtp.ehlo()

                # 设置调试模式，可以看到与服务器的交互信息
                smtp.set_debuglevel(1)

                smtp.login(user, passwd)
            except:
                if try_num > try_max:
                    break
                pass

    #删除附件
    if file_name != None:
        path = os.getcwd() + "\\" + file_name
        if os.path.exists(path):
            os.remove(path)

    #收拾数据
    # smtp.close()

def main():
    global config, popserver, popport, popstype, imapserver, imapport, imapstype, smtpserver, smtpport, smtpstype,\
    user, passwd, startsend, sleep, whitelist, tag_shutdown, tag_screen, tag_say, tag_cmd
    threads = []   #多线程

    # 载入配置
    config = ConfigParser.ConfigParser()
    config.read("config.ini")
    popserver = config.get("mail", "popserver")
    popport = config.get("mail", "popport")
    popstype = config.get("mail", "popstype")
    imapserver = config.get("mail", "imapserver")
    imapport = config.get("mail", "imapport")
    imapstype = config.get("mail", "imapstype")
    smtpserver = config.get("mail", "smtpserver")
    smtpport = config.get("mail", "smtpport")
    smtpstype = config.get("mail", "smtpstype")
    user = config.get("mail", "user")
    passwd = config.get("mail", "passwd")
    startsend = config.get("settings", "startsend")
    sleep = config.get("settings", "sleep")
    whitelist = config.get("settings", "whitelist")
    tag_shutdown = config.get("commands", "tag_shutdown")
    tag_screen = config.get("commands", "tag_screen")
    tag_say = config.get("commands", "tag_say")
    tag_cmd = config.get("commands", "tag_cmd")

    # 开启默认任务
    if startsend == '1':
        pc_name = socket.gethostname()
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        info = get_ip_info();

        print pc_name,current_time,info
        if info:
            title = "您的电脑当前有开机动作！"
            msg = "电脑名称：%s\n开机时间：%s\nIP地址：%s\n地理位置：%s" % (pc_name, current_time, info[0], info[1])

            mail_list = whitelist.split('\n')
            for mail in mail_list:
                print mail,title,msg

                #多线程调用的方法
                a = threading.Thread(target=send, args=(mail, title, msg))
                a.start()

    # return
    #开始循环读取邮件，处理收到的命令
    last_mail_time = time.time()
    run_num = 0
    while 1:
        run_num += 1
        try:
            resp, mails, octets = pop.list()
            print resp,mails,octets

            number = len(mails)
            for num in range(number, number - 5, -1):
                resp, lines, octets = pop.retr(num)
                # print resp, lines, octets

                msg_content = b'\r\n'.join(lines)
                msg = Parser().parsestr(msg_content)
                info = get_info(msg)
                subject = info[0].strip()
                addr = info[1].strip()
                content = info[2].strip()
                datetime = info[3]

                print datetime,time.asctime(time.localtime(datetime)),addr,subject

                # 如果是之前的邮件，那么跳过
                if datetime < last_mail_time:
                    print 'continue',time.asctime(time.localtime(datetime)),time.asctime(time.localtime(last_mail_time))
                    continue

                mail_list = whitelist.split('\n')
                if addr in mail_list:
                    title = None
                    msg = None
                    pic_name = None

                    # 初始化命令列表
                    parm = ""
                    tmp = subject.split(" ", 1)
                    cmd = tmp[0].lower()

                    if len(tmp) == 2:
                        parm = tmp[1]

                    if tag_shutdown.lower() == cmd:
                        subprocess.Popen('shutdown -s', shell=True)
                        title = "已成功执行关机命令！"

                    if tag_screen.lower() == cmd:
                        pic_name = time.strftime('%Y%m%d%H%M%S',time.localtime(time.time()))
                        pic_name = pic_name + ".jpg"
                        pic = ImageGrab.grab()
                        pic.save('%s' % pic_name)
                        title = "截屏成功！"

                    if tag_say.lower() == cmd:
                        showMessage(parm)
                        title = "已转达你所说的话！"

                    if tag_cmd.lower() == cmd:
                        subprocess.Popen(parm, shell=True)
                        title = "成功执行CMD命令！"

                    # 发送邮件
                    if title != None:
                        #发邮件
                        a = threading.Thread(target=send, args=(addr, title, msg, pic_name))
                        a.start()

                        #更新时间
                        last_mail_time = time.time()

            #断开连接，重新读取数据
            pop.close()
        except:
            try:
                if popstype == 'ssl':
                    pop = poplib.POP3_SSL(popserver, popport)
                else:
                    pop = poplib.POP3(popserver, popport)

                # 设置调试模式，可以看到与服务器的交互信息
                pop.set_debuglevel(1)

                pop.user(user)
                pop.pass_(passwd)
            except:
                pass

        print 'run num:',run_num
        time.sleep(float(sleep))
        pass

if __name__ == '__main__':
    main()
