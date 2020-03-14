# coding:utf-8
# Author:Zhoubin
# 2020-02-04
from email.mime.text import MIMEText
import smtplib
import os
import requests
import arrow
from lxml import etree
from tabulate import tabulate
import csv
# from apscheduler.schedulers.background import BackgroundScheduler   # 后台定时任务，多线程式，非阻塞
from apscheduler.schedulers.blocking import BlockingScheduler   # 定时任务，阻塞式
from apscheduler.triggers.cron import CronTrigger               # 定时任务

class WtExcel():
    def __init__(self):
        pass                     
    # 实时抓取IOT机器运行信息
    def get_produce_information(self,text):
        now = arrow.now().format('YYYY-MM-DD HH:mm:ss')
        for t in text:
            t.insert(0,now)
        with open('d:\\IOT_test.csv','a',newline='') as f:
            strings= csv.writer(f,dialect='excel')
            strings.writerows(text)

# 发送电邮
def Send_mail(text):
    content=MIMEText(text,'html','utf-8')
    reveivers = 'ok0755@126.com,zhoubin@spintecgear.com'
    content['To']=reveivers
    content['From']= str('ok0755@126.com')
    content['Subject']='IOT实时运行信息'
    smtp_server = smtplib.SMTP('smtp.126.com',25)
    smtp_server.login('ok0755@126.com','syoVIS78')
    smtp_server.sendmail('ok0755@126.com',['ok0755@126.com','zhoubin@spintecgear.com'],content.as_string())
    smtp_server.quit()

# IOT信息
def get_produce_information():
        machine_lists = ["C-WY-01","C-WY-02","C-WY-03","C-WY-04","C-WY-05","C-WY-06","C-WY-07","C-WY-08","C-WY-09",\
            "C-WY-10","C-WY-11","C-WY-12","C-WY-13","C-WY-14","C-WY-15","C-WY-16","C-WY-17","C-WY-18","C-WY-19",\
                "C-WY-20","C-WY-21","C-WY-22","C-WY-23","C-WY-24","C-WY-25","C-WY-26"]

        url = 'http://192.168.9.164:88/realTimeMonitor/update'   # IOT data Web entrance
        #url = 'http://59.40.183.220:8888/realTimeMonitor/update' # 外网IP
        header = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36'}        
        text_rows = []
        response = requests.get(url,headers=header)                                         
        data =response.json()
        response.close
        x = 0
        IOT_informations = data.get('serverinformations')
        for d in IOT_informations:
            if (d['machineId'] in machine_lists) and ('3263' in d['productName']):
                h = d['totalRunTime'].split(':')
                sec = int(h[0])*3600 + int(h[1])*60 + int(h[2])
                avilabily_percent = sec/43200
                try:
                    OEE = d['producedParts']/(d['producedParts'] + d['ngParts']) * (64 * (d['producedParts'] + d['ngParts']) / sec) * avilabily_percent
                    OEE = '{:.2f}%'.format(OEE*100)
                except:
                    OEE = "0"
                                    # 机台、图号、生产数量、不良品数量、运行时间、停机时间、OEE
                text_rows.append([d['machineId'],d['productName'],d['producedParts'],d['ngParts'],d['totalRunTime'][:5],d['totalDownTime'][:5],OEE])
                x = x + int(d.get('producedParts'))
        text_rows.append(['','','','','','',''])  
        text_rows.append(['合计:','06.3263.2301.1',x,'','',''])         
        tabulate_header = ['机台','图号','生产数量','不良品数量','运行时间','停机时间','OEE']
        txt = tabulate(text_rows,headers=tabulate_header,tablefmt='html',stralign='left',missingval='')
        Send_mail(txt)
        WtExcel().get_produce_information(text_rows[:-2])

if __name__=='__main__':
    # mytask = BackgroundScheduler()
    # get_produce_information()
    mytask = BlockingScheduler()
    exe = CronTrigger(day='1-31',hour='7,19',minute='59')
    mytask.add_job(get_produce_information,exe)
    mytask.start()
    
