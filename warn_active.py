# noinspection PyCompatibility
import configparser
import datetime
import platform

import os

from TimerTask import Task
from email_dict import to_send_email
from email_send import EmailSend
from loading import Loading
from mylogger import Logger
from warning_main import WarningPlay
from warnstone import stoneobject

if __name__ == '__main__':
    # 记录器 实例
    logname = "生日预警日志"
    log = Logger(logname)
    logger = log.getlogger()
    # 解析器实例
    conf = configparser.ConfigParser()
    path = 'warning.conf'
    assert os.path.exists(path), "{file}不存在".format(file=path)
    if platform.system() == 'Windows':
        conf.read(path, encoding="utf-8-sig")
    else:
        conf.read(path)
    # 数据库实例
    stone = stoneobject()
    # 初始化 定时器
    task = Task("08:00", logger)
    times = conf.get(section="time", option="now")
    if task.times != datetime.time(int(times.split(':')[0]), int(times.split(':')[1])):
        task.times = input("请输入开始时间，例如08:00")
    # task.times = "19:45"
    logger.debug("主程序开始运行")
    # 初始化 邮件发送、短信发送
    send_mail = EmailSend(smtp_server=conf.get(section='email', option='smtp_server'),
                          smtp_port=conf.get(section='email', option='smtp_port'),
                          from_addr=conf.get(section='email', option='from_addr'),
                          from_addr_str=conf.get(section='email', option='from_addr_str'),
                          password=conf.get(section='email', option='password'),
                          error_address=conf.get(section='email', option='error_email'), logger=logger, stone=stone,
                          special=conf.items(section='关注'))
    loading = Loading(stone=stone, logger=logger, conf=conf, )
    to_address_dict = to_send_email(file_name='主管及以上名单.xlsx')
    while True:
        # 到达预定时间后，执行 BlessingPlay.play
        #  收件人需要获取
        # 使用 to_address_dict 传递 收件人字典
        task.run(WarningPlay.play, to_address=to_address_dict, logger=logger, send_mail=send_mail, loading=loading,
                 date=datetime.date.today() + datetime.timedelta(days=4))
        logger.debug("当次执行完毕")
