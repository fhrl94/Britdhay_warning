import datetime
import time


class WarningPlay(object):
    def __init__(self):
        pass

    #  收件人地址（字典） to_address
    @classmethod
    def play(cls, to_address, logger, send_mail, loading, date):
        """
        类方法，不需要实例化
        功能：
        1、判断今天是不是工作日
        2、判断今天之后 days 天为工作日
        3、实例化 class Loading 对象loading ，并进行数据处理 run
        4、今天是工作日则发送 1 + days 天的邮件，否则跳过
        5、发送当天的短信
        6、休眠23H
        :param date: 开始日期
        :param to_address:收件地址字典
        :param logger:记录器实例
        :param send_mail:邮件发送实例
        :param loading:数据处理实例
        :return:
        """
        #  需要执行的动作
        if (date + datetime.timedelta(days=1)).month - date.month != 0:
            i = 28
            while True:
                if (date + datetime.timedelta(days=i + 1)).month != (date + datetime.timedelta(days=1)).month:
                    break
                i += 1
            logger.debug("月预警开始执行")
            loading.run(today=date, after_day=1, number=i, )
            #  邮件发送
            send_mail.send(to_address=to_address,
                           header="{month}月-生日预警".format(month=(date + datetime.timedelta(days=1)).month), )
            logger.debug("查询日期天数:{day}".format(day=i))
            logger.debug("月预警发送完成")
        # 判断时间，如果时间为周五，调用unloading 存到 周转存表
        # 不需要每周一次
        if date.isoweekday() == 5:
            logger.debug("周预警开始执行")
            loading.run(today=date, after_day=3, number=7, )
            #  邮件发送
            send_mail.send(to_address=to_address, header="{days}-下周生日预警".format(days=date), )
            logger.debug("查询日期天数:{day}".format(day=7))
            logger.debug("周预警发送完成")
        logger.debug("执行完毕开始休眠23H")
        time.sleep(60 * 60 * 23)
