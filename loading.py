import datetime

import pymssql
from sqlalchemy import text, and_

from warnstone import Relation, DaysMapping, EmployeeInfo


class Loading(object):
    def __init__(self, stone, logger, conf):
        """
        初始化 数据库连接，并将上次运行的遗留数据进行删除
        :param stone:数据库连接
        :param logger:记录器实例
        :param conf: 配置文件实例
        """
        self.stone = stone
        self.logger = logger
        self.conf = conf
        self._data_delete()
        self._create_data()
        self.logger.debug("数据库 EmployeeInfo 初始化完成")

    #  待优化，若每月和每周发送是同一天会2次初始化，可能会影响性能
    # 每月和每周是一个逻辑，但是时间长度不一致，2次初始化；我认为是正常
    def run(self, today, after_day, number, ):
        """

        :param today:
        :param after_day:
        :param number:
        :return:
        """
        self._data_delete()
        self.logger.debug("已删除上次所有数据")
        self._create_data()
        self.logger.debug("重新拉取所有数据")
        self._transform(today=today, after_day=after_day, number=number, )
        self.logger.debug("已更新{number}天数据".format(number=number))
        pass

    def _create_data(self):
        """
        数据初始化
        1、从金蝶中获取职员表（职员信息（姓名、工号、出生日期）、职务、职位ID）
        2、职位表（上下级关系）
        将金蝶数据转存到sqlite 数据库中
        转存（在职人员） 职员表 姓名、工号、出生日期、职位唯一ID（涉及到sqlite支不支持UID），
        需要处理 职务的获取（存在兼职的情况）
        :return:
        """
        conn = pymssql.connect(self.conf.get('server', 'ip'), self.conf.get('server', 'user'),
                               self.conf.get('server', 'password'), database=self.conf.get('server', 'database'))
        cur = conn.cursor()
        sql = """
        select he.Name,he.Code,he.Birthday,ope.PositionID,oj.Name,ou.Name,ope.IsPrimary from HM_Employees as he 
        join ORG_Position_Employee as ope on he.EM_ID=ope.EmID 
        join ORG_Position as op on ope.PositionID=op.ID
        join ORG_Job as oj on op.JobID=oj.ID
        join ORG_Unit as ou on op.UnitID=ou.ID
        where he.Status =1 and op.IsDelete=0 and ou.StatusID=1 and  ou.Name not like '%远郊%'"""
        cur.execute(sql)
        emp_cols = ['name', 'code', 'birthDate', 'positionID', 'job', 'departmentname', 'IsPrimary']
        for one in cur.fetchall():
            empinfo = EmployeeInfo()
            for count, col in enumerate(emp_cols):
                if col == 'positionID':
                    setattr(empinfo, col, str(one[count]))
                else:
                    setattr(empinfo, col, one[count])
            self.stone.add(empinfo)
        self.stone.commit()
        #  转存（存在的岗位） 职位表 当前职位ID，父级职位ID
        sql = """
        select ID,ParentID from ORG_Position"""
        cur.execute(sql)
        re_cols = ['positionID', 'parentID', ]
        print('_______')
        for one in cur.fetchall():
            print(one)
            relation = Relation()
            for count, col in enumerate(re_cols):
                setattr(relation, col, str(one[count]))
            self.stone.add(relation)
        print('_______')
        self.stone.commit()
        cur.close()
        conn.close()
        pass

    def _transform(self, today, after_day, number, ):
        """
        数据转存
        1、将指定日期期间的人员放入 转存表 的表头，通过迭代获取相应的上级
        2、兼职人员处理
        :param today:  当前日期
        :param after_day:  预警日期与当前日期相差天数
        :param number:  预警的周期天数
        :return:
        """
        #  日期处理迭代（可用函数实现），将时间区域的人转存到 月转存表 或 周转存表
        today = today + datetime.timedelta(days=after_day)
        for num in range(number):
            result = self.stone.query(EmployeeInfo).filter(
                and_(text("strftime('%m%d',DATE (birthDate,'1 day'))=strftime('%m%d',date(:date,:value)) ")),
                EmployeeInfo.IsPrimary == True).params(value='{num} day'.format(num=num + 1), date=today).all()
            for one in result:
                # print(one)
                tab = DaysMapping()
                # 处理后缀数值
                try:
                    int(one.name[len(one.name) - 1])
                    tab.name = one.name[:len(one.name) - 1]
                except ValueError:
                    tab.name = one.name
                tab.code = one.code
                tab.birthDate = one.birthDate
                tab.departmentname = one.departmentname
                tab.positionID = one.positionID
                tab.job = one.job
                tab.date = today + datetime.timedelta(days=num)
                tab.count = 0
                tab.director = None
                tab.director1 = None
                tab.manager = None
                tab.manager1 = None
                tab.majordomo = None
                tab.principal = None
                tab.general_manager = None
                self.stone.add(tab)
        self.stone.commit()
        #  上级获取迭代（可用函数实现）
        result = self.stone.query(DaysMapping).all()
        for one in result:
            position = one.positionID
            director = []
            manager = []
            majordomo = []
            principal = []
            parent = self.stone.query(Relation).filter(Relation.positionID == position).one_or_none()
            while parent:
                position = parent.parentID
                parent = self.stone.query(EmployeeInfo).filter(EmployeeInfo.positionID == position).one_or_none()
                if parent is None:
                    # print('空')
                    # self.logger.debug("职位ID:{ID} 无上级".format(ID=position))
                    break
                if parent.job == '主管':
                    director.append(parent.name)
                elif parent.job == '经理':
                    manager.append(parent.name)
                elif parent.job == '总监':
                    majordomo.append(parent.name)
                elif parent.job == '副总':
                    principal.append(parent.name)
                parent = self.stone.query(Relation).filter(Relation.positionID == position).one_or_none()
            if director:
                one.director1 = director.pop()
            if director:
                one.director = director.pop()
            if manager:
                one.manager1 = manager.pop()
            if manager:
                one.manager = manager.pop()
            if majordomo:
                one.majordomo = majordomo.pop()
            if principal and principal[-1] != self.conf.get(section='special', option='name'):
                one.general_manager = principal.pop()
            if principal:
                one.principal = principal.pop()
        self.stone.commit()

    def _data_delete(self):
        """
        删除全部数据
        :return:
        """
        self.stone.query(EmployeeInfo).delete()
        self.stone.query(Relation).delete()
        self.stone.query(DaysMapping).delete()
        self.stone.commit()
        pass

    pass
