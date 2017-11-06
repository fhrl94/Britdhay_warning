import glob
import os
import platform
import smtplib
import zipfile
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

import jinja2
from sqlalchemy import and_, distinct, text

from warningsend import Send
from warnstone import DaysMapping


class EmailSend(Send):
    def __init__(self, smtp_server, smtp_port, from_addr, from_addr_str, password, logger, stone, error_address,
                 special, *args, **kwargs):
        """
        初始化 【参数】，并获取模板
        :param smtp_server: SMTP地址
        :param smtp_port: SMTP端口
        :param from_addr: 发件人邮箱
        :param from_addr_str: 发件人友好名称
        :param password: 发件人邮件密码
        :param logger: 记录器实例
        :param stone: 数据库实例
        :param special: 关注人特例
        :param error_address 备份发送地址
        :param args:
        :param kwargs:
        """
        # noinspection PyCompatibility
        super().__init__(logger, stone, *args, **kwargs)
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.from_addr = from_addr
        self.from_addr_str = from_addr_str
        self.password = password
        self.error_address = error_address
        self.logger = logger
        self.stone = stone
        self.special = special
        self.brith_result = None
        self.templates = None
        self.temp_dir = os.path.abspath('.') + os.sep + 'temp' + os.sep
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
        self.zip_path = os.path.abspath('.') + os.sep + '发送情况' + r'.zip'
        self._get_template()

    #  实现邮件发送
    def send(self, to_address, header, ):
        """
        先调用 _get_data() 获取数据，在调用 _get_template() 获取模板，并将数据进行填充。再调用 _email_send() 发送邮件
        顺序不能调整
        :param to_address: 收件人地址（字典）
        :param header:邮件主题【标题】
        :return:
        """
        self.logger.debug("开始发送邮件")
        for emp in self._get_data():
            body = self.templates.render(brith_list=emp.get('employee_list'))
            if emp.get('superior', False):
                # 更新发送数据
                for one in emp.get('employee_list'):
                    one.count += 1
                try:
                    #  邮件发送，未测试
                    assert to_address.get(emp.get('superior'), False), "{name}无邮箱".format(name=emp.get('superior'))
                    self._email_send(to_address=to_address.get(emp.get('superior')), header=header, body=body)
                except AssertionError:
                    self.logger.warning("{name}无邮箱".format(name=emp.get('superior')))
            # 一次没有发送邮件，没有 superior属性
            self._save_html_file(name=emp.get('superior', header + '未发送'), str_html=body)
        self._save_zip()
        self._send_multimedia(to_address=self.error_address, body='邮件发送详情见附件', file=self.zip_path,
                              header=header + '发送情况')
        self.logger.debug("邮件发送完成")

        pass

    #  数据获取需要重新实现
    # 使用生成器实现数据返回
    def _get_data(self):
        """
        通过使用 【数据表】、【结果表】 2个字典，实现获取不同类型的数据，并进行存储
        :return:
        """
        # 生日名单 brith_result  司龄名单 siling_result
        self.logger.debug("开始获取邮件数据")
        cols = ['director', 'director1', 'manager', 'manager1', 'majordomo', 'principal']
        # 删除上一次生成的所有文件
        for dir_path, dir_names, file_names in os.walk(self.temp_dir):
            for file in file_names:
                os.remove(os.path.join(dir_path, file))
        # 总监级以下的人员内容生成（员工开始），部门第一负责人内容生成（主管开始）
        for col in cols:
            result = self.stone.query(distinct(getattr(DaysMapping, col))).filter(
                getattr(DaysMapping, col) != None).all()
            # print(result)
            #  result 是列表嵌套，每个元素都是list，第一个是col
            for one in result:
                # print(one)
                # print(stone.query(DaysMapping).filter(getattr(DaysMapping,col)==one[0]).all())
                if col == 'principal':
                    gather = self.stone.query(DaysMapping).filter(
                        and_(getattr(DaysMapping, col) == one[0], getattr(DaysMapping, 'job') != '员工')).all()
                else:
                    gather = self.stone.query(DaysMapping).filter(getattr(DaysMapping, col) == one[0]).all()
                #   返回管理人员数据，数据只能在返回后进行处理
                # email_draw(stone, gather, one[0])
                # gather.count += 1
                yield {'employee_list': gather, 'superior': one[0], }
        # 特殊人员配置
        assert len(self.special) == 2, '配置文件中人员指定格式错误'
        result = self.stone.query(DaysMapping).filter(
            text("name in ('{name}')".format(name="','".join(self.special[1][1].split(','))))).all()
        yield {'employee_list': result, 'superior': self.special[0][1], }
        self.stone.commit()
        # 将没有发送的人员进行返回，只有一次
        result = self.stone.query(DaysMapping).filter(DaysMapping.count == 0).all()
        yield {'employee_list': result, }
        self.logger.debug("邮件数据获取完成")
        # print(len(siling_result))
        pass

    #  实现模板获取
    def _get_template(self):
        # 使用 jinja2 生成邮件模板，并根据【数据获取】的结果进行填充
        # jinja2 需要定位到当前目录
        self.logger.debug("邮件模板获取开始")
        template_loader = jinja2.FileSystemLoader(searchpath=os.path.abspath(".") + os.sep + 'templates' + os.sep)
        template_env = jinja2.Environment(loader=template_loader)
        templates = template_env.get_or_select_template('warning.html')
        # return templates
        self.templates = templates
        self.logger.debug("邮件模板获取完成")

    def _email_send(self, to_address, header, body, ):
        """
        邮件投递
        :param to_address: 收件人地址，格式为字符串，以逗号隔开
        :param header: 主题内容
        :param body: 正文内容
        :return: 无返回值
        """

        def _format_addr(s):
            name, addr = parseaddr(s)
            return formataddr((Header(name, 'utf-8').encode(), addr))

        # 正文
        msg = MIMEText(body, 'html', 'utf-8')
        # 主题，
        msg['Subject'] = Header(header, 'utf-8').encode()
        # 发件人别名
        msg['From'] = _format_addr('{name}<{addr}>'.format(name=self.from_addr_str, addr=self.from_addr))
        # 收件人别名
        msg['To'] = to_address
        server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
        server.login(self.from_addr, self.password)
        server.sendmail(self.from_addr, to_address.split(','), msg.as_string())
        server.quit()
        # print('邮件投递成功')
        # self.logger.info("邮件投递成功")

    def _send_multimedia(self, to_address, header, body, file):
        """

        :param to_address: 收件人地址，格式为字符串，以逗号隔开
        :param header: 主题内容
        :param body: 正文内容
        :param file: 附件路径名称
        :return: 无返回值
        """

        def _format_addr(s):
            name, addr = parseaddr(s)
            return formataddr((Header(name, 'utf-8').encode(), addr))

        # 正文
        msg = MIMEMultipart()
        # 主题，
        msg['Subject'] = Header(header, 'utf-8').encode()
        # 发件人别名
        msg['From'] = _format_addr('{name}<{addr}>'.format(name=self.from_addr_str, addr=self.from_addr))
        # 收件人别名
        msg['To'] = to_address
        msg.attach(MIMEText(body, 'html', 'utf-8'))
        with open(file, 'rb') as f:
            # 设置附件的MIME和文件名，这里是png类型:
            mime = MIMEBase("application", "zip")
            # 加上必要的头信息:
            #  windows平台使用gbk ,无法全平台通用
            if platform.system() == 'Windows':
                mime.add_header('Content-Disposition', 'attachment', filename=('gbk', '', os.path.basename(file)))
            else:
                mime.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', os.path.basename(file)))
            # mime.add_header('Content-ID', '<0>')
            # mime.add_header('X-Attachment-Id', '0')
            # 把附件的内容读进来:
            mime.set_payload(f.read())
            # 用Base64编码:
            encoders.encode_base64(mime)
            # 添加到MIMEMultipart:
            msg.attach(mime)
        server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
        server.login(self.from_addr, self.password)
        server.sendmail(self.from_addr, to_address.split(','), msg.as_string())
        server.quit()
        # print('发送成功')
        self.logger.debug("邮件发送情况已备份")

    def _save_html_file(self, name, str_html):
        path = self.temp_dir + name + '.html'
        with open(path, 'a', encoding='utf-8') as file:
            file.write(str_html)

    def _save_zip(self):
        # 生成zip文件
        if os.path.exists(self.zip_path):
            os.remove(self.zip_path)
        f = zipfile.ZipFile(self.zip_path, 'w', zipfile.ZIP_DEFLATED)
        files = glob.glob(self.temp_dir + '*')
        for file in files:
            f.write(file, os.path.basename(file))
        f.close()

    pass
