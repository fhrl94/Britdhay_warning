# 生日预警思路

根据金蝶上下级关系，通过金蝶信息中的生日信息，实现发送预警邮件

通过使用 ```Python warn_active.py``` 设定定时为【4:00】，可以运行时自行更改。（依赖库我不太记得了，如果运行失败，
自行百度或提问）

【warn_active】 实现资源的声明、每天定时运行的时间（使用 TimerTask 实现）、日志输出（mylogger）

【warning_main】 实现了业务流程

email_send/EmailSend 继承了 warningsend/Send 【模板获取 _get_template】
【数据获取 _get_data】 【邮件发送 send】 并实现了这三个函数的具体细节

email_dict/to_send_email() 方法实现获取 excel中的高管收件人

【warnstone】 实现了sqlite中的表设计，共三张表【员工信息】、【上下级关系】、【转存表】（存储生日人员） 

# 配置文件
```
[server]
ip = 
user = 
password = 
database = 

[email]
smtp_server = 
smtp_port = 
from_addr = 
from_addr_str = 生日预警管理站
password = 
error_email = 

[time]
now = 4:00

[special]
name = 
#（特殊情况 无上级但属于部门第一负责人【详见loading】，文件中请删除此行）

[关注]
关注人 = 
被关注人员名单 = 
#（这里可以设置特殊关注人员，文件中请删除此行）
```
根据【email_dict】设置excel表 

