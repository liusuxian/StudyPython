import datetime


# 获取上周周一和周日的日期
def getLastWeek():
    """
    获取上周周一和周日的日期
    :return: 上周周一和周日的日期
    """
    # 获取当前日期, 因为要求时分秒为0, 所以不要求时间
    today = datetime.date.today()
    # 获取当前周的排序, 周一为0, 周日为6
    weekday = today.weekday()
    # 当前日期距离上周一的时间差
    monday_delta = datetime.timedelta(weekday + 7)
    # 获取上周一日期
    monday = today - monday_delta
    # 当前日期距离上周日的时间差
    sunday_delta = datetime.timedelta(weekday + 1)
    # 获取上周日日期
    sunday = today - sunday_delta
    return monday, sunday


# 获取本周周一和周日的日期
def getThisWeek():
    """
    获取本周周一和周日的日期
    :return: 本周周一和周日的日期
    """
    # 获取当前日期, 因为要求时分秒为0, 所以不要求时间
    today = datetime.date.today()
    # 获取当前周的排序, 周一为0, 周日为6
    weekday = today.weekday()
    # 当前日期距离这周一的时间差
    monday_delta = datetime.timedelta(weekday)
    # 获取这周一日期
    monday = today - monday_delta
    # 当前日期距离这周日的时间差
    sunday_delta = datetime.timedelta(6 - weekday)
    # 获取这周日日期
    sunday = today + sunday_delta
    return monday, sunday


# 获取下周周一和周日的日期
def getNextWeek():
    """
    获取下周周一和周日的日期
    :return: 下周周一和周日的日期
    """
    # 获取当前日期, 因为要求时分秒为0, 所以不要求时间
    today = datetime.date.today()
    # 获取当前周的排序, 周一为0, 周日为6
    weekday = today.weekday()
    # 当前日期距离下周一的时间差
    monday_delta = datetime.timedelta(7 - weekday)
    # 获取下周一日期
    monday = today + monday_delta
    # 当前日期距离下周日的时间差
    sunday_delta = datetime.timedelta(7 - weekday + 6)
    # 获取下周日日期
    sunday = today + sunday_delta
    return monday, sunday


# 删除日期中的年份并组装月份和天数
def delDateYear(date: datetime.date, connectStr: str = '-', suffixStr: str = ''):
    """
    删除日期中的年份并组装月份和天数
    :param date: 日期
    :param connectStr: 连接字符
    :param suffixStr: 后缀字符
    :return: 月份connectStr天数suffixStr
    """
    dateList = str(date).split('-')
    dateList.pop(0)
    return str(connectStr).join(dateList) + suffixStr


# 格式化日期
def formatDate(date: datetime.date):
    """
    格式化日期
    :param date: 日期
    :return: xxxx年xx月xx日
    """
    dateList = str(date).split('-')
    return dateList[0] + '年' + dateList[1] + '月' + dateList[2] + '日'
