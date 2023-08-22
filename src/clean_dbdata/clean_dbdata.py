'''
Author: liusuxian 382185882@qq.com
Date: 2023-08-18 15:21:21
LastEditors: liusuxian 382185882@qq.com
LastEditTime: 2023-08-22 10:40:59
Description: 

Copyright (c) 2023 by liusuxian email: 382185882@qq.com, All Rights Reserved.
'''
import pymysql

# 打开数据库连接
db = pymysql.connect(
    host='localhost',
    user='aiweiju',
    password='aiweiju!@#$%',
    database='aiweiju_test'
)
print("数据库连接成功！！！")
# 查询非活跃用户SQL语句
query_users_sql = "SELECT * FROM common_user WHERE lasttime <= '2023-06-01 00:00:00' AND riches_aidou = 0 AND (vip IS NULL OR vip < NOW())"
# 非活跃用户插入历史表SQL语句
insert_users_sql = "INSERT INTO common_user_history (id, nickname, openid, avatar, did, mpid, riches_aidou, vip, remark, channel, fromuid, device, creator, lasttime, createtime, created_at, updated_at) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"


def query_users(sql: str):
    try:
        # 使用 with 打开游标，自动关闭
        with db.cursor() as cursor:
            # 执行 SQL 语句
            cursor.execute(sql)
            # 获取所有记录列表
            return cursor.fetchall()
    except pymysql.Error as e:
        print("Error: unable to fetch data:", e)
        return []


# 查询非活跃用户
users = query_users(query_users_sql)
# 分批大小
batch_size = 1000
# 分批插入数据
count = 0
# 插入数据总量
insert_total = 0
# 删除数据总量
delete_total = 0
while count < len(users):
    batch_data = users[count:count+batch_size]
    try:
        # 使用 with 打开游标，自动关闭
        with db.cursor() as cursor:
            # 执行批量插入操作
            cursor.executemany(insert_users_sql, batch_data)
            db.commit()
            insert_total += len(batch_data)
            print("本次成功插入了", len(batch_data), "条记录")
            # 批量删除操作
            user_ids = [r[0] for r in batch_data]
            try:
                # 执行批量删除操作
                delete_user_ids = ",".join([str(uid) for uid in user_ids])
                delete_users_sql = f"DELETE FROM common_user WHERE id IN ({delete_user_ids})"
                cursor.execute(delete_users_sql)
                db.commit()
                delete_total += len(user_ids)
                print("本次成功删除了", len(user_ids), "个用户")
            except pymysql.Error as e:
                print("Error deleting data:", e)
                db.rollback()
    except pymysql.Error as e:
        print("Error inserting data:", e)
        db.rollback()
    count += batch_size
print("成功插入了", insert_total, "条记录")
print("成功删除了", delete_total, "个用户")
# 关闭数据库连接
db.close()
print("数据库关闭成功！！！")
