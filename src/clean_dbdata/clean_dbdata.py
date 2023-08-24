'''
Author: liusuxian 382185882@qq.com
Date: 2023-08-18 15:21:21
LastEditors: liusuxian 382185882@qq.com
LastEditTime: 2023-08-24 20:38:07
Description: 

Copyright (c) 2023 by liusuxian email: 382185882@qq.com, All Rights Reserved.
'''
import json
import time
import pymysql
import redis

# 打开数据库连接
db = pymysql.connect(
    host='localhost',
    user='aiweiju',
    password='aiweiju!@#$%',
    database='aiweiju_test'
)
print("数据库连接成功！！！")
# 打开 redis 连接
redis_pool_0 = redis.ConnectionPool(
    host='localhost',
    port=6379,
    db=0,
    password='',
    decode_responses=True,
    max_connections=10
)
redis_pool_1 = redis.ConnectionPool(
    host='localhost',
    port=6379,
    db=1,
    password='',
    decode_responses=True,
    max_connections=10
)
redis_client_0 = redis.Redis(connection_pool=redis_pool_0)
pipeline_0 = redis_client_0.pipeline()
redis_client_1 = redis.Redis(connection_pool=redis_pool_1)
pipeline_1 = redis_client_1.pipeline()
print("redis 连接成功！！！")
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


def get_keys_to_delete(redisClient, pattern):
    start_time = time.time()  # 记录开始时间
    cursor = 0
    keys_to_delete = []
    while True:
        cursor, keys = redisClient.scan(cursor, match=pattern, count=500)
        keys_to_delete.extend(keys)
        if cursor == 0:
            break
    end_time = time.time()  # 记录结束时间
    elapsed_time = end_time - start_time  # 计算耗时
    return keys_to_delete, elapsed_time


def delete_redis_keys(redisPipeline, keysList):
    for keys in keysList:
        if len(keys) > 0:
            redisPipeline.delete(*keys)
    redisPipeline.execute()


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
        # 批量删除 Redis 数据
        delete_keys_list_0 = []
        delete_keys_list_1 = []
        # 微信openid、微信Unionid、微信用户登录过的appid列表key、微信用户登录过的openid列表key、微信小程序会话密钥key
        wxOpenidKeys = []
        wxUnionidKeys = []
        wxLoginAppidKeys = []
        wxLoginOpenidKeys = []
        wxSessionKeys = []
        # 抖音DeviceId、抖音匿名openid、抖音openid、抖音Unionid、抖音用户登录过的appid列表key、抖音用户登录过的openid列表key、抖音小程序会话密钥key
        dyDeviceidKeys = []
        dyOpenidKeys = []
        dyUnionidKeys = []
        dyLoginAppidKeys = []
        dyLoginOpenidKeys = []
        dySessionKeys = []
        # 快手openid、快手用户登录过的appid列表key、快手用户登录过的openid列表key、快手小程序会话密钥key
        ksOpenidKeys = []
        ksLoginAppidKeys = []
        ksLoginOpenidKeys = []
        ksSessionKeys = []
        # 给第三方平台开放的注册用户账号
        openUuidKeys = []

        userKeys = ['user:' + str(uid) for uid in user_ids]
        userList = redis_client_1.mget(*userKeys)
        for user in userList:
            user_data = json.loads(user)
            uid = str(user_data["id"])
            appid = str(user_data["appid"])
            openid = str(user_data["openid"])
            unionid = str(user_data["unionid"])
            deviceid = str(user_data["device"])
            if appid[:2] == 'wx':
                # 微信小程序、微信H5
                wxOpenidKeys.append('openid:' + openid)
                if len(unionid) > 0:
                    wxUnionidKeys.append('wx:unionid:' + unionid)
                wxLoginAppidKeys.append('wx:appid:list:' + uid)
                wxLoginOpenidKeys.append('wx:openid:list:' + uid)
                wxSessionKeys.append('wx:sessionKey:' + uid)
            elif appid[:2] == 'tt':
                # 抖音小程序
                dyOpenidKeys.append('dy:openid:' + openid)
                if len(unionid) > 0:
                    dyUnionidKeys.append('dy:unionid:' + unionid)
                dyLoginAppidKeys.append('dy:appid:list:' + uid)
                dyLoginOpenidKeys.append('dy:openid:list:' + uid)
                dySessionKeys.append('dy:sessionKey:' + uid)
            elif appid[:2] == 'ks':
                # 快手小程序
                ksOpenidKeys.append('ks:openid:' + openid)
                ksLoginAppidKeys.append('ks:appid:list:' + uid)
                ksLoginOpenidKeys.append('ks:openid:list:' + uid)
                ksSessionKeys.append('ks:sessionKey:' + uid)
            elif appid[:3] == 'awj':
                # 互动广告H5
                openUuidKeys.append('open:uuid:' + openid)
            elif appid == 'awqqy314q01jrapu':
                # 抖音H5
                dyDeviceidKeys.append('dy:deviceid:' + deviceid)
                dyLoginAppidKeys.append('dy:appid:list:' + uid)
        # 用户数据
        # 渠道信息
        channelInfoKeys = ['channel:info:' + str(uid) for uid in user_ids]
        # 用户追剧列表
        collectListKeys = ['collect:list:' + str(uid) for uid in user_ids]
        # 用户观看剧列表
        watchListKeys = ['watch:list:' + str(uid) for uid in user_ids]
        # 用户剧集对应视频解锁列表、用户观看剧信息、IOS用户累计观看时长、IOS用户充值支付权限
        unlockList = []
        watchInfoList = []
        iosWatchTimeList = []
        iosPayAuthList = []
        for k in watchListKeys:
            pipeline_1.zrange(k, 0, -1)
        pipelineResults = pipeline_1.execute()  # 执行 Pipeline 中的命令
        for uid, sidList in zip(user_ids, pipelineResults):
            unlockKeys = [
                'sp:vd:unlock:' + str(uid) + ':' + str(sid)
                for sid in sidList
            ]
            unlockList.extend(unlockKeys)
            watchInfoKeys = [
                'watch:info:' + str(uid) + ':' + str(sid)
                for sid in sidList
            ]
            watchInfoList.extend(watchInfoKeys)
            iosWatchTimeKeys = [
                'ios:watchtime:' + str(uid) + ':' + str(sid)
                for sid in sidList
            ]
            iosWatchTimeList.extend(iosWatchTimeKeys)
            iosPayAuthKeys = [
                'ios:payauth:' + str(uid) + ':' + str(sid)
                for sid in sidList
            ]
            iosPayAuthList.extend(iosPayAuthKeys)
        # 用户签到信息
        qiandaoKeys = ['qiandao:' + str(uid) for uid in user_ids]
        # 客户端useragent信息
        launchKeys = ['launch:' + str(uid) for uid in user_ids]
        # 小程序广告投放信息
        advertisingKeys = ['advertising:' + str(uid) for uid in user_ids]

        delete_keys_list_0.append(advertisingKeys)

        delete_keys_list_1.append(wxOpenidKeys)
        delete_keys_list_1.append(wxUnionidKeys)
        delete_keys_list_1.append(wxLoginAppidKeys)
        delete_keys_list_1.append(wxLoginOpenidKeys)
        delete_keys_list_1.append(wxSessionKeys)

        delete_keys_list_1.append(dyDeviceidKeys)
        delete_keys_list_1.append(dyOpenidKeys)
        delete_keys_list_1.append(dyUnionidKeys)
        delete_keys_list_1.append(dyLoginAppidKeys)
        delete_keys_list_1.append(dyLoginOpenidKeys)
        delete_keys_list_1.append(dySessionKeys)

        delete_keys_list_1.append(ksOpenidKeys)
        delete_keys_list_1.append(ksLoginAppidKeys)
        delete_keys_list_1.append(ksLoginOpenidKeys)
        delete_keys_list_1.append(ksSessionKeys)

        delete_keys_list_1.append(openUuidKeys)

        delete_keys_list_1.append(userKeys)
        delete_keys_list_1.append(channelInfoKeys)
        delete_keys_list_1.append(collectListKeys)
        delete_keys_list_1.append(watchListKeys)
        delete_keys_list_1.append(unlockList)
        delete_keys_list_1.append(watchInfoList)
        delete_keys_list_1.append(iosWatchTimeList)
        delete_keys_list_1.append(iosPayAuthList)
        delete_keys_list_1.append(qiandaoKeys)
        delete_keys_list_1.append(launchKeys)
        delete_redis_keys(pipeline_0, delete_keys_list_0)
        delete_redis_keys(pipeline_1, delete_keys_list_1)
    except pymysql.Error as e:
        print("Error inserting data:", e)
        db.rollback()
    count += batch_size
print("成功插入了", insert_total, "条记录")
print("成功删除了", delete_total, "个用户")
# 关闭数据库连接
db.close()
print("数据库关闭成功！！！")
# 关闭 redis 连接
redis_client_0.close()
redis_pool_0.disconnect()
redis_client_1.close()
redis_pool_1.disconnect()
print("redis 关闭成功！！！")
