# coding: utf-8

from easytrader import remoteclient

user = remoteclient.use('ths', host='192.168.80.141', port=1430)

# 通用同花顺版本需在远程服务器上手动登录
user.connect()

print(user.balance)