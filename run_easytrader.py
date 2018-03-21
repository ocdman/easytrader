# coding: utf-8

import easytrader
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from pywinauto import Application, win32defines
from pywinauto.win32functions import SetForegroundWindow, ShowWindow

user = easytrader.use('ths')
user.connect()

'''
为了触发点击按钮事件
需将交易系统窗口置于最前
'''
def set_foreground_window():
    w = user._app.top_window()
    #bring window into foreground
    if w.HasStyle(win32defines.WS_MINIMIZE): # if minimized
        ShowWindow(w.wrapper_object(), 9) # restore window state
    else:
        SetForegroundWindow(w.wrapper_object()) #bring to front

'''
获取站点的连接状态
True  表示连接成功
False 表示连接失败
'''
def get_connection_status():
    statusBar = user._main.window(
        control_id=user._config.STATUS_BAR_CONTROL_ID,
        class_name='msctls_statusbar32'
    )
    #status = statusBar.GetProperties()['texts'][2]
    status = statusBar.GetPartText(1)
    return status != '断开'

'''
最多尝试连接站点3次
'''
def try_connection(times):
    print('第' + str(4 - times) + '次连接站点')
    if times == 0:
        return
    if get_connection_status():
        return
    else:
        #set_foreground_window()
        toolbar = user._main.window(
            control_id=user._config.TOOL_BAR_CONTROL_ID
        )
        #toolbar.PressButton(3) #点击刷新按钮
        toolbar.Button(3).Click()
        user._wait(5)
    times -= 1
    try_connection(times)

'''
定时任务
'''
def job():
    print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    try_connection(3)
    if get_connection_status():
        print('已连接站点')
        for position in user.position:
            profit = position['参考盈亏']
            ratio = position['参考盈亏比(%)']
            marketPrice = position['市价']
            security = position['证券代码']
            name = position['证券名称']
            print('当前持股：' + name + '，证券代码：' + security + '，市价：' + str(marketPrice) + '，参考盈亏：' + str(profit) + '，参考盈亏比(%)：' + str(ratio))
            if profit > 500 or ratio > 25 :
                # user._switch_left_menus(['卖出[F2]'])
                # user._set_trade_params(security, price = marketPrice + 1, amount = 100)
                # user._submit_trade()
                user.sell(security, marketPrice + 1, 100)
                print('委托卖出股票：' + name + '，卖价：' + str(marketPrice + 1) + '，数量：100')
    else:
        print('连接站点失败，请检查网络')

'''
程序入口
每周一到周五，9:00-12:00，13:00-15:00，每分钟执行一次
'''
scheduler = BlockingScheduler()
scheduler.add_job(job, 'cron', day_of_week='0-4', hour='9-11,13-14', minute='*/1')
scheduler.start()

