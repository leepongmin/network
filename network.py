#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
#
# Copyright (C) 2024 leepongmin, Inc. All Rights Reserved 
#
# @Time    : ${DATE} ${TIME}
# @Author  : ####leepongmin####
# @Email   : ####leepongmin@hotmail.com####
# @File    : ${NAME}.py
# @Software: ${PRODUCT_NAME}

import os
import sys
import time
import pandas
import threading
from netmiko import ConnectHandler

INFO_PATH = os.path.join(os.getcwd(), 'Sino-Bridge.xlsx') 
LOCAL_TIME = time.strftime('%Y.%m.%d', time.localtime())
LOCK = threading.Lock() 
POOL = threading.BoundedSemaphore(100) 

def get_devices_info(info_file): 
    try:
        devices_dataframe = pandas.read_excel(info_file, sheet_name=0, dtype=str, keep_default_na=False)

    except FileNotFoundError: 
        print(f'\n没有找到表格文件！\n') 
        input('Press Enter to exit.')
        sys.exit(1) 
    else:
        devices_dict = devices_dataframe.to_dict('records') 
        return devices_dict


def get_cmds_info(info_file):
    try:
        cmds_dataframe = pandas.read_excel(info_file, sheet_name=1, dtype=str)

    except ValueError: 
        print(f'\n表格文件缺失子表格信息！\n')  
        input('Press Enter to exit.') 
        sys.exit(1) 
    else:
        cmds_dict = cmds_dataframe.to_dict('list') 
        return cmds_dict


def inspection(login_info, cmds_dict):
    t11 = time.time() 
    ssh = None 

    try:
        ssh = ConnectHandler(**login_info) 
        ssh.enable() 
    except Exception as ssh_error:  
        with LOCK:  
            match type(ssh_error).__name__:  
                case 'AttributeError':  
                    print(f'设备 {login_info["host"]} 缺少设备管理地址！')  
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 缺少设备管理地址！\n')
                case 'NetmikoTimeoutException':
                    print(f'设备 {login_info["host"]} 管理地址或端口不可达！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 管理地址或端口不可达！\n')
                case 'NetmikoAuthenticationException':
                    print(f'设备 {login_info["host"]} 用户名或密码认证失败！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 用户名或密码认证失败！\n')
                case 'ValueError':
                    print(f'设备 {login_info["host"]} Enable密码认证失败！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} Enable密码认证失败！\n')
                case 'TimeoutError':
                    print(f'设备 {login_info["host"]} Telnet连接超时！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} Telnet连接超时！\n')
                case 'ReadTimeout':
                    print(f'设备 {login_info["host"]} Enable密码认证失败！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} Enable密码认证失败！\n')
                case 'ConnectionRefusedError':
                    print(f'设备 {login_info["host"]} 远程登录协议错误！')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 远程登录协议错误！\n')
                case _:
                    print(f'设备 {login_info["host"]} 未知错误！{type(ssh_error).__name__}')
                    with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'a', encoding='utf-8') as log:
                        log.write(f'设备 {login_info["host"]} 未知错误！{type(ssh_error).__name__}\n')
    else:  
        with open(os.path.join(os.getcwd(), LOCAL_TIME, login_info['host'] + '.log'), 'w', encoding='utf-8') as device_log_file:

            with LOCK:  
                print(f'设备 {login_info["host"]} 正在采集...')  
            for cmd in cmds_dict[login_info['device_type']]:  
                if type(cmd) is str: 
                    device_log_file.write('=' * 10 + ' ' + cmd + ' ' + '=' * 10 + '\n\n')  
                    show = ssh.send_command(cmd, read_timeout=30) 
                    device_log_file.write(show + '\n\n') 
        t12 = time.time()  
        with LOCK:  
            print(f'设备 {login_info["host"]} 采集完成，用时 {round(t12 - t11, 1)} 秒。')  
    finally: 
        if ssh is not None:  
            ssh.disconnect()  
        POOL.release()  


if __name__ == '__main__':
    t1 = time.time()  
    threading_list = []  
    devices_info = get_devices_info(INFO_PATH)  
    cmds_info = get_cmds_info(INFO_PATH)  

    print(f'\n北京神州新桥科技有限公司西安分公司') 
    print(f'\n采集开始...')  
    print(f'\n' + '-' * 40 + '\n')  

    if not os.path.exists(LOCAL_TIME):  
        os.makedirs(LOCAL_TIME)  
    else:  
        try:  
            os.remove(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'))  
        except FileNotFoundError:  
            pass 

    for device_info in devices_info: 
        pre_device = threading.Thread(target=inspection, args=(device_info, cmds_info))

        threading_list.append(pre_device)  
        POOL.acquire() 
        pre_device.start() 

    for _ in threading_list:  
        _.join()  

    try: 
        with open(os.path.join(os.getcwd(), LOCAL_TIME, 'err-log.log'), 'r', encoding='utf-8') as log_file:
            file_lines = len(log_file.readlines()) 
    except FileNotFoundError:  
        file_lines = 0  
    t2 = time.time()  
    print(f'\n' + '-' * 40 + '\n')  
    print(f'采集完成，共采集 {len(threading_list)} 台设备，{file_lines} 台异常，共用时 {round(t2 - t1, 1)} 秒。\n')  
    input('Press Enter to exit.')  
