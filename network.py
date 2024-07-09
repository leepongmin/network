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
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from netmiko import ConnectHandler

INFO_PATH = os.path.join(os.getcwd(), 'leepongmin.xlsx')
LOCAL_TIME = time.strftime('%Y.%m.%d', time.localtime())
LOG_DIR = os.path.join(os.getcwd(), LOCAL_TIME)

if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

ERROR_LOG = os.path.join(LOG_DIR, 'err-log.log')

def log_error(message):
    with open(ERROR_LOG, 'a', encoding='utf-8') as log:
        log.write(message + '\n')

def get_devices_info(info_file):
    try:
        devices_dataframe = pd.read_excel(info_file, sheet_name=0, dtype=str, keep_default_na=False)
        return devices_dataframe.to_dict('records')
    except FileNotFoundError:
        print('\n没有找到表格文件！\n')
        input('Press Enter to exit.')
        sys.exit(1)

def get_cmds_info(info_file):
    try:
        cmds_dataframe = pd.read_excel(info_file, sheet_name=1, dtype=str)
        return cmds_dataframe.to_dict('list')
    except ValueError:
        print('\n表格文件缺失子表格信息！\n')
        input('Press Enter to exit.')
        sys.exit(1)

def handle_exception(device, error_type, error_message):
    print(f'设备 {device["host"]} {error_message}')
    log_error(f'设备 {device["host"]} {error_message}')

def inspection(device, cmds_dict):
    start_time = time.time()
    ssh = None

    try:
        ssh = ConnectHandler(**device)
        ssh.enable()
    except Exception as e:
        error_map = {
            'AttributeError': '缺少设备管理地址！',
            'NetmikoTimeoutException': '管理地址或端口不可达！',
            'NetmikoAuthenticationException': '用户名或密码认证失败！',
            'ValueError': 'Enable密码认证失败！',
            'TimeoutError': 'Telnet连接超时！',
            'ReadTimeout': 'Enable密码认证失败！',
            'ConnectionRefusedError': '远程登录协议错误！',
        }
        handle_exception(device, type(e).__name__, error_map.get(type(e).__name__, f'未知错误！{type(e).__name__}'))
    else:
        log_file_path = os.path.join(LOG_DIR, f'{device["host"]}.log')
        with open(log_file_path, 'w', encoding='utf-8') as device_log_file:
            print(f'设备 {device["host"]} 正在采集...')
            for cmd in cmds_dict[device['device_type']]:
                device_log_file.write(f'{"=" * 10} {cmd} {"=" * 10}\n\n')
                show = ssh.send_command(cmd, read_timeout=30)
                device_log_file.write(show + '\n\n')
        elapsed_time = time.time() - start_time
        print(f'设备 {device["host"]} 采集完成，用时 {round(elapsed_time, 1)} 秒。')
    finally:
        if ssh:
            ssh.disconnect()

if __name__ == '__main__':
    start_time = time.time()
    devices_info = get_devices_info(INFO_PATH)
    cmds_info = get_cmds_info(INFO_PATH)

    print('\n采集开始...')
    print('\n' + '-' * 40 + '\n')

    with ThreadPoolExecutor(max_workers=200) as executor:
        futures = [executor.submit(inspection, device, cmds_info) for device in devices_info]
        for future in futures:
            future.result()

    try:
        with open(ERROR_LOG, 'r', encoding='utf-8') as log_file:
            error_count = len(log_file.readlines())
    except FileNotFoundError:
        error_count = 0

    total_time = time.time() - start_time
    print('\n' + '-' * 40 + '\n')
    print(f'采集完成，共采集 {len(devices_info)} 台设备，{error_count} 台异常，共用时 {round(total_time, 1)} 秒。\n')
    input('Press Enter to exit.')
