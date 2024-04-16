#!/usr/bin/env python
# -*- coding: utf-8 -*-

import socket
import subprocess
import sys
from datetime import datetime

# 获取本地主机名
host = socket.gethostname()
# 设置目标IP段
ip_range = "172.16.0."

# 扫描函数
def scan_ip(ip):
    try:
        # 创建Socket连接
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # 设置超时时间为1秒
        sock.settimeout(1)
        # 连接到目标IP的80端口
        result = sock.connect_ex((ip, 80))
        if result == 0:
            print(f"Host {ip} is up")
        sock.close()
    except socket.error:
        pass

# 扫描IP段内的主机
def scan_ip_range():
    # 获取当前时间
    start_time = datetime.now()

    # 使用多线程扫描IP段内的主机
    for i in range(1, 255):
        ip = ip_range + str(i)
        scan_ip(ip)

    # 输出扫描所花费的时间
    end_time = datetime.now()
    total_time = end_time - start_time
    print(f"Scanning completed in {total_time}")

# 执行IP扫描
scan_ip_range()
