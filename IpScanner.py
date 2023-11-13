import socket
import subprocess
import sys
from datetime import datetime

# ��ȡ����������
host = socket.gethostname()
# ����Ŀ��IP��
ip_range = "172.16.0."

# ɨ�躯��
def scan_ip(ip):
    try:
        # ����Socket����
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # ���ó�ʱʱ��Ϊ1��
        sock.settimeout(1)
        # ���ӵ�Ŀ��IP��80�˿�
        result = sock.connect_ex((ip, 80))
        if result == 0:
            print(f"Host {ip} is up")
        sock.close()
    except socket.error:
        pass

# ɨ��IP���ڵ�����
def scan_ip_range():
    # ��ȡ��ǰʱ��
    start_time = datetime.now()

    # ʹ�ö��߳�ɨ��IP���ڵ�����
    for i in range(1, 255):
        ip = ip_range + str(i)
        scan_ip(ip)

    # ���ɨ�������ѵ�ʱ��
    end_time = datetime.now()
    total_time = end_time - start_time
    print(f"Scanning completed in {total_time}")

# ִ��IPɨ��
scan_ip_range()
