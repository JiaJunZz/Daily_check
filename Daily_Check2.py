#!/usr/bin/env python
# -*- coding: utf-8 -*-

import paramiko
import xlrd,xlwt
import time
import wmi
import shutil
import re
from xlutils3.copy import copy
from multiprocessing import Process



class Myprocess(Process):

    def __init__(self,hostname,username,password,txt_fname):
        # 定义构造函数
        super(Myprocess,self).__init__()
        #Process.__init__(self)
        self.hostname = hostname
        self.username = username
        self.password = password
        self.txt_fname = txt_fname


    def execute_cmd(self,hostname,username,password,txt_fname):
        #########################################
        #    远程到服务器上执行命令获取状态值存入文件
        ########################################

        # 针对不同主机IP获取不同的shell命令
        cmd_res = []
        cmd_shell = []
        with open('./cmd/' + hostname + '_cmd', 'r', encoding='utf-8') as f:
            for i in f.readlines():
                if not i.startswith('#'):
                    cmd_shell.append(i.strip('\n'))

        print(hostname)
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy)
        try:
            ssh.connect(hostname=hostname, username=username, password=password)
            for x in cmd_shell:
                stdin, stdout, stderr = ssh.exec_command(x)
                res=stdout.read().decode().strip('\n')
                cmd_res.append(res)
            ssh.close()

            with open(txt_fname,'a',encoding='utf-8') as file:
                 file.write(hostname + '  ' + str(cmd_res) +'\n')
        except Exception as e:

            with open(txt_fname,'a',encoding='utf-8') as file:
                 file.write(hostname + '  ' + str(e) +'\n')

    def run(self):
        self.execute_cmd(self.hostname,self.username,self.password,self.txt_fname)



def WriteEecel(txt_fname,xls_fname):
    #################################
    #    导入列表数据至csv文件
    #################################


    shutil.copy("template.xls",xls_fname)           ###copy模板文件，生成对应日期的excel文件

    rb = xlrd.open_workbook(xls_fname,formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)

    table = rb.sheet_by_name(u'sheet1')
    cols = table.col_values(1)  # 获取第1列内容


    with open(txt_fname,'r',encoding='utf-8') as f:
        for line in f.readlines():
            datas = line.split('  ')
            csv_ip = datas[0]
            if re.search('Error',datas[1]):
                continue
            else:
                csv_data = eval(datas[1])
                csv_body = [csv_data[0], csv_data[1], csv_data[2], csv_data[3]]
                print(csv_body)

                for ip_col in cols:
                    if csv_ip == ip_col:       #如果找到对应第一列的IP，即导入数据到对应IP的行
                        row = cols.index(ip_col)
                        i = 0
                        for col in range(3,7):
                            ws.write(row,col,csv_body[i])
                            i = i + 1

    wb.save(xls_fname)


def get_win():

    c = wmi.WMI()

    # cpu 使用率
    cpu_usage = str(c.Win32_Processor()[0].LoadPercentage) + '%'

    # 内存使用率
    TotalMemory = int(c.Win32_ComputerSystem()[0].TotalPhysicalMemory)/1024
    FreeMemory = int(c.Win32_OperatingSystem()[0].FreePhysicalMemory)
    memory_usage = str(int((TotalMemory - FreeMemory)/TotalMemory*100)) + '%'

    # 硬盘使用率
    TotalDisk = int(c.Win32_LogicalDisk(DeviceID = "C:")[0].Size)
    FreeDisk = int(c.Win32_LogicalDisk(DeviceID = "C:")[0].FreeSpace)
    Disk_usage = str(int((TotalDisk-FreeDisk)/TotalDisk*100)) + '%'

    # 监控数据列表
    win_list = [cpu_usage,'/',memory_usage,Disk_usage,'/']

    print(win_list)




if __name__ == '__main__':

    start = time.time()

    #  以时间戳定义txt日志以及xls日志文件名
    time_local = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    txt_fname = './cmd_res/' + time_local + '.txt'
    xls_fname = './logs/' + time_local + '.xls'

    #  多进程获取主机巡检值
    with open('hosts.info','r',encoding='utf-8') as f:
        processes = []
        for line in f.readlines():
            line_list = line.split()
            app = line_list[0]
            hostname = line_list[1]
            username = line_list[2]
            password = line_list[3]
            process = line_list[4]
            p = Myprocess(hostname,username,password,txt_fname)
            processes.append(p)
            p.start()

        for p in processes:
            # 等待进程退出
            p.join()

    WriteEecel(txt_fname,xls_fname)

    end = time.time()
    print('Task  runs %0.2f seconds.' % (end - start))
