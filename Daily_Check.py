#!/usr/bin/env python
# -*- coding: utf-8 -*-

import paramiko
import xlwt
import xlrd
import time
import wmi
import shutil
import re
from xlutils3.copy import copy



def execute_cmd(hostname,username,password):
    #########################################
    #    ssh远程到服务器上执行命令获取服务器状态值
    ########################################
    cmd_res = []
    cmd_shell = []
    with open('./cmd/' + hostname + '_cmd', 'r', encoding='utf-8') as f:
    # with open('./cmd/centos_cmd', 'r', encoding='utf-8') as f:
        for i in f.readlines():
            if not i.startswith('#'):
                cmd_shell.append(i.strip('\n'))
        #print(cmd_shell)

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy)
    try:
        ssh.connect(hostname=hostname, username=username, password=password)
        for x in cmd_shell:
            stdin, stdout, stderr = ssh.exec_command(x)
            res=stdout.read().decode().strip('\n')
            cmd_res.append(res)
        ssh.close()
        return cmd_res
    except Exception as e:
        return str(e)

def save_csv(txt_fname):
    #################################
    #    导入列表数据至csv文件
    #################################

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')
    ####title_style样式
    # title_style = xlwt.XFStyle()
    # ####body_style样式
    # #body_style = xlwt.XFStyle
    # ######字体
    # font = xlwt.Font()
    # font.name = "SimSun"  # 宋体
    # font.height = 20 * 11  # 字体大小为11，字体大小的基本单位是20.
    # font.bold = False  # 设置字体为不加粗
    # #font.colour_index = 0x01  # 字体颜色默认为黑色，此处设置字体颜色为白色， 颜色范围为：0x00-0xff
    # title_style.font = font
    # ######背景颜色
    # pat = xlwt.Pattern()
    # pat.pattern = xlwt.Pattern.SOLID_PATTERN  # 设置单元格背景颜色
    # pat.pattern_fore_colour = xlwt.Style.colour_map['gray25']  # 设置单元格背景颜色为深蓝
    # title_style.pattern = pat
    # ###边框
    # borders = xlwt.Borders()
    # borders.left = 1
    # borders.right = 1
    # borders.top = 1
    # borders.bottom = 1
    # title_style.borders = borders
    # title_style.alignment.horz = 2  # 水平居中 值为2
    # title_style.alignment.vert = 1  # 垂直居中 值为1


    # title = [u'应用说明',u'IP地址',u'进程名',u'CPU使用率',u'5分钟负载',u'内存使用率',u'/空闲率',u'/home空闲率',u'进程',u'备注']
    # for i in range(0,len(title)):
    #     worksheet.write(0,i,title[i],title_style)

    k = 1
    with open(txt_fname,'r',encoding='utf-8') as f:
        for line in f.readlines():
            datas = line.split('@@')
            csv_app = datas[0]
            csv_ip = datas[1]
            csv_data = eval(datas[3])
            csv_process = datas[2]
            #csv_body = [csv_app, csv_ip, csv_process, csv_data[0], csv_data[1], csv_data[2], csv_data[3], csv_data[4]]
            csv_body = [csv_data[0], csv_data[1], csv_data[2], csv_data[3], csv_data[4]]
            print(csv_body)

            for j in range(3,8):
                worksheet.write(k,j,csv_body[j])
            k = k + 1

    daily_xls = './logs/' + time.strftime("%Y%m%d%H_%M",time.localtime()) + '.xls'
    shutil.copy("template.xls",daily_xls)

    workbook.save(daily_xls)


def WriteEecel(txt_fname):
    #################################
    #    导入列表数据至csv文件
    #################################

    daily_xls = './logs/' + time.strftime("%Y%m%d_%H%M%S",time.localtime()) + '.xls'
    shutil.copy("template.xls",daily_xls)           ###copy模板文件，生成对应日期的excel文件

    rb = xlrd.open_workbook(daily_xls,formatting_info=True)
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

    wb.save(daily_xls)



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
    res_fname = './cmd_res/' + time.strftime("%Y%m%d_%H%M%S", time.localtime()) + '.txt'
    with open('hosts.info','r',encoding='utf-8') as f:
        for line in f.readlines():
            line_list = line.split()
            app = line_list[0]
            hostname = line_list[1]
            username = line_list[2]
            password = line_list[3]
            process = line_list[4]
            excel_res = execute_cmd(hostname,username,password)
            with open(res_fname,'a',encoding='utf-8') as file:
                file.write(hostname + '  ' + str(excel_res) +'\n')

    WriteEecel(txt_fname = res_fname)
    end = time.time()

    print("total time is %0.2f" % (end - start))
    # #get_win()



