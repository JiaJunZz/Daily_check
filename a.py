#!/usr/bin/env python
# -*- coding: utf-8 -*-

from multiprocessing import Pool
from multiprocessing import Process


import time
import random


def test2():
    for i in range(random.randint(1, 5)):
        print("----子进程中%d___" % i)
        time.sleep(1)


if __name__ == '__main__':
    p = Process(target=test2)
    p.start()
    p.join()  # 这句话保证子进程结束后再向下执行
    # p.join(2)#等待2s
    # p.terminate() #进行结束
    print("----等待子进程结束后进行-----")

# 整个子进程结束后主进程才结束，p.join保证p进程结束后，才继续向下执行
''''' 
----子进程中0___ 
----子进程中1___ 
----等待子进程结束后进行----- 
'''