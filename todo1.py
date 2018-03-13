#!/usr/bin/env python
# -*- coding: utf-8 -*-

from multiprocessing import Process  
import time  
import random  
  
def test():  
    for i in range(random.randint(1,5)):  
        print("----sub processing%d___"%i)  
        time.sleep(1)  
  
if __name__=='__main__':  
    p=Process(target=test)  
    p.start()  
    p.join()#这句话保证子进程结束后再向下执行  
    #p.join(2)#等待2s  
    #p.terminate() #进行结束  
    print("----after subprocesing 哈哈ding-----")  
  
#整个子进程结束后主进程才结束，p.join保证p进程结束后，才继续向下执行  
''''' 
----子进程中0___ 
----子进程中1___ 
----等待子进程结束后进行----- 
'''  