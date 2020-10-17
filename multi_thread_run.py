import os
import threading
import time

from Conf.ProjVar import proj_path
from TestScript.hybrid import execute_testcase_by_excel
from selenium import webdriver
import queue


def count_success_num():
    global success_num
    lock.acquire()
    success_num += 1
    lock.release()


def count_fail_num():
    global fail_num
    lock.acquire()
    fail_num += 1
    lock.release()


def task(q):
    try:
        data_excel_file_path = q.get()
    except:
        return
    execute_testcase_by_excel(data_excel_file_path)


if __name__ == '__main__':
    file_path_list = []
    data_dir = os.path.join(proj_path,r"TestData")
    for root,dirs,files in os.walk(data_dir):
        for file in files:
            file_path_list.append(os.path.join(root,file))
    print(file_path_list)

    lock = threading.Lock()
    success_num = 0
    fail_num = 0
    thread_list = []
    test_data = file_path_list

    q = queue.Queue()
    for data in test_data:
        q.put(data)

    for i in range(len(file_path_list)):
        t = threading.Thread(target=task, args=(q,))
        t.start()
        thread_list.append(t)

    for t in thread_list:
        t.join()

    print("执行完毕")