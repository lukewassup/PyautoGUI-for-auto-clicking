import pyautogui
import subprocess
import time
import psutil
import requests
import win32com.client
import socket
import json
from py3nvml.py3nvml import *
 
# 任务栏游戏图标检测
def game_icon_check():
    try:
        pyautogui.locateCenterOnScreen('images\game_icon.png', confidence = 0.8)
    except Exception:
        return False
    return True
 
# windows开始图标检测
def is_win_icon_exsist():
    try:
        pyautogui.locateCenterOnScreen('images\win_check.png', confidence = 0.8)
    except Exception:
        return False
    return True
 
# 以每次游戏进程启动的PID作为检测进程是否存在的条件
def is_process_running(pid):
    return psutil.pid_exists(pid)
 
# 通过pid来kill后台进程
def kill_process_by_pid(process_name):
    for proc in psutil.process_iter(['pid']):
        if proc.info['pid'] == pid:
            try:
                proc.kill()
                time.sleep(1)
                print(f'此进程{process_name}-{proc.info["pid"]}已关闭。')
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                print(f'此进程{process_name}-{proc.info["pid"]}不存在。')
 
# 全屏使用windows开始图标和游戏任务栏图标检查（进入大厅也放在这一步）
def fullscreen_check():
    try:
        loc1 = pyautogui.locateCenterOnScreen('images\win_check.png')
        loc2 = pyautogui.locateCenterOnScreen('images\game_icon.png')
        if (loc1 is not None) and (loc2 is not None):
            print('游戏不为全屏，切换为全屏......')
            pyautogui.hotkey('alt', 'enter')  
        time.sleep(3)
    except Exception:
        print('游戏为全屏')
        time.sleep(3)      
 
    while True:
        try:
            loc = pyautogui.locateCenterOnScreen('images\start_lobby.png', confidence = 0.8) 
            pyautogui.click(loc, clicks = 1, interval = 0.5, duration = 1)
            if loc is not None:
                break
        except Exception:
            print('大厅连接中，请稍后......')
            time.sleep(5)
 
def click_to_start():
    while True:
        # 如有断线重连提示，点击确认重新加入
        try:
            # 设置建议进入大厅时间等待可以久一点
            time.sleep(10)
            loc1 = pyautogui.locateCenterOnScreen('images\\reconnet.png', confidence = 0.8)
            if loc1 is not None:
                loc2 = pyautogui.locateCenterOnScreen('images\confirm_button.png', confidence = 0.6)
                pyautogui.click(loc2, clicks = 1, interval = 0.5, duration = 1)
                break
        except Exception:
            print('无需重连') 
 
        try:
            for step in steps:
                loc3 = pyautogui.locateCenterOnScreen(step, confidence = 0.8) 
                print(f'点击【{step}】')
                if loc3 is not None:
                    time.sleep(1)
                    pyautogui.click(loc3, clicks = 1, interval = 0.5, duration = 1)
        except Exception:
            print('未找到按钮，请检查主屏幕显示！')
        break
    print('进入对局中......')
 
def start_process(process_name):
    global pid
    global cnt
    
    proc = subprocess.Popen(process_name, cwd = 'C:\\Windows\\System32') 
    pid = proc.pid
    print(f'游戏开启，版本为【{this_time_version}服】-【{process_name}】-【{pid}】')
    time.sleep(10)
    cnt += 1
 
def notice():
    global CPU
    global GPU
    global IP
 
    try:
        # CPU信息
        for cpu in wmi.InstancesOf("Win32_Processor"):  
            CPU = cpu.Name
 
        # 显卡信息
        for gpu in wmi.InstancesOf("Win32_VideoController"):
            GPU = gpu.Name
 
        # 内存信息
        for mem in wmi.InstancesOf("Win32_PhysicalMemory"):
            RAM = f'{int(mem.Capacity) / 1024**3:.0f}'
 
        # 内网IP地址
        IP = socket.gethostbyname(socket.gethostname())
 
        headers = {"Content-Type": "application/json; charset = utf-8"}
        message = {
            "msgtype": "markdown",
            "markdown": {
                "content": f"<font color=\"warning\">【稳定性测试中此机型游戏进程因崩溃或其他原因退出多次，请相关同学注意！】</font>\n"
                            f"> <font color=\"info\">CPU：{CPU}</font>\n"
                            f"> <font color=\"info\">显卡：{GPU}</font>\n"
                            f"> <font color=\"info\">内存：{RAM}GB</font>\n"
                            f"> <font color=\"info\">内网IP地址：{IP}</font>\n"
                            f"> <font color=\"info\">版本：{this_time_version}</font>\n"
                            f"> <font color=\"info\">Crash-Report：[点击查看]({crash_report_url})</font>\n"
            }
        }
        requests.post(notice_url, headers = headers, json = message)
    except Exception:
        print('硬件信息收集失败！')
 
# 上传客户端崩溃日志
def upload_log():
    # 进入客户端日志目录将最近崩溃日志拷贝至公盘61.175的崩溃日志汇总内
    command = f'New-Item -path {pub_crash_dir}\{CPU}-{GPU}-{IP} -itemtype directory -force; cd {log_dir}; Get-ChildItem -Path .\*.log | Sort-Object LastWritetime -Descending | Select-Object -First {restart_time} | Copy-Item -Destination {pub_crash_dir}\{CPU}-{GPU}'
    subprocess.run(['powershell', '-Command', command], capture_output=True, text=True)
 
# 未知启动问题，暂时先通过拷贝指定dll文件到游戏目录下解决
def copy_dll():
    command = f'Copy-Item -Path .\\images\\msvcp140_2.dll, .\\images\\vcruntime140_1.dll -Destination {install_dir}\\ -Force'
    subprocess.run(['powershell', '-Command', command], capture_output=True, text=True)    
 
if __name__ == '__main__':
    wmi = win32com.client.GetObject("winmgmts:")
    notice_url= '******************************'
    crash_report_url = '***********************'
    dgs_dir = r'*************\*********\******.json'
    steps = ['images\\fight.png', 'images\\start_match.png']
    pub_crash_dir = '*************'
    cnt = 0
 
    # 初始化参数
    with open('config.json', 'r', encoding = 'utf-8') as t:
        this_time = json.load(t)    
        this_time_version = this_time.get('Version')
        restart_time = int(this_time.get('RT'))
        wait_time = int(this_time.get('WT'))
        params = this_time.get('Params')
 
    # 读取对应DGS目录json文件以获取游戏安装目录并启动
    with open(dgs_dir, 'r', encoding = 'utf-8') as f:
        configs = json.load(f)
    for config in configs:
        if config['GameVersionName'] == this_time_version:
            install_dir = config['InstallDirectory']
            break
 
    process_name = f'{install_dir}\*****.exe'
    log_dir = f'{install_dir}'
 
    copy_dll()
 
    while True:
        # 首次开始游戏
        if cnt == 0:
            start_process(process_name)
            fullscreen_check()
            click_to_start()
 
        # 进程未运行且能检测到windows图标，判断为游戏进程结束
        if (not is_process_running(pid)) and is_win_icon_exsist():
            print('游戏意外退出，进程重启中......当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
            if cnt == restart_time:
                print(f'严重问题！取消重启！崩溃次数：【{cnt}次】 当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
                kill_process_by_pid(process_name)
                notice()
                upload_log()
                break
            start_process(process_name)
            fullscreen_check()
            click_to_start()
        # 退到桌面可以检测到windows图标但进程仍然存在，仍判断为游戏进程结束
        elif is_process_running(pid) and is_win_icon_exsist() and not game_icon_check():
            # 等待尽可能长让进程自行销毁
            time.sleep(30)
            print('游戏意外退出，进程重启中......当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
            if cnt == restart_time:
                print(f'严重问题！取消重启！崩溃次数：【{cnt}次】 当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
                kill_process_by_pid(process_name)
                notice()
                upload_log()
                break
            kill_process_by_pid(process_name)
            time.sleep(1) 
            start_process(process_name)
            fullscreen_check()
            click_to_start()
 
        # 检测游戏对局是否结束，如结束重新点击进行匹配
        try:
            loc1 = pyautogui.locateCenterOnScreen('images\\fight.png', confidence = 0.8)
            # 检测正常对局结束退到大厅后短暂等待10s
            time.sleep(10)
            pyautogui.click(loc1, clicks = 1, interval = 0.5, duration = 1)
            time.sleep(3)
            loc2 = pyautogui.locateCenterOnScreen('images\start_match.png', confidence = 0.8)
            pyautogui.click(loc2, clicks = 1, interval = 0.5, duration = 1)
            print('重新开始匹配......')
        except Exception:
            print('游戏对局进行中......')
 
        # 进程循环检测
        time.sleep(wait_time) 