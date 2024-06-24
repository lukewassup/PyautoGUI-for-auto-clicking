import pyautogui
import subprocess
import time
import psutil
import requests
import win32com.client
import socket
import json
import configparser
from py3nvml.py3nvml import *

# 任务栏游戏图标检测
def game_icon_check():
    try:
        pyautogui.locateCenterOnScreen(game_icon, confidence = 0.8)
    except Exception:
        return False
    return True

# windows开始图标检测
def is_win_icon_exsist():
    try:
        pyautogui.locateCenterOnScreen(win_check, confidence = 0.8)
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

# 进入大厅单独处理，各个机型情况不一
def into_lobby():
    while True:
        try:
            loc = pyautogui.locateCenterOnScreen(start_lobby, confidence = 0.8)
            if loc is not None:
                time.sleep(2)
                pyautogui.click(loc, clicks = 1, interval = 0.5, duration = 1)
                break
        except Exception:
            print('大厅连接中，请稍后......')
            time.sleep(3)      

def click_to_start():
    while True:
        # 如有断线重连提示，点击确认重新加入
        try:
            # 设置建议进入大厅时间等待可以久一点
            time.sleep(10)
            loc1 = pyautogui.locateCenterOnScreen(reconnet, confidence = 0.8)
            if loc1 is not None:
                loc2 = pyautogui.locateCenterOnScreen(confirm_button, confidence = 0.6)
                pyautogui.click(loc2, clicks = 1, interval = 0.5, duration = 1)
                break
        except Exception:
            print('无需重连') 
        time.sleep(10)

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
    # pyautogui.press('`')
    # time.sleep(1)
    # pyautogui.write('gm startlevel withrobot')
    # time.sleep(1)
    # pyautogui.press('enter')
    print('进入对局中......')

def start_process(process_name):
    global pid
    global cnt
    
    proc = subprocess.Popen(process_name) 
    pid = proc.pid
    print(f'游戏开启，版本为【{this_time_version}服】-【{process_name}】-【{pid}】')
    time.sleep(10)
    cnt += 1

def notice_conf():
    headers = {"Content-Type": "application/json; charset = utf-8"}
    message = {
        "msgtype": "markdown",
        "markdown": {
            "content": f"<font color=\"info\">【以下为本次测试机器配置与内网IP对应关系】</font>\n"
                        f"> <font color=\"info\">CPU：{CPU}</font>\n"
                        f"> <font color=\"info\">显卡：{GPU}</font>\n"
                        f"> <font color=\"info\">内存：{RAM}GB</font>\n"
                        f"> <font color=\"info\">内网IP地址：{IP}</font>\n"
        }
    }
    requests.post(notice_url, headers = headers, json = message)

def notice_crash():
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

# 上传客户端崩溃日志
def upload_log():
    # 进入客户端日志目录将最近崩溃日志拷贝至公盘61.175的崩溃日志汇总内
    command = f'New-Item -path {pub_crash_dir}\{CPU}-{GPU}-{IP} -itemtype directory -force; cd {log_dir}; Get-ChildItem -Path .\*.log | Sort-Object LastWritetime -Descending | Select-Object -First {restart_time} | Copy-Item -Destination {pub_crash_dir}\{CPU}-{GPU}'
    subprocess.run(['powershell', '-Command', command], capture_output=True, text=True)

# 未知启动问题，暂时先通过拷贝指定dll文件到游戏目录下解决
def copy_dll():
    command = f'Copy-Item -Path .\\images\\msvcp140_2.dll, .\\images\\vcruntime140_1.dll -Destination {install_dir}\FromTheForgotten\Binaries\Win64 -Force'
    subprocess.run(['powershell', '-Command', command], capture_output=True, text=True)    

if __name__ == '__main__':
    wmi = win32com.client.GetObject("winmgmts:")
    notice_url= '********************************************************'
    crash_report_url = '********************************'
    dgs_dir = r'C:\Users\Public\Documents\DigiskyGameSync\game_install_info.json'
    pub_crash_dir = '\\192.168.61.175\wtshare\稳定性测试\客户端崩溃日志'
    cnt = 0

    # 根据显示屏分辨率选择不同的参考图片 
    screen = pyautogui.size()
    if screen == (1920, 1080):
        steps = ['images\\fight.png', 'images\\start_match.png']
        fight = 'images\\fight.png'
        game_icon = 'images\\game_icon.png'
        win_check = 'images\\win_check.png'
        start_lobby = 'images\\start_lobby.png'
        start_match = 'images\\start_match.png'
        reconnet = 'images\\reconnet.png'
        confirm_button = 'images\\confirm_button.png'
        lose_connect  = 'images\\lose_connect.png'
        lose_connect_confirm  = 'images\\lose_connect_confirm.png'
    else:
        steps = ['images\\fight_2k.png', 'images\\start_match_2k.png']
        fight = 'images\\fight_2k.png'
        game_icon = 'images\\game_icon_2k.png'
        win_check = 'images\\win_check_2k.png'
        start_lobby = 'images\\start_lobby_2k.png'
        start_match = 'images\\start_match_2k.png'
        reconnet = 'images\\reconnet_2k.png'
        confirm_button = 'images\\confirm_button_2k.png'  
        lose_connect  = 'images\\lose_connect_2k.png'
        lose_connect_confirm  = 'images\\lose_connect_confirm_2k.png' 
        
    # 收集硬件信息
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
    except Exception:
        print('硬件信息收集失败！')

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

    process_name = f'{install_dir}\FromTheForgottenClient.exe -unattended -RLPlayerAutoTest -forceinputaccount=false {params}'
    log_dir = f'{install_dir}\FromTheForgotten\Saved\Logs'
    # 手动修改配置使得每次打开游戏强制全屏
    ini = configparser.ConfigParser()
    ini.read(f'{install_dir}\\FromTheForgotten\\Saved\\Config\\WindowsClient\\GameUserSettings.ini')
    ini['/Script/DSMutilMonitor.DSGameUserSettings']['LastConfirmedFullscreenMode'] = '1'
    with open(f'{install_dir}\\FromTheForgotten\\Saved\\Config\\WindowsClient\\GameUserSettings.ini', 'w') as inifile:
        ini.write(inifile)

    copy_dll()
    # notice_conf()

    while True:
        # 首次开始游戏
        if cnt == 0:
            start_process(process_name)
            into_lobby()
            click_to_start()
            
        # 进程未运行且能检测到windows图标，判断为游戏进程结束
        if (not is_process_running(pid)) and is_win_icon_exsist():
            print('游戏意外退出，进程重启中......当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
            if cnt == restart_time:
                print(f'严重问题！取消重启！崩溃次数：【{cnt}次】 当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
                kill_process_by_pid(process_name)
                notice_crash()
                upload_log()
                break
            start_process(process_name)
            into_lobby()
            click_to_start()
        # 退到桌面可以检测到windows图标但进程仍然存在，仍判断为游戏进程结束
        elif is_process_running(pid) and is_win_icon_exsist() and not game_icon_check():
            # 等待尽可能长让进程自行销毁
            time.sleep(30)
            print('游戏意外退出，进程重启中......当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
            if cnt == restart_time:
                print(f'严重问题！取消重启！崩溃次数：【{cnt}次】 当前时间:', time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
                kill_process_by_pid(process_name)
                notice_crash()
                upload_log()
                break
            kill_process_by_pid(process_name)
            time.sleep(1) 
            start_process(process_name)
            into_lobby()
            click_to_start()

        # 检测游戏是否与服务器断开连接，如断开点击确定退回大厅并重新匹配
        try:
            loc3 = pyautogui.locateCenterOnScreen(lose_connect, confidence = 0.8)
            if loc3 is not None:
                loc4 = pyautogui.locateCenterOnScreen(lose_connect_confirm, confidence = 0.6)
                pyautogui.click(loc4, clicks = 1, interval = 0.5, duration = 1)
                print('与服务器失去连接，退回大厅......')
        except Exception:
            pass

        # 检测游戏对局是否结束，如结束重新点击进行匹配
        try:
            loc1 = pyautogui.locateCenterOnScreen(fight, confidence = 0.8)
            # 检测正常对局结束退到大厅后短暂等待10s
            time.sleep(10)
            pyautogui.click(loc1, clicks = 1, interval = 0.5, duration = 1)
            time.sleep(3)
            loc2 = pyautogui.locateCenterOnScreen(start_match, confidence = 0.8)
            pyautogui.click(loc2, clicks = 1, interval = 0.5, duration = 1)
            print('重新匹配......')
        except Exception:
            print('游戏对局进行中......')

        # 进程循环检测
        time.sleep(wait_time) 