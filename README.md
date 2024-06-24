一、说明：
    基于pyautogui开发，用于稳定性测试中机器崩溃后自动重新拉起游戏进程。

二、依赖资源文件说明：
    【config.json】：【Version】:表示每次需要启动的游戏版本（如Q5）。如后续版本不一，可进行手动变更（均大写）;【RT】表示可以接受的崩溃次数（以数字表示）;【WT】表示多久检测一次进程是否存在。
    【images】目录：存放pyautogui每次识别所需的参考图片。

pyinstaller --onefile .\auto_restart.py