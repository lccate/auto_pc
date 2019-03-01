# python控制电脑鼠标键盘自动运行
相关资料查阅：  
pywin32模拟鼠标查阅：https://www.cnblogs.com/huidaoli/p/7398392.html  
pywin32模拟键盘同时按下几个按键：https://blog.csdn.net/u014595019/article/details/49131321  
pywin32.keybd_event模拟键盘输入：https://blog.csdn.net/polyhedronx/article/details/81988948  
使用pyUserInput：https://blog.csdn.net/u013783095/article/details/79630358  
pyautogui自动输入：https://blog.csdn.net/qq_40379759/article/details/80427235  
pyautogui：https://www.jianshu.com/p/ff337d381a64  
pyautogui：https://hugit.app/posts/doc-pyautogui.html  

出现的问题：  
分别尝试了3种方法pywin32，pyuserinput，pyautogui  
1.使用pywin32时，运行程序改变了原有的鼠标和键盘，得重新开机关机才能恢复正常  
2.安装pyUserInput时一直出现以下错误提示，手动安装也未能解决问题  
```
Could not find a version that satisfies the requirement pyHook (from PyUserInput) (from versions: )
No matching distribution found for pyHook (from PyUserInput)
```
3.安装pyautogui时，出现错误提示
```
Command "python setup.py egg_info" failed with error code 1 in C:\Users\lcc\AppData\Local\Temp\pip-install-5sm1527n\pygetwindow\
```
查找了大量的
