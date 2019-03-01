# python控制电脑鼠标键盘自动运行
### 相关资料查阅：  
pywin32模拟鼠标查阅：https://www.cnblogs.com/huidaoli/p/7398392.html  
pywin32模拟键盘同时按下几个按键：https://blog.csdn.net/u014595019/article/details/49131321  
pywin32.keybd_event模拟键盘输入：https://blog.csdn.net/polyhedronx/article/details/81988948  
使用pyUserInput：https://blog.csdn.net/u013783095/article/details/79630358  
pyautogui自动输入：https://blog.csdn.net/qq_40379759/article/details/80427235  
pyautogui：https://www.jianshu.com/p/ff337d381a64  
pyautogui：https://hugit.app/posts/doc-pyautogui.html  

### 出现的问题：  
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
查找了大量的资料，终于找到解决方法，这个错误是安装pyautogui的核心错误，解决方法就是降低PyGetWindow的版本，运行以下命令： 
```
pip install PyGetWindow==0.01
```
安装完PyGetWindow后出现以下提示，将下列包一个个pip即可  
```
pyautogui 0.9.41 requires pymsgbox, which is not installed.
pyautogui 0.9.41 requires pyscreeze, which is not installed.
pyautogui 0.9.41 requires PyTweening>=1.0.1, which is not installed.
```
最终pyautogui安装成功，import后无错误提示  
4.直接使用pyautogui.typewrite函数会导致只能输入一个‘ ’

### 程序设计
实现功能：将excel表格中身份证号信息输入系统文本框，截图出现的信息并保存  
* 1.提取表格前100人的身份证号放入数组  
* 2.遍历数组  
* 3.对于每个人都执行以下操作：  
1定位系统文本框坐标（定位坐标的方法很简单，打开qq的截图功能，从左上角0，0位置开始截图，显示数字多少就说明右下角的坐标是多少）  
2输入身份证号字符串（注意是字符串不是整数，因为有的身份证号后面带有X，从excel直接提取的字符串后面带有一个空格，直接使用pyautogui.typewrite函      数会导致只能输入一个‘ ’，解决方法见‘出现的问题4’）  
3定位系统
