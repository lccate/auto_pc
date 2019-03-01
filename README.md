# python控制电脑鼠标键盘自动运行
拟实现在系统中输入身份证号查询信息并截图的功能，程序比较简单，难点是用对工具并安装一些工具包，多次试验后最终选择用pyautogui模块  
## 相关资料查阅：  
pywin32模拟鼠标查阅：https://www.cnblogs.com/huidaoli/p/7398392.html  
pywin32模拟键盘同时按下几个按键：https://blog.csdn.net/u014595019/article/details/49131321  
pywin32.keybd_event模拟键盘输入：https://blog.csdn.net/polyhedronx/article/details/81988948  
使用pyUserInput：https://blog.csdn.net/u013783095/article/details/79630358  
pyautogui自动输入：https://blog.csdn.net/qq_40379759/article/details/80427235  
pyautogui：https://www.jianshu.com/p/ff337d381a64  
pyautogui：https://hugit.app/posts/doc-pyautogui.html  

## 出现的问题：  
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
使用pyperclip包，直接pip安装即可，利用copy函数，将字符串黏贴到文本框中  
```
def paste(foo):
    pyperclip.copy(foo)
    pyautogui.hotkey('ctrl','v')
```
5.如果搜索的信息不存在，没有出现详细信息窗口，而接下来关闭的动作依然会执行，会出现把主窗口关闭的情况，增加一个鼠标拖拽窗口往下拉动的动作，这样
## 程序设计
实现功能：将excel表格中身份证号信息输入系统文本框，截图出现的信息并保存  
### 系统1：
#### 1.提取表格前10人的身份证号放入数组    
#### 2.遍历数组x信息并输出，对于每个人都执行以下操作：  
1 定位系统文本框坐标（724，226）（定位坐标的方法很简单，打开qq的截图功能，从左上角0，0位置开始截图，显示数字多少就说明右下角的坐标是多少）  
2 删除原有的身份证号信息 （773，228）
3 输入身份证号字符串（注意是字符串不是整数，因为有的身份证号后面带有X，从excel直接提取的字符串后面带有一个空格，直接使用pyautogui.typewrite函      数会导致只能输入一个‘ ’，解决方法见‘出现的问题4’）  
4 输入文本框中的身份证号后面带有一个空格要进行删除  
```
pyautogui.press(‘backspace’)
```
5 定位系统“查询”按钮（751，346）  
6 等待5s系统出现查询信息  
7 定位系统“详细信息”按钮 (579,415)  
8 等待5s
9 截图需要的信息，左上角的xy坐标，宽度和高度（260，387，360，407），命名方式为“序号姓名”,保存到文件夹中  
10 关闭“详细信息窗口”(1275,14)  
11 输出“姓名已完成”
#### 3.遍历完成后输出“程序结束，前100人已统计完毕”  

### 系统2：
系统与表格窗口并排放置，快捷键tab+alt
提前将窗口调整为640X某的大小，主要是窗口长度，宽度不重要
#### 1.提取表格前n人的身份证号放入数组    
#### 2.现将光标定位到excel的第一个表格上    
#### 3.遍历数组x信息并输出，对于每个人都执行以下操作：
1 双击系统文本框坐标（294，280）  
```
pyautogui.click(clicks=2)
```
2 输入身份证号字符串  
4 身份证号后面空格删除    
```
pyautogui.press(‘backspace’)
```
5 在旁边空白位置点一下等待2秒出现结果（107，369）    
6 双击定位需要复制的信息(266,360)   
7 同时按下ctrl+c复制   
```
pyautogui.hotkey('ctrl', 'c')
```
8 在excel表格靠右的空白位置点击一次（1317，335）  
9 同时按下ctrl+v黏贴  
10 控制键盘向下箭头，下移一格down
11 输出“姓名已完成”
#### 3.遍历完成后输出“程序结束，前100人已统计完毕”  
