import pyautogui
import xlrd
import xlwt
import pyperclip
import time

def paste(foo):
    pyperclip.copy(foo)
    pyautogui.hotkey('ctrl','v')


def over_all():
    # 打开文件
    workbook = xlrd.open_workbook(r'123.xlsx')

    # 获取所有sheet
    sheet1_name = workbook.sheet_names()[0]

    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
    sheet1 = workbook.sheet_by_name('Sheet1')

    # 获取整列的值（数组）
    cols2 = sheet1.col_values(1)  # 获取第2列内容（身份证号）
    cols1 = sheet1.col_values(0)  # 获取第1列内容（备注名）
    # 输出第一列第2列内容
    print(cols2)
    print(cols1)
    pyautogui.click(x=936,y=230)
    for i in range(0,p_m):
    # 为了避免出现未注册的情况
        pyautogui.click(x=331, y=435)
        time.sleep(0.3)
    # 双击定位文本框
        pyautogui.click(x=294,y=280,clicks=2,interval=0.25)
        time.sleep(0.3)
    # 输入身份证号
        sheetcols2='%s'%cols2[i]
        paste(sheetcols2)
        time.sleep(0.3)
    # 删除空格
        pyautogui.press('backspace')
        time.sleep(0.3)
    # 点击旁边空白处
        pyautogui.click(x=107, y=369)
        time.sleep(1)
    # 双击定位需要复制的信息
        pyautogui.click(x=266, y=360,clicks=2,interval=0.25)
        time.sleep(0.3)
    # 同时ctrl+c进行复制
        pyautogui.hotkey('ctrl','c')
        time.sleep(0.3)
    # 点击旁边空白处
        pyautogui.click(x=107, y=369)
        time.sleep(0.5)
    # 点击excel空白处
        pyautogui.click(x=1317, y=335)
        time.sleep(0.3)
    # ctrl+v黏贴
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
    # 下移一格
        pyautogui.hotkey('down')
        time.sleep(0.3)
        print('%s已完成'%cols1[i])



if __name__ == '__main__':
    p_m = int(input("输入你想整理的人数："))
    print('程序10s后即将开始，请将此窗口缩小')
    time.sleep(10)
    over_all()
    print('前50人已完成，请及时保存、删除')
