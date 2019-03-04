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
    for i in range(0,2):
    #定位文本框
        pyautogui.click(x=724,y=226)
        time.sleep(1)
    #删除原来的身份证号
        pyautogui.click(x=773, y=228)
        time.sleep(1)
    #输入身份证号
        sheetcols2='%s'%cols2[i]
        paste(sheetcols2)
        time.sleep(1)
    #删除空格
        pyautogui.press('backspace')
        time.sleep(1)
    #点击查询
        pyautogui.click(x=751, y=346)
        time.sleep(12)
    #详细信息
        pyautogui.click(x=579, y=415)
        time.sleep(12)
    #截图并保存
        pics = pyautogui.screenshot("%d_%s.png"%(i,cols1[i]),region=(260,387,105,26))
    #关闭小窗口
        pyautogui.click(x=1275, y=14)
        print('%s已完成'%cols1[i])



if __name__ == '__main__':
    print('表格123人名后不能有空格')
    time.sleep(2)
    print('每次整理10人，电话截图在文件夹images中')
    time.sleep(2)
    print('程序即将开始，请将系统窗口靠左')
    time.sleep(5)
    over_all()
    print('前10人已完成，请及时删除')
