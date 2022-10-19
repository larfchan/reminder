import pyautogui as pg
import pyperclip as pc
import sys
import time
import xlrd


inpath = r"D:\微信文件\WeChat Files\wxid_0ejsryummijk21\FileStorage\File\2022-10\testcoding.xls"  # excel文件所在路径
data = xlrd.open_workbook(inpath, encoding_override='utf-8')
table = data.sheets()[0]  # 选定第一张表
nrows = table.nrows  # 获取行号

for i in range(1, nrows):  # 第0行为表头
    alldata = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
    user_name = alldata[0]  # 取出表中第0列数据：学生姓名
    RNA_days = alldata[1]  # 取出表中第1列数据：核酸天数
    class_name = alldata[2]  # 取出表中第2列数据：班级
    if class_name != '电子20-1': #判断班级
        continue
    print(user_name, RNA_days, class_name)


    class SendMsg(object):  # 自动发送程序

        def __init__(self):
            self.name = user_name
            self.msg = '''{}的：\n\n{}\n\n同学，您的核酸天数是:\n\n{}\n\n，请在今天下午17点前完成核酸检测哦！(python机器人自动发送，免回复）'''.format(
                class_name, user_name, RNA_days)
            # 发送不同的消息

        def send_msg(self):
            # 操作间隔为1秒
            pg.PAUSE = 1
            pg.hotkey('ctrl', 'alt', 'w')
            pg.hotkey('ctrl', 'f')

            # 找到好友
            pc.copy(self.name)
            pg.hotkey('ctrl', 'v')
            pg.press('enter')

            # 发送消息
            pc.copy(self.msg)
            pg.hotkey('ctrl', 'v')
            pg.hotkey('ctrl', 'enter')

            # 隐藏微信
            time.sleep(0.5)
            pg.hotkey('ctrl', 'alt', 'w')


    if __name__ == '__main__':
        s = SendMsg()
        # while True:
        for i in range(1):  # 每句话发送几次，如001发送两次=001 001 ，002两次=002、002:
            s.send_msg()
            # n +=/ 1
            #
sys.exit(0)  # 发送完成后，退出