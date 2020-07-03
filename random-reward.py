# coding=utf-8
# auther:wangc
# 2020-07-02

from openpyxl import load_workbook
from random import randint
import tkinter as tk

dict1 = {}
def read_excel():
    wb = load_workbook(filename='123.xlsx')
    ws = wb['Sheet1']
    for row in ws.rows:
        dict1[row[0].value] = row[1].value
    print(dict1)
    return dict1

def lottery(t):

    luncky_num = randint(1, 49)
    print(luncky_num)
    t.insert('end', f"本次抽奖的结果是：{luncky_num}, {dict1[luncky_num]}\n")
    t.update()


def show_gui():
    gui = tk.Tk()
    gui.title('抽奖')
    gui.geometry('900x600+280+100')
    # gui.iconbitmap('icon.ico')
    t = tk.Text(gui, width=50, height=30)
    t.place(x=50, y=50)
    t1 = tk.Text(gui, width=50, height=30)
    t1.place(x=510, y=50)

    t.insert('end', f"奖池如下：\n")
    for key, value in dict1.items():
        t.insert('end', f"{key}:{value}\n")
    t.update()
    b1 = tk.Button(gui, text='抽奖', bg='lightblue', width=15, height=2, command=lambda: lottery(t1))
    b1.place(x=400, y=470)

    gui.mainloop()


if __name__ == '__main__':
    read_excel()
    show_gui()
