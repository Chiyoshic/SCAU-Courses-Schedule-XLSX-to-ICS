from tkinter import *
import tkinter.filedialog
from tkinter.messagebox import *
import openpyxl
import courseget
import courseclass
import datetime
import os
import sys


weekOptions = [
    1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
]
reminderOptions = {
    '5mins':5,
    '10mins':10,
    '15mins':15,
    '30mins':30,
    '1hr':60,
    '2hrs':120,
    '1day':1440,

}


def selectInputPath():
    path_ = tkinter.filedialog.askopenfilename(title="选择一个文件", filetypes = [("Excel文件",'.xlsx'),("Excel文件",'.xls')])
    print(path_)
    inputPath.set(path_)

def selectOutputPath():
    path_ = tkinter.filedialog.askdirectory()
    print(path_)
    outputPath.set(path_)

"""def transform():"""


main = Tk()
main.title("SCAU个人课表转ics小工具 created by 2024CS6王崇熙")
main.geometry("650x150")
main.resizable(False, False)

inputPath = StringVar()
inputLB = (Label(main, text="导入文件路径：")).grid(row = 0, column = 0)
inputEntry = Entry(main, width = 60, textvariable = inputPath)
inputEntry.grid(row=0, column=1, columnspan= 4)
inputButton = Button(main, text="选择文件", command=selectInputPath).grid(row=0, column=5)

outputPath = StringVar()
outputLB = (Label(main, text="输出文件路径：")).grid(row = 1, column = 0)
outputEntry = Entry(main, width = 60, textvariable = outputPath)
outputEntry.grid(row=1, column=1, columnspan= 4)
outputButton = Button(main, text="选择路径", command=selectOutputPath).grid(row=1, column=5)

firstDateLB = Label(main, text="学期第一天(YYYYMMDD)：").grid(row = 2, column = 0)
firstDateEntry = Entry(main)
firstDateEntry.grid(row=2, column=1)
firstDateEntry.insert(0, "20240902")

weekChooseLB = Label(main, text="选择周数：").grid(row=2, column=2)
weekVar = StringVar()
weekVar.set(weekOptions[0])
weekChooseList = OptionMenu(main, weekVar,*weekOptions).grid(row=2, column=3)

reminder1ChooseLB = Label(main, text="第一次提醒时间：").grid(row=3, column=0)
reminder1Var = StringVar()
reminder1Var.set('5mins')
reminder1ChooseList = OptionMenu(main, reminder1Var,*reminderOptions.keys()).grid(row=3, column=1)

reminder2ChooseLB = Label(main, text="第二次提醒时间：").grid(row=3, column=2)
reminder2Var = StringVar()
reminder2Var.set('10mins')
reminder2ChooseList = OptionMenu(main, reminder2Var,*reminderOptions.keys()).grid(row=3, column=3)



def transform():
    fileName = f"{int(weekVar.get())} week from {firstDateEntry.get()}"
    inputFile = inputEntry.get()
    outputFile = f"{outputEntry.get()}/{fileName}.ics"
    nowWeek = int(weekVar.get())
    reminderMins1 = int(reminderOptions[reminder1Var.get()])
    reminderMins2 = int(reminderOptions[reminder2Var.get()])
    firstDateOfTheTerm = firstDateEntry.get()

    if (inputFile == ''):
        showerror('错误', '输入路径空白')
        return
    
    if (outputEntry.get() == ''):
        showerror('错误', '输出路径空白')
        return
    
    if (courseget.isDateValid(firstDateOfTheTerm) == False):
        showerror('错误', '日期输入非法')
        return

    startingDate = courseget.getTheNowDate(firstDateOfTheTerm, nowWeek)

    startingDateDatetime = datetime.datetime.strptime(startingDate, "%Y%m%d")
    weekList = courseget.makeAListOfWeek(startingDateDatetime)
    """make a weeklist which stores the date of a certain week of the course list."""

    startTimeDic = {4: '000000', 5: '015500', 6: '044000', 7:'063000', 8:'082500', 9:'113000'}
    endTimeDic = {4: '012500', 5: '040500', 6: '060500', 7:'075500', 8:'103500', 9:'134000'}
    """define the course time"""


    file = open(outputFile, "w", encoding= 'utf-8') 
    file.write("BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\nMETHOD:PUBLISH\n\n")
    """create the ical file head"""

    for colunm in range(2, 9):
        for row in range(4, 10):
            weekListColunm = colunm - 2
            """set the cell location"""
            nowDate = weekList[weekListColunm]
            sequence = row - 4

            reminderTime1 = courseget.setReminderTime(nowDate, startTimeDic[row], reminderMins1)
            reminderTime2 = courseget.setReminderTime(nowDate, startTimeDic[row], reminderMins2)
            """set the reminder time"""

            cell_obj = courseget.getaCourse(inputFile, row, colunm, nowWeek) 
            """get cell in the course excel"""

            """print(f"this is the cell:{cell_obj}")"""

            if cell_obj is None:
                pass
            else:
                arrangement = courseget.divide(cell_obj)
                """divide the gotten content into a list"""


                """print(f"after get is{arrangement}\n")"""
                
                cUID = courseget.makeUID(nowDate, sequence)

                if arrangement[0].startswith('体育'):
                    cLocation = courseget.getLocation(arrangement[4])
                    cCourse = courseclass.course(arrangement[0], f'老师: {arrangement[1]}，类型:{arrangement[2]}，编号：{arrangement[3]}，上课地点：{arrangement[4]}，班别：{arrangement[5]}。',startTimeDic[row], endTimeDic[row],f'华南农业大学{cLocation}', cUID)
                elif arrangement[2].startswith('实验'):
                    cLocation = courseget.getLocation('校园内')
                    cCourse = courseclass.course(arrangement[0], f'老师: {arrangement[1]}，类型:{arrangement[2]}，编号：{arrangement[3]}，{arrangement[4]}。，班别：{arrangement[5]}。',startTimeDic[row], endTimeDic[row],f'华南农业大学{cLocation}', cUID)
                else:
                    cLocation = courseget.getLocation(arrangement[4])
                    cCourse = courseclass.course(arrangement[0], f'老师: {arrangement[1]}，类型:{arrangement[2]}，编号：{arrangement[3]}，上课地点：{arrangement[4]}，{arrangement[5]}。，班别：{arrangement[6]}。',startTimeDic[row], endTimeDic[row],f'华南农业大学{cLocation}', cUID)

                """body making[begin]"""
                file.write("BEGIN:VEVENT\n")

                file.write(f"DTSTART:{nowDate}T{startTimeDic[row]}Z\n")
                file.write(f"DTEND:{nowDate}T{endTimeDic[row]}Z\n")
                file.write(f"UID:{cCourse.UID}\n")
                file.write(f"DESCRIPTION:{cCourse.des}\n")
                file.write(f"LOCATION:{cCourse.location}\n")
                file.write(f"SEQUENCE:{sequence}\n")
                file.write("STATUS:CONFIRMED\n")
                file.write(f"SUMMARY:{cCourse.courseName}\n")
                file.write("TRANSP:OPAQUE\n")

                """make reminder1[begin]"""
                file.write("\nBEGIN:VALARM\nACTION:DISPLAY\n")
                file.write(f"TRIGGER:{reminderTime1}\n")
                file.write("DESCRIPTION:This is an event reminder\nEND:VALARM\n")
                """make reminder1[begin]"""

                """make reminder2[begin]"""
                file.write("\nBEGIN:VALARM\nACTION:DISPLAY\n")
                file.write(f"TRIGGER:{reminderTime2}\n")
                file.write("DESCRIPTION:This is an event reminder\nEND:VALARM\n\n")
                """make reminder2[begin]"""

                file.write("END:VEVENT\n")
                """body making[end]"""

    file.write("\nEND:VCALENDAR")
    """create the ical file bottom"""
    file.close

    showinfo('提示', '转换成功')



beginButton = Button(main, text="开始转换", command=transform, height=2).grid(row=2, column=5, rowspan=2)


main.mainloop()