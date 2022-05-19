import openpyxl as op
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import numpy as np
import matplotlib.pyplot as plt
import math

from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score

import statsmodels.api as sm
from scipy.stats import t
from scipy.stats import f
from scipy.stats import chi2

import tkinter as tk
from tkinter import *
import os

# creating a window
tsk = Tk()
tsk.title("Econometrics - seminar 2")

tsk.configure(width=400, height=210)
tsk.configure(bg='lightgray')

lab_1 = Label(tsk, text="Enter your excel file name")
lab_1.place(x=10, y=30)
a = Entry(tsk)
a.place(x=180, y=30)
a.get()

lab_2 = Label(tsk, text="Enter the number of independent variables")
lab_2.place(x=10, y=60)
b = Entry(tsk)
b.place(x=180, y=60)
b.get()

lab_3 = Label(tsk, text="After saving your values to the excel file, please, push the next button")
lab_3.place(x=10, y=120)


# creating an excel file
def func1():
    wb = Workbook()
    ws = wb.active

    wb.save(filename=str(a.get() + ".xlsx"))

    ws.cell(column=1, row=1).value = "Y"

    for i in range(1, int(b.get())+1):
        ws.cell(column=1 + i, row=1).value = "X" + str(i)

    wb.save(filename=str(a.get() + ".xlsx"))

    os.startfile(str(a.get() + ".xlsx"))

# input values
def func2():
    wb = op.load_workbook(str(a.get() + ".xlsx"))
    ws = wb.active





# buttons
new_excel = Button(tsk, text="Create an excel file", command=func1)
new_excel.place(x=160, y=90)

calc_button = Button(tsk, text="*The next button*")
calc_button.place(x=162, y=150)

graph = Button(tsk, text="Create a graph")
graph.place(x=170, y=180)

mainloop()
1