import openpyxl as op
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import math
import statistics

from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score

import statsmodels.api as sm
from scipy.stats import t
from scipy.stats import f
from scipy.stats import chi2

import tkinter as tk
from tkinter import *
import os

# # creating a window
# tsk = Tk()
# tsk.title("Econometrics - seminar 2")
#
# tsk.configure(width=400, height=210)
# tsk.configure(bg='lightgray')
#
# lab_1 = Label(tsk, text="Enter your excel file name")
# lab_1.place(x=10, y=30)
# a = Entry(tsk)
# a.place(x=180, y=30)
# a.get()
#
# lab_2 = Label(tsk, text="Enter the number of independent variables")
# lab_2.place(x=10, y=60)
# b = Entry(tsk)
# b.place(x=180, y=60)
# b.get()
#
# lab_3 = Label(tsk, text="After saving your values to the excel file, please, push the next button")
# lab_3.place(x=10, y=120)
#
#
# # creating an excel file
# def func1():
#     wb = Workbook()
#     ws = wb.active
#
#     wb.save(filename=str(a.get() + ".xlsx"))
#
#     ws.cell(column=1, row=1).value = "Y"
#
#     for i in range(1, int(b.get())+1):
#         ws.cell(column=1 + i, row=1).value = "X" + str(i)
#
#     wb.save(filename=str(a.get() + ".xlsx"))
#
#     os.startfile(str(a.get() + ".xlsx"))
#
# # input values
# def func2():
#     wb = op.load_workbook(str(a.get() + ".xlsx"))
#     ws = wb.active
#
#
#
#
#
# # buttons
# new_excel = Button(tsk, text="Create an excel file", command=func1)
# new_excel.place(x=160, y=90)
#
# calc_button = Button(tsk, text="*The next button*")
# calc_button.place(x=162, y=150)
#
# graph = Button(tsk, text="Create a graph")
# graph.place(x=170, y=180)
#
# mainloop()

df = pd.read_excel("25635247.xlsx")
y = df["Y"]
x = df.drop("Y", axis=1)
x = sm.add_constant(x)
model = sm.OLS(y, x).fit()
model_sk = LinearRegression().fit(x, y)
y_pred = model.predict()
y_av = statistics.mean(y)
# x_np, y_np = np.array(x), np.array(y)

# print(model.summary())
# print(model.params)

# a)
pr = model.params
a = str(round((pr[0]), 4))
for i in range(0, len(pr)):
    if pr[i] > 0:
        a += str("+" + str(round((pr[i]), 4)) + "x" + str(i))
    elif pr[i] < 0:
        a += str(str(round((pr[i]), 4)) + "x" + str(i))
    elif pr[i] == 0:
        a = a
# print(a)

# b)
corr = df.corr()
fig = plt.figure()
ax = fig.add_subplot(111)
cax = ax.matshow(corr, cmap='coolwarm', vmin=-1, vmax=1)
fig.colorbar(cax)
ticks = np.arange(0, len(df.columns), 1)
ax.set_xticks(ticks)
plt.xticks(rotation=90)
ax.set_yticks(ticks)
ax.set_xticklabels(df.columns)
ax.set_yticklabels(df.columns)
r = math.sqrt(model_sk.score(x, y))
# print(r)
# print(corr)
# plt.show()
