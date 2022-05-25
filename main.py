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

from scipy import stats
from scipy.stats import spearmanr

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

    ws_rename = wb['Sheet']
    ws_rename.title = "Values"
    sh1 = wb["Values"]

    ws1 = wb.create_sheet("Predictions", 1)
    ws1.title = "Predictions"
    sh2 = wb["Predictions"]

    wb.save(filename=str(a.get() + ".xlsx"))

    ws.cell(column=1, row=1).value = "Y"

    for i in range(1, int(b.get()) + 1):
        sh1.cell(column=1 + i, row=1).value = "X" + str(i)

    sh2.cell(column=1, row=1).value = "Y"

    for i in range(1, int(b.get()) + 1):
        sh2.cell(column=1 + i, row=1).value = "X" + str(i)

    wb.save(filename=str(a.get() + ".xlsx"))

    os.startfile(str(a.get() + ".xlsx"))


# calculations
def func2(c=str(a.get())):
    wb = op.load_workbook(str(a.get() + ".xlsx"))
    ws = wb.active

    df = pd.read_excel(str(a.get() + ".xlsx"), sheet_name='Values')
    df2 = pd.read_excel(str(a.get() + ".xlsx"), sheet_name="Predictions")
    y = df["Y"]
    x = df.drop("Y", axis=1)
    x = sm.add_constant(x)
    x_pred = df2.drop("Y", axis=1)
    x_pred = sm.add_constant(x_pred, has_constant='add')
    model = sm.OLS(y, x).fit()
    model_sk = LinearRegression().fit(x, y)
    y_pred = model.predict()
    y_av = statistics.mean(y)
    # x_np, y_np = np.array(x), np.array(y)

    # print(model.summary())
    # print(model.params)

    # a) find a regression equation of Y on Х1 and X2 and explain the meaning of regression coefficients b0, b1, b2
    pr = model.params
    g = str(round((pr[0]), 4))
    for i in range(0, len(pr)):
        if pr[i] > 0:
            g += str("+" + str(round((pr[i]), 4)) + "x" + str(i))
        elif pr[i] < 0:
            g += str(str(round((pr[i]), 4)) + "x" + str(i))
        elif pr[i] == 0:
            g = g
    # print(a)

    # b) estimate the tightness and the direction of the relation between variables Х1, X2 and Y by computing multiple
    # correlation coefficient
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

    SSR = np.sum([(i - y.mean()) ** 2 for i in y_pred])
    SST = np.sum([(i - y.mean()) ** 2 for i in y])
    SSE_array = np.subtract(y, y_pred)
    SSE = np.sum([i ** 2 for i in SSE_array])

    r2 = model_sk.score(x, y)
    r = math.sqrt(r2)

    # print(r)
    # print(corr)
    # plt.show()

    # c) estimate the significance of the regression equation of Y on Х1, X2 by F-statistics at level α=0.05
    p = int(b.get())
    dof = len(y) - p - 1
    # alpha = float(c.get())
    # confidence = 1 - alpha
    confidence = 0.95

    f_crit = np.abs(f.ppf(confidence, p, dof))
    F = (r2 / (1 - r2)) * (dof / p)

    # d) estimate the significance of regression coefficients b0, b1, b2 by t-statistics at level α=0.05
    t_crit = np.abs(t.ppf((1 - confidence) / 2, dof))

    S2 = SSE / dof

    xt = np.matrix.transpose(x.to_numpy())
    mult1 = np.dot(xt, x)
    inv = np.linalg.inv(mult1)
    mult2 = np.dot(inv, xt)

    b_array = np.dot(mult2, y)

    # e) find determination coefficient and explain its meaning
    r2 = model_sk.score(x, y)

    # f) estimate 95% confidence interval for the average company’s profit, given that X1 = 11, X2 = 8
    x_predt = np.matrix.transpose(x_pred.to_numpy())
    mult3 = np.dot(x_pred, inv)
    mult4 = np.dot(mult3, x_predt)
    # this is useless, i found formula in module

    y_new = model.predict(x_pred)
    conf1 = np.array([y_new - t_crit * math.sqrt(S2 * mult4), y_new + t_crit * math.sqrt(S2 * mult4)])

    # g) estimate 95% confidence interval for the individual value of company’s profit, given that X1 = 11, X2 = 8
    conf2 = np.array([y_new - t_crit * math.sqrt(S2 * (1 + mult4)), y_new + t_crit * math.sqrt(S2 * (1 + mult4))])

    # h) find an interval estimator for regression coefficients β0, β1, β2
    t_crit = np.abs(t.ppf((1 - confidence) / 2, dof))
    varb = []
    for i in range(0, int(b.get()) + 1):
        varb.append(math.sqrt(S2 * inv[i, i]))

    T = []
    for i in range(0, int(b.get()) + 1):
        T.append(abs(b_array[i]) / varb[i])

    conf3 = []
    for i in range(0, int(b.get()) + 1):
        conf3.append([b_array[i] - varb[i] * t_crit, b_array[i] + varb[i] * t_crit])

    # i) find an interval estimator for the error’s variance with 0.95 confidence
    chi2_1 = np.abs(chi2.ppf(1 - (1 - confidence) / 2, dof))
    chi2_2 = np.abs(chi2.ppf((1 - confidence) / 2, dof))
    conf4 = np.array([S2 * dof / chi2_1, S2 * dof / chi2_2])
    # print(conf4)

    # 2. a) apply the Spearman rank correlation test to assess heteroscedasticity at a 5% significance level for
    # both x1 and x2
    # s_coef = []
    # for i in range(0, int(b.get())+1):
    #     s_coef.append(spearmanr(x.to_numpy()[i], y.to_numpy()))
    x_new = np.zeros([len(y), int(b.get())])

    for i in range(1, int(b.get()) + 1):
        # x_new.append(df["X" + str(i)].to_numpy)
        x_new[:, i - 1] = df["X" + str(i)]
    model_array = np.zeros([len(y), int(b.get())])
    # model12 = LinearRegression().fit(x_new[:, 0].reshape(10, 1), y)
    counter = 0
    while counter < int(b.get()):
        model11 = sm.OLS(y, x_new[:, counter].reshape(-1, 1)).fit()
        # print(model11.summary())
        # for i in range(0, int(b.get())):
        model_array[:, counter] = model11.resid
        counter += 1
        # print(model_array)
    # print(model_array[:, 1])
    # print(model_array, '\n', x_new)

    # counter = 0
    # while counter < int(b.get()):
    #     print(spearmanr(x_new[:, counter], model_array[:, counter]))
    #     counter += 1

    # while counter < int(b.get()):
    #     for i in range ()
    argsort_array = x_new[:, 0].argsort()
    ranks_array = np.empty_like(argsort_array)
    ranks_array[argsort_array] = np.arange(len(y))
    duplicate_check = []
    # for i in argsort_array:
    #     for j in range(0, len(y)):
    #         if argsort_array[i] == argsort_array[j]:
    #             duplicate_check.append(i)
    #             duplicate_check.append(j)
    def duplicates(argsort_array):
        return [elem in argsort_array[:i] for i, elem in enumerate(argsort_array)]


    print(duplicate_check)

    print("\nRank of each item of the said array:")
    print(ranks_array)
    # s_coef = np.zeros([int(b.get())])
    # for i in range(0, int(b.get())+1):
    #     s_coef = spearmanr(x_new[i], model_array[i])
    # print(s_coef)
    # y_pred_new = []
    # for i in range(0, int(b.get())):
    #     y_pred_new.append(model_array.predict(x_new[i]))
    # print(y_pred_new)


# buttons
new_excel = Button(tsk, text="Create an excel file", command=func1)
new_excel.place(x=160, y=90)

calc_button = Button(tsk, text="*The next button*", command=func2)
calc_button.place(x=162, y=150)

graph = Button(tsk, text="Create a graph")
graph.place(x=170, y=180)

mainloop()
