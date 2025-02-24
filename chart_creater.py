import openpyxl
from matplotlib.font_manager import FontProperties as font
import matplotlib.pyplot as plt

font1 = font(fname=r'fonts\NotoSansTC-Black.ttf')

import openpyxl

def load_excel_to_data(filename: str) -> list:
    workbook = openpyxl.load_workbook(filename, read_only=True)
    worksheet = workbook.active
    data = [[], []]
    for i in range(2):
        for j in range(2, worksheet.max_column + 1):
            data[i].append(worksheet.cell(row=i + 1, column=j).value)
    return data


def create_chart(filename: str, title: str, x_label: str, y_label: str, chart_type: str):
    data = load_excel_to_data(filename)
    fig, ax = plt.subplots(figsize=(6, 3.7))
    if chart_type == '折線圖':
        ax.plot(data[0], data[1])
    elif chart_type == '長條圖':
        ax.bar(data[0], data[1])
    ax.set_title(title, fontproperties=font1, fontsize=20)
    ax.set_xlabel(x_label, fontproperties=font1, fontsize=15)
    ax.set_ylabel(y_label, fontproperties=font1, fontsize=15)
    return fig

