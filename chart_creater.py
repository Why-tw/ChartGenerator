import openpyxl
from matplotlib.font_manager import FontProperties as font
import matplotlib.pyplot as plt

font1 = font(fname=r'fonts\NotoSansTC-Black.ttf')
import openpyxl

def load_excel_to_data(filename: str) -> list:
    wb = openpyxl.load_workbook('data.xlsx')
    ws = wb.active
    data = [[]for _ in range(ws.max_row)]
    x_list = []
    y_list = []
    result = []

    for i in range(1, ws.max_row+1):
        for j in range(1, ws.max_column+1):
            data[i-1].append(ws.cell(row=i, column=j).value)

    for i in range(2, ws.max_row+1):
        x_list.append(data[i-1][0])
        y_list.append(data[i-1][1])
    
    result.append(x_list)
    result.append(y_list)
    return result

def create_chart(filename: str, title: str, x_label: str,
                  y_label: str, chart_type: str) -> plt.figure:
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

print(load_excel_to_data('data.xlsx'))
