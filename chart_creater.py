import openpyxl
from matplotlib.font_manager import FontProperties as font
import matplotlib.pyplot as plt

font1 = font(fname=r'fonts\NotoSansTC-Black.ttf')

def load_excel_to_data(filename: str) -> list:
    workbook = openpyxl.load_workbook(filename, read_only=True)
    worksheet = workbook.active
    data = [[]for _ in range(2)]
    for row in range(2):
        for col in range(1, worksheet.max_column):
            data[row].append(worksheet.cell(row=row+1, column=col+1).value)
    return data

def create_chart(filename: str, title: str, x_label: str, y_label: str, chart_type: str):
    data = load_excel_to_data(filename)
    if chart_type == 'line':
        plt.plot(data[0], data[1])
    elif chart_type == 'bar':
        plt.bar(data[0], data[1])
    plt.title(title, fontproperties=font1, fontsize=20)
    plt.xlabel(x_label, fontproperties=font1, fontsize=15)
    plt.ylabel(y_label, fontproperties=font1, fontsize=15)
    plt.show()
