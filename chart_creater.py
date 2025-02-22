import openpyxl
from matplotlib.font_manager import FontProperties as font
import matplotlib.pyplot as plt

font1 = font(fname=r'fonts\NotoSansTC-Black.ttf')

def load_excel_to_data(filename: str) -> list:
    workbook = openpyxl.load_workbook(filename, read_only=True)
    worksheet = workbook.active
    data = [[], []]
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        data[0].append(row[0])
        data[1].append(row[1])
    return data

def create_chart(filename: str, title: str, x_label: str, y_label: str, chart_type: str):
    data = load_excel_to_data(filename)
    fig, ax = plt.subplots()
    if chart_type == 'line':
        ax.plot(data[0], data[1])
    elif chart_type == 'bar':
        ax.bar(data[0], data[1])
    ax.set_title(title, fontproperties=font1, fontsize=20)
    ax.set_xlabel(x_label, fontproperties=font1, fontsize=15)
    ax.set_ylabel(y_label, fontproperties=font1, fontsize=15)
    plt.savefig('temp/chart.png')
    plt.close(fig)
