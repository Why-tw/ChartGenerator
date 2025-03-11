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
                  y_label: str, chart_type: str,
                  x_min: int, x_max: int, y_min: int, y_max: int, 
                  point_type_text: str, data_display: bool) -> plt.figure:
    data = load_excel_to_data(filename)
    point_type=''
    if point_type_text == '無':
        point_type = ''
    elif point_type_text == '圓形':
        point_type = 'o'
    elif point_type_text == '方形':
        point_type = 's'
    elif point_type_text == '三角形':
        point_type = '^'
    fig, ax = plt.subplots(figsize=(6, 3.85))
    #　選擇圖表類型
    if chart_type == '折線圖':
        ax.plot(data[0], data[1], marker=point_type)
    elif chart_type == '長條圖':
        ax.bar(data[0], data[1], marker=point_type)
    # 設置顯示範圍
    plt.xlim(x_min, x_max)
    plt.ylim(y_min, y_max)

    # 設置標題和標籤
    ax.set_title(title, fontproperties=font1, fontsize=20)
    ax.set_xlabel(x_label, fontproperties=font1, fontsize=15)
    ax.set_ylabel(y_label, fontproperties=font1, fontsize=15)
    
    # 數據顯示
    if data_display:
        for a, b in zip(data[0], data[1]):
            ax.text(a, b, b, ha='center', va='bottom', fontsize=10)
    
    return fig
