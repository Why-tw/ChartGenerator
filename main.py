from PyQt6 import QtWidgets
from chart_creater import create_chart, load_excel_to_data, analyze_data
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import sys
import os

# 自動排列介面元件
start = 30
r = 70

class MainWindows(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Chart Generator')
        info =  analyze_data(load_excel_to_data('data.xlsx'))
        self.x_min = info[2]
        self.x_max = info[3]
        self.y_min = info[4]
        self.y_max = info[5]
        self.ui()
        self.update_chart() # 初始化圖表

    def ui(self):
        # 圖表顯示區
        self.grview = QtWidgets.QGraphicsView(self)
        self.grview.setGeometry(25, 25, 830, 480)
        self.grview.setStyleSheet('background-color:#6C6C6C')
        self.scene = QtWidgets.QGraphicsScene()
        self.scene.setSceneRect(15, 25, 800, 420)
        self.grview.setScene(self.scene)
        
        # 圖表更新按鈕
        update_btn = QtWidgets.QPushButton(self)
        update_btn.setText('更新圖表')
        update_btn.setGeometry(23, start+7*r, 100, 30)
        update_btn.setStyleSheet('font-size:20px')
        update_btn.clicked.connect(self.update_chart)

        # 圖表類型選擇按鈕
        chart_type_label = QtWidgets.QLabel(self)
        chart_type_label.setText('圖表類型:')
        chart_type_label.setGeometry(950, start, 100, 30)
        chart_type_label.setStyleSheet('font-size:20px')
        self.chart_type = QtWidgets.QComboBox(self)
        self.chart_type.addItem('折線圖')
        self.chart_type.addItem('長條圖')
        self.chart_type.setGeometry(1050, start, 200, 30)
        self.chart_type.setStyleSheet('font-size:20px')

        # 圖表標題輸入框
        chart_title_label = QtWidgets.QLabel(self)
        chart_title_label.setText('圖表標題:')
        chart_title_label.setGeometry(950, start+r, 100, 30)
        chart_title_label.setStyleSheet('font-size:20px')
        self.chart_title = QtWidgets.QLineEdit(self)
        self.chart_title.setGeometry(1050, start+r, 200, 30)
        self.chart_title.setStyleSheet('font-size:20px')
        self.chart_title.setText('圖表')

        # X軸標籤輸入框
        x_label_label = QtWidgets.QLabel(self)
        x_label_label.setText('X軸標籤:')
        x_label_label.setGeometry(950, start+2*r, 100, 30)
        x_label_label.setStyleSheet('font-size:20px')
        self.x_label = QtWidgets.QLineEdit(self)
        self.x_label.setGeometry(1050, start+2*r, 200, 30)
        self.x_label.setStyleSheet('font-size:20px')
        self.x_label.setText('X軸標籤')

        # Y軸標籤輸入框
        y_label_label = QtWidgets.QLabel(self)
        y_label_label.setText('Y軸標籤:')
        y_label_label.setGeometry(950, start+3*r, 100, 30)
        y_label_label.setStyleSheet('font-size:20px')
        self.y_label = QtWidgets.QLineEdit(self)
        self.y_label.setGeometry(1050, start+3*r, 200, 30)
        self.y_label.setStyleSheet('font-size:20px')
        self.y_label.setText('Y軸標籤')

        # 儲存圖表按鈕
        save_btn = QtWidgets.QPushButton(self)
        save_btn.setText('儲存圖表')
        save_btn.setGeometry(525, start+7*r, 100, 30)
        save_btn.setStyleSheet('font-size:20px')
        save_btn.clicked.connect(self.save_chart)

        # 表格點樣式選擇按鈕
        self.point_style_label = QtWidgets.QLabel(self)
        self.point_style_label.setText('點樣式:')
        self.point_style_label.setGeometry(950, start+4*r, 100, 30)
        self.point_style_label.setStyleSheet('font-size:20px')
        self.point_style = QtWidgets.QComboBox(self)
        self.point_style.setGeometry(1050, start+4*r, 200, 30)
        self.point_style.setStyleSheet('font-size:20px')
        self.point_style.addItem('無')
        self.point_style.addItem('圓形')
        self.point_style.addItem('方形')
        self.point_style.addItem('三角形')

        # 顯示範圍
        range_label = QtWidgets.QLabel(self)
        range_label.setText('顯示範圍')
        range_label.setGeometry(950, start+5*r, 100, 30)
        range_label.setStyleSheet('font-size:20px')

        xrange_label = QtWidgets.QLabel(self)
        xrange_label.setText('X軸:')
        xrange_label.setGeometry(1050, start+5*r, 50, 30)
        xrange_label.setStyleSheet('font-size:20px')
        self.range_xmin = QtWidgets.QLineEdit(self)
        self.range_xmin.setGeometry(1100, start+5*r, 50, 30)
        self.range_xmin.setStyleSheet('font-size:20px')
        self.range_xmin.setText(str(self.x_min-10))
        rlabel1 = QtWidgets.QLabel(self)
        rlabel1.setText('~')
        rlabel1.setGeometry(1160, start+5*r, 50, 30)
        rlabel1.setStyleSheet('font-size:30px')
        self.range_xmax = QtWidgets.QLineEdit(self)
        self.range_xmax.setGeometry(1190, start+5*r, 50, 30)
        self.range_xmax.setStyleSheet('font-size:20px')
        self.range_xmax.setText(str(self.x_max+10))

        yrange_label = QtWidgets.QLabel(self)
        yrange_label.setText('Y軸:')
        yrange_label.setGeometry(1050, 450, 50, 30)
        yrange_label.setStyleSheet('font-size:20px')
        self.range_ymin = QtWidgets.QLineEdit(self)
        self.range_ymin.setGeometry(1100, 450, 50, 30)
        self.range_ymin.setStyleSheet('font-size:20px')
        self.range_ymin.setText(str(self.y_min-10))
        rlabel2 = QtWidgets.QLabel(self)
        rlabel2.setText('~')
        rlabel2.setGeometry(1160, 450, 50, 30)
        rlabel2.setStyleSheet('font-size:30px')
        self.range_ymax = QtWidgets.QLineEdit(self)
        self.range_ymax.setGeometry(1190, 450, 50, 30)
        self.range_ymax.setStyleSheet('font-size:20px')
        self.range_ymax.setText(str(self.y_max+10))

        # 數據顯示勾選
        self.data_display = QtWidgets.QCheckBox(self)
        self.data_display.setText('數據顯示')
        self.data_display.setGeometry(950, start+7*r, 100, 30)
        self.data_display.setStyleSheet('font-size:20px')

    def update_chart(self):
        self.fig = create_chart('data.xlsx', self.chart_title.text(),
                                 self.x_label.text(), self.y_label.text(),
                                   self.chart_type.currentText(), int(self.range_xmin.text()),
                                     int(self.range_xmax.text()), int(self.range_ymin.text()),
                                       int(self.range_ymax.text()), self.point_style.currentText(),
                                       self.data_display.isChecked())
        if self.fig:
            canvas = FigureCanvas(self.fig)
            self.scene.clear()
            self.scene.addWidget(canvas)
            self.grview.setScene(self.scene)
    
    def save_chart(self):
        if os.path.exists('outputs/chart.png'):
            os.remove('outputs/chart.png')
        self.fig.savefig('outputs/chart.png')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    Form = MainWindows()
    Form.showMaximized()
    sys.exit(app.exec())
