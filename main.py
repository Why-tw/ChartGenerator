from PyQt6 import QtWidgets
from chart_creater import create_chart
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import sys
import os

class MainWindows(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Chart Generator')
        self.ui()

    def ui(self):
        # 圖表顯示區
        self.grview = QtWidgets.QGraphicsView(self)
        self.grview.setGeometry(25, 25, 600, 400)
        self.grview.setStyleSheet('background-color:#6C6C6C')
        self.scene = QtWidgets.QGraphicsScene()
        self.scene.setSceneRect(0, 0, 590, 390)
        self.grview.setScene(self.scene)
        
        # 圖表更新按鈕
        update_btn = QtWidgets.QPushButton(self)
        update_btn.setText('更新圖表')
        update_btn.setGeometry(23, 430, 100, 30)
        update_btn.setStyleSheet('font-size:20px')
        update_btn.clicked.connect(self.update_chart)

        # 圖表類型選擇按鈕
        chart_type_label = QtWidgets.QLabel(self)
        chart_type_label.setText('圖表類型:')
        chart_type_label.setGeometry(700, 30, 100, 30)
        chart_type_label.setStyleSheet('font-size:20px')
        self.chart_type = QtWidgets.QComboBox(self)
        self.chart_type.addItem('折線圖')
        self.chart_type.addItem('長條圖')
        self.chart_type.setGeometry(800, 30, 200, 30)
        self.chart_type.setStyleSheet('font-size:20px')

        # 圖表標題輸入框
        chart_title_label = QtWidgets.QLabel(self)
        chart_title_label.setText('圖表標題:')
        chart_title_label.setGeometry(700, 100, 100, 30)
        chart_title_label.setStyleSheet('font-size:20px')
        self.chart_title = QtWidgets.QLineEdit(self)
        self.chart_title.setGeometry(800, 100, 200, 30)
        self.chart_title.setStyleSheet('font-size:20px')

        # X軸標籤輸入框
        x_label_label = QtWidgets.QLabel(self)
        x_label_label.setText('X軸標籤:')
        x_label_label.setGeometry(700, 170, 100, 30)
        x_label_label.setStyleSheet('font-size:20px')
        self.x_label = QtWidgets.QLineEdit(self)
        self.x_label.setGeometry(800, 170, 200, 30)
        self.x_label.setStyleSheet('font-size:20px')

        # Y軸標籤輸入框
        y_label_label = QtWidgets.QLabel(self)
        y_label_label.setText('Y軸標籤:')
        y_label_label.setGeometry(700, 240, 100, 30)
        y_label_label.setStyleSheet('font-size:20px')
        self.y_label = QtWidgets.QLineEdit(self)
        self.y_label.setGeometry(800, 240, 200, 30)
        self.y_label.setStyleSheet('font-size:20px')

        # 儲存圖表按鈕
        save_btn = QtWidgets.QPushButton(self)
        save_btn.setText('儲存圖表')
        save_btn.setGeometry(525, 430, 100, 30)
        save_btn.setStyleSheet('font-size:20px')
        save_btn.clicked.connect(self.save_chart)

        # 表格點樣式選擇按鈕
        point_style_label = QtWidgets.QLabel(self)
        point_style_label.setText('點樣式:')
        point_style_label.setGeometry(700, 310, 100, 30)
        point_style_label.setStyleSheet('font-size:20px')
        point_style = QtWidgets.QComboBox(self)
        point_style.setGeometry(800, 310, 200, 30)
        point_style.setStyleSheet('font-size:20px')
        point_style.addItem('圓形')
        point_style.addItem('方形')
        point_style.addItem('三角形')

        # # 最適直線按鈕
        # best_line_check = QtWidgets.QCheckBox(self)
        # best_line_check.setText('最適直線')
        # best_line_check.setGeometry(700, 380, 100, 30)
        # best_line_check.setStyleSheet('font-size:20px')

    def update_chart(self):
        self.fig = create_chart('data.xlsx', self.chart_title.text(),
                                 self.x_label.text(), self.y_label.text(),
                                   self.chart_type.currentText())
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