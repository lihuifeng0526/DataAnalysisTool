import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLabel
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
import abnormal_order_data
import apriori_app
import fee  # 新增导入 fee.py

class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.apriori_window = None
        self.order_processor_window = None
        self.fee_window = None  # 新增 fee_window
        self.initUI()

    def initUI(self):
        self.setWindowTitle('多功能数据分析工具')
        self.setGeometry(100, 100, 400, 300)

        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'icons', 'app_icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            fallback_icon = 'C:\\Windows\\System32\\shell32.dll,4'  # Windows 购物车图标
            self.setWindowIcon(QIcon(fallback_icon))

        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 标题
        title_label = QLabel('选择功能模块', self)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # 按钮：启动异常订单处理
        self.order_processor_button = QPushButton('超区发货异常订单数据(包含超区发货总数据)', self)
        self.order_processor_button.clicked.connect(self.open_order_processor)
        layout.addWidget(self.order_processor_button)

        # 按钮：启动商品关联性分析
        self.apriori_button = QPushButton('商品关联性分析', self)
        self.apriori_button.clicked.connect(self.open_apriori)
        layout.addWidget(self.apriori_button)

        # 按钮：启动超区发货费用数据
        self.fee_button = QPushButton('超区发货费用数据', self)
        self.fee_button.clicked.connect(self.open_fee)
        layout.addWidget(self.fee_button)

        # 填充布局
        layout.addStretch()

    def open_order_processor(self):
        if self.order_processor_window is None:
            self.order_processor_window = abnormal_order_data.OrderDataProcessor(self)
            self.order_processor_window.closed.connect(self.show)  # 连接关闭信号
        self.order_processor_window.show()
        self.hide()

    def open_apriori(self):
        if self.apriori_window is None:
            self.apriori_window = apriori_app.AprioriApp(self)
            self.apriori_window.closed.connect(self.show)  # 连接关闭信号
        self.apriori_window.show()
        self.hide()

    def open_fee(self):
        if self.fee_window is None:
            self.fee_window = fee.OrderDataProcessor(self)  # 使用 fee.py 的 OrderDataProcessor
            self.fee_window.closed.connect(self.show)  # 连接关闭信号
        self.fee_window.show()
        self.hide()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())