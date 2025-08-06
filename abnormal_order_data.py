# abnormal_order_data.py
import sys
import os
import pandas as pd
import re
from datetime import datetime
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QPushButton, 
                             QTextEdit, QFileDialog, QLabel, QProgressBar, QMessageBox,QApplication)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, pyqtSignal

class OrderDataProcessor(QMainWindow):

    closed = pyqtSignal()  # 自定义信号，用于窗口关闭时通知
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.order_files = []
        self.inventory_file = None
        self.output_dir = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle("订单数据处理工具")
        self.setGeometry(100, 100, 600, 500)

        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'icons', 'app_icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            fallback_icon = 'C:\\Windows\\System32\\shell32.dll,4'
            self.setWindowIcon(QIcon(fallback_icon))

        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 订单文件选择按钮
        self.select_order_button = QPushButton("选择订单 Excel 文件", self)
        self.select_order_button.clicked.connect(self.select_order_files)
        layout.addWidget(self.select_order_button)

        # 库存文件选择按钮
        self.select_inventory_button = QPushButton("选择库存 Excel 文件", self)
        self.select_inventory_button.clicked.connect(self.select_inventory_file)
        layout.addWidget(self.select_inventory_button)

        # 运行按钮
        self.run_button = QPushButton("运行处理", self)
        self.run_button.clicked.connect(self.run_processing)
        self.run_button.setEnabled(False)
        layout.addWidget(self.run_button)

        # 返回按钮
        self.back_button = QPushButton("返回主菜单", self)
        self.back_button.clicked.connect(self.close)
        layout.addWidget(self.back_button)

        # 状态标签
        self.status_label = QLabel("状态：等待选择订单和库存文件", self)
        layout.addWidget(self.status_label)

        # 进度条
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # 输出文本区域
        self.output_text = QTextEdit(self)
        self.output_text.setReadOnly(True)
        layout.addWidget(self.output_text)

    def closeEvent(self, event):
        self.closed.emit()  # 发出关闭信号
        event.accept()
        
    def select_order_files(self):
        """打开订单文件选择对话框"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择订单数据 Excel 文件", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if files:
            self.order_files = files
            self.output_dir = os.path.dirname(files[0])
            self.output_text.append(f"已选择 {len(files)} 个订单文件：\n{chr(10).join([os.path.basename(f) for f in files])}")
            self.check_files_selected()
            QApplication.processEvents()

    def select_inventory_file(self):
        """打开库存文件选择对话框"""
        file, _ = QFileDialog.getOpenFileName(
            self, "选择库存数据 Excel 文件", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file:
            self.inventory_file = file
            if not self.output_dir:
                self.output_dir = os.path.dirname(file)
            self.output_text.append(f"已选择库存文件：{os.path.basename(file)}")
            self.check_files_selected()
            # QApplication.processEvents()

    def check_files_selected(self):
        """检查是否已选择订单和库存文件"""
        if self.order_files and self.inventory_file:
            self.output_text.append(f"输出目录：{self.output_dir}")
            self.status_label.setText("状态：订单和库存文件已选择，点击运行处理")
            self.run_button.setEnabled(True)
            self.progress_bar.setValue(0)
        else:
            self.status_label.setText("状态：请同时选择订单和库存文件")
            self.run_button.setEnabled(False)

    def run_processing(self):
        """运行数据处理流程"""
        self.run_button.setEnabled(False)
        self.status_label.setText("状态：正在处理...")
        self.progress_bar.setValue(0)
        self.output_text.append(f"\n=== 开始处理 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===")

        # 总步骤数（数据清洗、省份提取、异常检测及库存合并、商家编码筛选）
        total_steps = 4
        step_value = 100 // total_steps

        # 步骤 1：数据清洗
        self.output_text.append("\n=== 步骤 1：数据清洗 ===")
        cleaned_file = self.data_clean_1()
        self.progress_bar.setValue(step_value)
        # QApplication.processEvents()

        if cleaned_file:
            # 步骤 2：提取省份
            self.output_text.append("\n=== 步骤 2：提取省份 ===")
            processed_file = self.data_clean_2(cleaned_file)
            self.progress_bar.setValue(step_value * 2)
            # QApplication.processEvents()

            if processed_file:
                # 步骤 3：检测异常并合并库存
                self.output_text.append("\n=== 步骤 3：检测异常数据及库存合并 ===")
                abnormal_file = self.abnormal_process(processed_file)
                self.progress_bar.setValue(step_value * 3)
                # QApplication.processEvents()

                if abnormal_file:
                    # 步骤 4：筛选商家编码
                    self.output_text.append("\n=== 步骤 4：筛选商家编码 ===")
                    self.filter_merchant_codes(abnormal_file)
                    self.progress_bar.setValue(100)
                    # QApplication.processEvents()
                else:
                    self.output_text.append("\n异常数据处理失败，终止流程！")
                    QMessageBox.critical(self, "错误", "异常数据处理失败，请检查输入文件！")
                    self.progress_bar.setValue(0)
            else:
                self.output_text.append("\n省份提取失败，终止流程！")
                QMessageBox.critical(self, "错误", "省份提取失败，请检查输入文件！")
                self.progress_bar.setValue(0)
        else:
            self.output_text.append("\n数据清洗失败，终止流程！")
            QMessageBox.critical(self, "错误", "数据清洗失败，请检查输入文件！")
            self.progress_bar.setValue(0)

        self.output_text.append(f"\n=== 处理完成 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===")
        self.status_label.setText("状态：处理完成")
        self.run_button.setEnabled(True)
        QMessageBox.information(self, "完成", "数据处理已完成，请检查输出目录！")

    def data_clean_1(self):
        """数据清洗：合并多个订单 Excel 文件，保留指定字段"""
        columns_to_keep = ["订单编号", "店铺", "仓库", "子单原始单号", "付款时间", "收货地区", "商家编码", "货品名称", "下单数量"]
        all_data = []
        total_records = 0

        for file in self.order_files:
            if os.path.exists(file):
                self.output_text.append(f"\n正在读取订单文件: {os.path.basename(file)}")
                # QApplication.processEvents()
                try:
                    df = pd.read_excel(file)
                    total_records += len(df)
                    self.output_text.append(f"文件包含 {len(df)} 条记录")
                    self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")
                    missing_columns = [col for col in columns_to_keep if col not in df.columns]
                    if missing_columns:
                        self.output_text.append(f"警告: 缺少字段 {missing_columns}")
                    else:
                        df_cleaned = df[columns_to_keep].copy()
                        all_data.append(df_cleaned)
                except Exception as e:
                    self.output_text.append(f"读取错误: {e}")
                    QMessageBox.warning(self, "警告", f"读取 {os.path.basename(file)} 失败: {e}")
            else:
                self.output_text.append(f"文件 {os.path.basename(file)} 不存在！")
                QMessageBox.warning(self, "警告", f"文件 {os.path.basename(file)} 不存在！")
            # QApplication.processEvents()

        if all_data:
            combined_data = pd.concat(all_data, ignore_index=True)
            self.output_text.append(f"\n合并完成，共 {len(combined_data)} 条记录")
            self.output_text.append(f"合并后的数据前 5 行：\n{combined_data.head().to_string()}")
            output_file = os.path.join(self.output_dir, "中间处理过程_cleaned_order_data.xlsx")
            try:
                combined_data.to_excel(output_file, index=False)
                self.output_text.append(f"\n清洗后的数据已保存到: {os.path.basename(output_file)}")
                return output_file
            except Exception as e:
                self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                return None
        else:
            self.output_text.append("\n没有成功读取任何订单数据！")
            return None

    def data_clean_2(self, input_file):
        """提取省份信息"""
        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            # QApplication.processEvents()
            try:
                df = pd.read_excel(input_file)
                self.output_text.append(f"文件包含 {len(df)} 条记录")
                self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")

                def extract_province(address):
                    if pd.isna(address):
                        return None
                    try:
                        return str(address).split()[0]
                    except IndexError:
                        return None

                df['省份'] = df['收货地区'].apply(extract_province)
                self.output_text.append(f"\n提取省份后的前 5 行数据：\n{df.head().to_string()}")
                self.output_text.append(f"省份字段缺失值统计：{df['省份'].isna().sum()} 条记录未提取到省份")

                output_file = os.path.join(self.output_dir, "中间处理过程_processed_order_data.xlsx")
                try:
                    df.to_excel(output_file, index=False)
                    self.output_text.append(f"\n处理后的数据已保存到: {os.path.basename(output_file)}")
                    return output_file
                except Exception as e:
                    self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                    QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                    return None
            except Exception as e:
                self.output_text.append(f"读取错误: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(input_file)} 失败: {e}")
                return None
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")
            return None

    def abnormal_process(self, input_file):
        """检测异常数据，添加月份，合并库存数据"""
        jinan_coverage = [
            '北京', '天津', '河北省', '山西省', '内蒙古自治区', '辽宁省', '吉林省', '黑龙江省',
            '上海', '江苏省', '浙江省', '安徽省', '山东省', '河南省', '湖北省', '北京市', '上海市', '天津市'
        ]

        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            # QApplication.processEvents()
            try:
                df = pd.read_excel(input_file)
                self.output_text.append(f"文件包含 {len(df)} 条记录")
                self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")

                # 添加月份字段
                df['付款时间'] = pd.to_datetime(df['付款时间'], errors='coerce')
                df['月份'] = df['付款时间'].dt.strftime('%Y-%m')
                self.output_text.append(f"\n添加月份字段后的前 5 行数据：\n{df.head().to_string()}")

                # 筛选指定仓库
                df = df[df['仓库'].isin(['佛山-优赛-三水仓', '济南-优赛-市中'])]
                self.output_text.append(f"\n筛选后（仅包含指定仓库）的记录数: {len(df)}")

                # 检测异常数据
                abnormal_data = []
                for index, row in df.iterrows():
                    province = row['省份']
                    warehouse = row['仓库']
                    if pd.isna(province):
                        abnormal_data.append(row)
                    elif warehouse == '佛山-优赛-三水仓' and province in jinan_coverage:
                        abnormal_data.append(row)
                    elif warehouse == '济南-优赛-市中' and province not in jinan_coverage:
                        abnormal_data.append(row)
                    # QApplication.processEvents()

                abnormal_df = pd.DataFrame(abnormal_data)
                if not abnormal_df.empty:
                    self.output_text.append(f"\n发现异常数据：\n{abnormal_df.head().to_string()}")
                    self.output_text.append(f"\n异常数据记录数: {len(abnormal_df)}")
                    output_file = os.path.join(self.output_dir, "中间处理过程_超区发货数据(不区分超区发货原因).xlsx")
                    try:
                        abnormal_df.to_excel(output_file, index=False)
                        self.output_text.append(f"\n异常数据已保存到: {os.path.basename(output_file)}")
                    except Exception as e:
                        self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                        QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                        return None
                else:
                    self.output_text.append("\n未发现异常数据！")
                    output_file = os.path.join(self.output_dir, "中间处理过程_超区发货数据(不区分超区发货原因).xlsx")
                    abnormal_df.to_excel(output_file, index=False)  # 保存空文件以便后续处理
                    self.output_text.append(f"\n无异常数据，保存空文件到: {os.path.basename(output_file)}")

                # 合并库存数据
                if os.path.exists(self.inventory_file):
                    self.output_text.append(f"\n正在读取库存数据: {os.path.basename(self.inventory_file)}")
                    try:
                        inventory_df = pd.read_excel(self.inventory_file)
                        self.output_text.append(f"库存数据前 5 行：\n{inventory_df.head().to_string()}")
                    except Exception as e:
                        self.output_text.append(f"读取库存数据错误: {e}")
                        QMessageBox.critical(self, "错误", f"读取 {os.path.basename(self.inventory_file)} 失败: {e}")
                        return None
                else:
                    self.output_text.append(f"\n库存文件 {os.path.basename(self.inventory_file)} 不存在！")
                    QMessageBox.critical(self, "错误", f"库存文件 {os.path.basename(self.inventory_file)} 不存在！")
                    return None

                # 确保商家编码列为字符串类型
                inventory_df['货品编号'] = inventory_df['货品编号'].astype(str)
                abnormal_df['商家编码'] = abnormal_df['商家编码'].astype(str)

                # 按仓库名称和货品编号透视库存数据
                pivot_inventory = inventory_df.pivot_table(
                    values=['期初库存', '期末库存'],
                    index='货品编号',
                    columns='仓库名称',
                    aggfunc='sum',
                    fill_value=0
                )

                # 重命名列名
                pivot_inventory.columns = [
                    '佛山仓期初库存' if '佛山-优赛-三水仓' in col and '期初库存' in col else
                    '佛山仓期末库存' if '佛山-优赛-三水仓' in col and '期末库存' in col else
                    '济南仓期初库存' if '济南-优赛-市中' in col and '期初库存' in col else
                    '济南仓期末库存' if '济南-优赛-市中' in col and '期末库存' in col else col
                    for col in pivot_inventory.columns
                ]
                pivot_inventory = pivot_inventory.reset_index()

                # 合并库存数据
                merged_df = abnormal_df.merge(
                    pivot_inventory,
                    left_on='商家编码',
                    right_on='货品编号',
                    how='left'
                )
                merged_df = merged_df.drop(columns=['货品编号'], errors='ignore')

                # 确保库存列存在并将NaN替换为0
                expected_columns = [
                    '佛山仓期初库存', '济南仓期初库存', '佛山仓期末库存', '济南仓期末库存'
                ]
                for col in expected_columns:
                    if col not in merged_df.columns:
                        merged_df[col] = 0
                    else:
                        merged_df[col] = merged_df[col].fillna(0)

                # 检查未匹配的商家编码
                missing_inventory = merged_df[
                    (merged_df['佛山仓期初库存'] == 0) &
                    (merged_df['济南仓期初库存'] == 0) &
                    (merged_df['佛山仓期末库存'] == 0) &
                    (merged_df['济南仓期末库存'] == 0)
                ]
                if not missing_inventory.empty:
                    self.output_text.append("\n警告：以下商家编码在库存数据中未找到对应的库存信息：")
                    self.output_text.append(missing_inventory[['商家编码', '货品名称']].to_string())

                # 保存合并后的数据
                inventory_output_file = os.path.join(self.output_dir, "中间处理过程_abnormal_order_data_with_inventory.xlsx")
                try:
                    merged_df.to_excel(inventory_output_file, index=False)
                    self.output_text.append(f"\n合并库存数据已保存到: {os.path.basename(inventory_output_file)}")
                    return inventory_output_file
                except Exception as e:
                    self.output_text.append(f"\n保存 {os.path.basename(inventory_output_file)} 错误: {e}")
                    QMessageBox.critical(self, "错误", f"保存 {os.path.basename(inventory_output_file)} 失败: {e}")
                    return None
            except Exception as e:
                self.output_text.append(f"读取错误: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(input_file)} 失败: {e}")
                return None
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")
            return None

    def filter_merchant_codes(self, input_file):
        """筛选商家编码"""
        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            # QApplication.processEvents()
            try:
                df = pd.read_excel(input_file)
                self.output_text.append(f"文件包含 {len(df)} 条记录")
                self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")

                # 筛选商家编码
                exclude_patterns = [
                    r'250g冰袋\*2\+500g干冰\*1',
                    r'250g冰袋\*4',
                    r'XDJXN',
                    r'XDJLW'
                ]
                df['商家编码'] = df['商家编码'].astype(str)
                mask = ~df['商家编码'].str.contains('|'.join(exclude_patterns), case=False, na=False, regex=True)
                cleaned_df = df[mask]

                # 检查被筛掉的记录
                excluded_df = df[~mask]
                if not excluded_df.empty:
                    self.output_text.append("\n被筛掉的记录（包含指定商家编码模式）：")
                    self.output_text.append(excluded_df[['订单编号', '商家编码', '货品名称']].to_string())

                # 保存清洗后的数据
                output_file = os.path.join(self.output_dir, "最终结果_缺货导致的超区发货数据.xlsx")
                try:
                    cleaned_df.to_excel(output_file, index=False)
                    self.output_text.append(f"\n清洗后的数据已保存到: {os.path.basename(output_file)}")
                except Exception as e:
                    self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                    QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
            except Exception as e:
                self.output_text.append(f"读取错误: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(input_file)} 失败: {e}")
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")