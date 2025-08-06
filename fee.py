import sys
import os
import pandas as pd
import re
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, 
                             QTextEdit, QFileDialog, QLabel, QProgressBar, QMessageBox, QInputDialog)
from PyQt5.QtCore import Qt, pyqtSignal
import logging

# 设置日志记录
logging.basicConfig(filename='app.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def resource_path(relative_path):
    """获取打包后的资源路径"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)

class OrderDataProcessor(QMainWindow):
    closed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.order_files = []
        self.inventory_file = None
        self.shipping_file = None
        self.foshan_sheet = "佛山发货数据"
        self.jinan_sheet = "济南发货数据"
        self.output_dir = ""
        self.initUI()
        logging.debug("OrderDataProcessor initialized")

    def initUI(self):
        self.setWindowTitle("超区发货费用数据处理工具")
        self.setGeometry(100, 100, 600, 500)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.select_order_button = QPushButton("选择订单 Excel 文件", self)
        self.select_order_button.clicked.connect(self.select_order_files)
        layout.addWidget(self.select_order_button)

        self.select_inventory_button = QPushButton("选择库存 Excel 文件", self)
        self.select_inventory_button.clicked.connect(self.select_inventory_file)
        layout.addWidget(self.select_inventory_button)

        self.select_shipping_button = QPushButton("选择账单数据 Excel 文件", self)
        self.select_shipping_button.clicked.connect(self.select_shipping_file)
        layout.addWidget(self.select_shipping_button)

        self.run_button = QPushButton("运行处理", self)
        self.run_button.clicked.connect(self.run_processing)
        self.run_button.setEnabled(False)
        layout.addWidget(self.run_button)

        self.back_button = QPushButton("返回主菜单", self)
        self.back_button.clicked.connect(self.close)
        layout.addWidget(self.back_button)

        self.status_label = QLabel("状态：等待选择订单、库存和发货数据文件", self)
        layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.output_text = QTextEdit(self)
        self.output_text.setReadOnly(True)
        layout.addWidget(self.output_text)

    def closeEvent(self, event):
        self.closed.emit()
        event.accept()

    def select_order_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择订单数据 Excel 文件", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if files:
            self.order_files = files
            self.output_dir = os.path.dirname(files[0])
            self.output_text.append(f"已选择 {len(files)} 个订单文件：\n{chr(10).join([os.path.basename(f) for f in files])}")
            logging.debug(f"Selected order files: {files}")
            self.check_files_selected()
            QApplication.processEvents()

    def select_inventory_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "选择库存数据 Excel 文件", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file:
            self.inventory_file = file
            if not self.output_dir:
                self.output_dir = os.path.dirname(file)
            self.output_text.append(f"已选择库存文件：{os.path.basename(file)}")
            logging.debug(f"Selected inventory file: {file}")
            self.check_files_selected()
            QApplication.processEvents()

    def select_shipping_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "选择发货数据 Excel 文件", "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file:
            self.shipping_file = file
            if not self.output_dir:
                self.output_dir = os.path.dirname(file)
            self.output_text.append(f"已选择发货数据文件：{os.path.basename(file)}")
            logging.debug(f"Selected shipping file: {file}")
            
            try:
                xl = pd.ExcelFile(file)
                available_sheets = xl.sheet_names
                xl.close()
                logging.debug(f"Available sheets in shipping file: {available_sheets}")
                
                foshan_sheet, ok1 = QInputDialog.getText(
                    self, "输入 sheet 名称", 
                    f"请输入佛山发货数据的 sheet 名称（可用 sheet: {', '.join(available_sheets)}）:",
                    text=self.foshan_sheet
                )
                if ok1 and foshan_sheet:
                    self.foshan_sheet = foshan_sheet
                    self.output_text.append(f"佛山发货数据 sheet 名称：{foshan_sheet}")
                    logging.debug(f"Foshan sheet name: {foshan_sheet}")
                
                jinan_sheet, ok2 = QInputDialog.getText(
                    self, "输入 sheet 名称", 
                    f"请输入济南发货数据的 sheet 名称（可用 sheet: {', '.join(available_sheets)}）:",
                    text=self.jinan_sheet
                )
                if ok2 and jinan_sheet:
                    self.jinan_sheet = jinan_sheet
                    self.output_text.append(f"济南发货数据 sheet 名称：{jinan_sheet}")
                    logging.debug(f"Jinan sheet name: {jinan_sheet}")
                
                if not (ok1 and ok2):
                    self.output_text.append("未输入有效的 sheet 名称，取消选择！")
                    logging.warning("Invalid sheet names provided, resetting shipping file")
                    self.shipping_file = None
                    self.foshan_sheet = "佛山发货数据"
                    self.jinan_sheet = "济南发货数据"
                
            except Exception as e:
                self.output_text.append(f"读取 sheet 名称错误: {e}")
                logging.error(f"Failed to read sheet names from {file}: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(file)} 的 sheet 名称失败: {e}")
                self.shipping_file = None
                
            self.check_files_selected()
            QApplication.processEvents()

    def check_files_selected(self):
        if self.order_files and self.inventory_file and self.shipping_file and self.foshan_sheet and self.jinan_sheet:
            self.output_text.append(f"输出目录：{self.output_dir}")
            self.status_label.setText("状态：订单、库存和发货数据文件及 sheet 名称已选择，点击运行处理")
            self.run_button.setEnabled(True)
            self.progress_bar.setValue(0)
            logging.debug(f"Output directory: {self.output_dir}")
        else:
            self.status_label.setText("状态：请同时选择订单、库存和发货数据文件并输入 sheet 名称")
            self.run_button.setEnabled(False)

    def run_processing(self):
        self.run_button.setEnabled(False)
        self.status_label.setText("状态：正在处理...")
        self.progress_bar.setValue(0)
        self.output_text.append(f"\n=== 开始处理 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===")
        logging.debug("Starting data processing")

        total_steps = 6
        step_value = 100 // total_steps

        self.output_text.append("\n=== 步骤 1：数据清洗 ===")
        cleaned_file = self.data_clean_1()
        self.progress_bar.setValue(step_value)
        QApplication.processEvents()

        if cleaned_file:
            self.output_text.append("\n=== 步骤 2：提取省份 ===")
            processed_file = self.data_clean_2(cleaned_file)
            self.progress_bar.setValue(step_value * 2)
            QApplication.processEvents()

            if processed_file:
                self.output_text.append("\n=== 步骤 3：检测异常数据及库存合并 ===")
                abnormal_file = self.abnormal_process(processed_file)
                self.progress_bar.setValue(step_value * 3)
                QApplication.processEvents()

                if abnormal_file:
                    self.output_text.append("\n=== 步骤 4：筛选商家编码 ===")
                    cleaned_abnormal_file = self.filter_merchant_codes(abnormal_file)
                    self.progress_bar.setValue(step_value * 4)
                    QApplication.processEvents()

                    if cleaned_abnormal_file:
                        self.output_text.append("\n=== 步骤 5：追加佛山及济南发货数据 ===")
                        shipping_file = self.append_shipping_data(cleaned_abnormal_file)
                        self.progress_bar.setValue(step_value * 5)
                        QApplication.processEvents()

                        if shipping_file:
                            self.output_text.append("\n=== 步骤 6：订单编号去重及删除货品字段 ===")
                            final_file = self.process_final_shipping_data(shipping_file)
                            self.progress_bar.setValue(100)
                            QApplication.processEvents()

                            if not final_file:
                                self.output_text.append("\n去重及删除货品字段失败，终止流程！")
                                logging.error("Failed to process final shipping data")
                                QMessageBox.critical(self, "错误", "去重及删除货品字段失败，请检查输出文件！")
                                self.progress_bar.setValue(0)
                        else:
                            self.output_text.append("\n追加发货数据失败，终止流程！")
                            logging.error("Failed to append shipping data")
                            QMessageBox.critical(self, "错误", "追加发货数据失败，请检查输入文件！")
                            self.progress_bar.setValue(0)
                    else:
                        self.output_text.append("\n筛选商家编码失败，终止流程！")
                        logging.error("Failed to filter merchant codes")
                        QMessageBox.critical(self, "错误", "筛选商家编码失败，请检查输入文件！")
                        self.progress_bar.setValue(0)
                else:
                    self.output_text.append("\n异常数据处理失败，终止流程！")
                    logging.error("Failed to process abnormal data")
                    QMessageBox.critical(self, "错误", "异常数据处理失败，请检查输入文件！")
                    self.progress_bar.setValue(0)
            else:
                self.output_text.append("\n省份提取失败，终止流程！")
                logging.error("Failed to extract provinces")
                QMessageBox.critical(self, "错误", "省份提取失败，请检查输入文件！")
                self.progress_bar.setValue(0)
        else:
            self.output_text.append("\n数据清洗失败，终止流程！")
            logging.error("Failed to clean data")
            QMessageBox.critical(self, "错误", "数据清洗失败，请检查输入文件！")
            self.progress_bar.setValue(0)

        self.output_text.append(f"\n=== 处理完成 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===")
        self.status_label.setText("状态：处理完成")
        self.run_button.setEnabled(True)
        QMessageBox.information(self, "完成", "数据处理已完成，请检查输出目录！")
        logging.debug("Data processing completed")

    def data_clean_1(self):
        columns_to_keep = ["订单编号", "店铺", "仓库", "子单原始单号", "付款时间", "收货地区", "商家编码", "货品名称", "下单数量", "物流单号", "拆自组合装"]
        all_data = []
        total_records = 0

        for file in self.order_files:
            file_path = file  # 用户选择的文件无需 resource_path
            if os.path.exists(file_path):
                self.output_text.append(f"\n正在读取订单文件: {os.path.basename(file_path)}")
                logging.debug(f"Reading order file: {file_path}")
                try:
                    df = pd.read_excel(file_path)
                    total_records += len(df)
                    self.output_text.append(f"文件包含 {len(df)} 条记录")
                    self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")
                    logging.debug(f"Order file {file_path} contains {len(df)} records")
                    missing_columns = [col for col in columns_to_keep if col not in df.columns]
                    if missing_columns:
                        self.output_text.append(f"警告: 缺少字段 {missing_columns}")
                        logging.warning(f"Missing columns in {file_path}: {missing_columns}")
                    else:
                        df_cleaned = df[columns_to_keep].copy()
                        all_data.append(df_cleaned)
                except Exception as e:
                    self.output_text.append(f"读取错误: {e}")
                    logging.error(f"Failed to read order file {file_path}: {e}")
                    QMessageBox.warning(self, "警告", f"读取 {os.path.basename(file_path)} 失败: {e}")
            else:
                self.output_text.append(f"文件 {os.path.basename(file_path)} 不存在！")
                logging.error(f"Order file not found: {file_path}")
                QMessageBox.warning(self, "警告", f"文件 {os.path.basename(file_path)} 不存在！")
            QApplication.processEvents()

        if all_data:
            combined_data = pd.concat(all_data, ignore_index=True)
            self.output_text.append(f"\n合并完成，共 {len(combined_data)} 条记录")
            self.output_text.append(f"合并后的数据前 5 行：\n{combined_data.head().to_string()}")
            logging.debug(f"Combined {len(combined_data)} records")

            initial_count = len(combined_data)
            combined_data = combined_data[combined_data['物流单号'].notna() & (combined_data['物流单号'] != '')]
            filtered_count = len(combined_data)
            self.output_text.append(f"\n剔除物流单号为空的记录后，剩余 {filtered_count} 条记录（原 {initial_count} 条，剔除了 {initial_count - filtered_count} 条）")
            logging.debug(f"Filtered to {filtered_count} records (from {initial_count})")

            output_file = os.path.join(self.output_dir, "中间过程处理_合并订单数据.xlsx")
            try:
                combined_data.to_excel(output_file, engine='openpyxl')
                self.output_text.append(f"\n清洗后的数据已保存到: {os.path.basename(output_file)}")
                logging.debug(f"Saved cleaned data to: {output_file}")
                return output_file
            except Exception as e:
                self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                logging.error(f"Failed to save {output_file}: {e}")
                QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                return None
        else:
            self.output_text.append("\n没有成功读取任何订单数据！")
            logging.error("No order data read successfully")
            return None

    def data_clean_2(self, input_file):
        input_file = resource_path(input_file)
        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            logging.debug(f"Reading file: {input_file}")
            try:
                df = pd.read_excel(input_file)
                self.output_text.append(f"文件包含 {len(df)} 条记录")
                self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")
                logging.debug(f"File {input_file} contains {len(df)} records")

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
                logging.debug(f"Province extraction completed, missing provinces: {df['省份'].isna().sum()}")

                output_file = os.path.join(self.output_dir, "中间过程处理_添加省份字段.xlsx")
                try:
                    df.to_excel(output_file, engine='openpyxl')
                    self.output_text.append(f"\n处理后的数据已保存到: {os.path.basename(output_file)}")
                    logging.debug(f"Saved province data to: {output_file}")
                    return output_file
                except Exception as e:
                    self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                    logging.error(f"Failed to save {output_file}: {e}")
                    QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                    return None
            except Exception as e:
                self.output_text.append(f"读取错误: {e}")
                logging.error(f"Failed to read {input_file}: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(input_file)} 失败: {e}")
                return None
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            logging.error(f"Input file not found: {input_file}")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")
            return None

    def abnormal_process(self, input_file):
        jinan_coverage = [
            '北京', '天津', '河北省', '山西省', '内蒙古自治区', '辽宁省', '吉林省', '黑龙江省',
            '上海', '江苏省', '浙江省', '安徽省', '山东省', '河南省', '湖北省', '北京市', '上海市', '天津市'
        ]
        input_file = resource_path(input_file)
        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            logging.debug(f"Reading file: {input_file}")
            try:
                df = pd.read_excel(input_file)
                self.output_text.append(f"文件包含 {len(df)} 条记录")
                self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")
                logging.debug(f"File {input_file} contains {len(df)} records")

                df['付款时间'] = pd.to_datetime(df['付款时间'], errors='coerce')
                df['月份'] = df['付款时间'].dt.strftime('%Y-%m')
                self.output_text.append(f"\n添加月份字段后的前 5 行数据：\n{df.head().to_string()}")
                logging.debug("Added month column")

                df = df[df['仓库'].isin(['佛山-优赛-三水仓', '济南-优赛-市中'])]
                self.output_text.append(f"\n筛选后（仅包含指定仓库）的记录数: {len(df)}")
                logging.debug(f"Filtered to {len(df)} records with specified warehouses")

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
                    QApplication.processEvents()

                abnormal_df = pd.DataFrame(abnormal_data)
                if not abnormal_df.empty:
                    self.output_text.append(f"\n发现异常数据：\n{abnormal_df.head().to_string()}")
                    self.output_text.append(f"\n异常数据记录数: {len(abnormal_df)}")
                    logging.debug(f"Found {len(abnormal_df)} abnormal records")
                    output_file = os.path.join(self.output_dir, "中间过程处理_异常数据.xlsx")
                    try:
                        abnormal_df.to_excel(output_file, engine='openpyxl')
                        self.output_text.append(f"\n异常数据已保存到: {os.path.basename(output_file)}")
                        logging.debug(f"Saved abnormal data to: {output_file}")
                    except Exception as e:
                        self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                        logging.error(f"Failed to save {output_file}: {e}")
                        QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                        return None
                else:
                    self.output_text.append("\n未发现异常数据！")
                    output_file = os.path.join(self.output_dir, "中间过程处理_异常数据.xlsx")
                    try:
                        abnormal_df.to_excel(output_file, engine='openpyxl')
                        self.output_text.append(f"\n无异常数据，保存空文件到: {os.path.basename(output_file)}")
                        logging.debug(f"Saved empty abnormal data to: {output_file}")
                    except Exception as e:
                        self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                        logging.error(f"Failed to save {output_file}: {e}")
                        QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                        return None

                if os.path.exists(self.inventory_file):
                    self.output_text.append(f"\n正在读取库存数据: {os.path.basename(self.inventory_file)}")
                    logging.debug(f"Reading inventory file: {self.inventory_file}")
                    try:
                        inventory_df = pd.read_excel(self.inventory_file)
                        self.output_text.append(f"库存数据前 5 行：\n{inventory_df.head().to_string()}")
                        logging.debug(f"Inventory data head: {inventory_df.head().to_string()}")
                        logging.debug(f"Inventory columns: {inventory_df.columns.tolist()}")
                    except Exception as e:
                        self.output_text.append(f"读取库存数据错误: {e}")
                        logging.error(f"Failed to read inventory file {self.inventory_file}: {e}")
                        QMessageBox.critical(self, "错误", f"读取 {os.path.basename(self.inventory_file)} 失败: {e}")
                        return None
                else:
                    self.output_text.append(f"\n库存文件 {os.path.basename(self.inventory_file)} 不存在！")
                    logging.error(f"Inventory file not found: {self.inventory_file}")
                    QMessageBox.critical(self, "错误", f"库存文件 {os.path.basename(self.inventory_file)} 不存在！")
                    return None

                inventory_df['货品编号'] = inventory_df['货品编号'].astype(str)
                abnormal_df['商家编码'] = abnormal_df['商家编码'].astype(str)

                pivot_inventory = inventory_df.pivot_table(
                    values=['期初库存', '期末库存'],
                    index='货品编号',
                    columns='仓库名称',
                    aggfunc='sum',
                    fill_value=0
                )

                pivot_inventory.columns = [
                    '佛山仓期初库存' if '佛山-优赛-三水仓' in col and '期初库存' in col else
                    '佛山仓期末库存' if '佛山-优赛-三水仓' in col and '期末库存' in col else
                    '济南仓期初库存' if '济南-优赛-市中' in col and '期初库存' in col else
                    '济南仓期末库存' if '济南-优赛-市中' in col and '期末库存' in col else col
                    for col in pivot_inventory.columns
                ]
                pivot_inventory = pivot_inventory.reset_index()

                merged_df = abnormal_df.merge(
                    pivot_inventory,
                    left_on='商家编码',
                    right_on='货品编号',
                    how='left'
                )
                merged_df = merged_df.drop(columns=['货品编号'], errors='ignore')

                expected_columns = [
                    '佛山仓期初库存', '济南仓期初库存', '佛山仓期末库存', '济南仓期末库存'
                ]
                for col in expected_columns:
                    if col not in merged_df.columns:
                        merged_df[col] = 0
                    else:
                        merged_df[col] = merged_df[col].fillna(0)

                missing_inventory = merged_df[
                    (merged_df['佛山仓期初库存'] == 0) &
                    (merged_df['济南仓期初库存'] == 0) &
                    (merged_df['佛山仓期末库存'] == 0) &
                    (merged_df['济南仓期末库存'] == 0)
                ]
                if not missing_inventory.empty:
                    self.output_text.append("\n警告：以下商家编码在库存数据中未找到对应的库存信息：")
                    self.output_text.append(missing_inventory[['商家编码', '货品名称']].to_string())
                    logging.debug(f"Missing inventory data: {missing_inventory[['商家编码', '货品名称']].to_string()}")

                inventory_output_file = os.path.join(self.output_dir, "中间过程处理_合并库存数据.xlsx")
                try:
                    merged_df.to_excel(inventory_output_file, engine='openpyxl')
                    self.output_text.append(f"\n合并库存数据已保存到: {os.path.basename(inventory_output_file)}")
                    logging.debug(f"Saved merged inventory data to: {inventory_output_file}")
                    return inventory_output_file
                except Exception as e:
                    self.output_text.append(f"\n保存 {os.path.basename(inventory_output_file)} 错误: {e}")
                    logging.error(f"Failed to save {inventory_output_file}: {e}")
                    QMessageBox.critical(self, "错误", f"保存 {os.path.basename(inventory_output_file)} 失败: {e}")
                    return None
            except Exception as e:
                self.output_text.append(f"读取错误: {e}")
                logging.error(f"Failed to read {input_file}: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(input_file)} 失败: {e}")
                return None
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            logging.error(f"Input file not found: {input_file}")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")
            return None

    def filter_merchant_codes(self, input_file):
        input_file = resource_path(input_file)
        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            logging.debug(f"Reading file: {input_file}")
            try:
                df = pd.read_excel(input_file)
                self.output_text.append(f"文件包含 {len(df)} 条记录")
                self.output_text.append(f"前 5 行数据：\n{df.head().to_string()}")
                logging.debug(f"File {input_file} contains {len(df)} records")

                exclude_patterns = [
                    r'250g冰袋\*2\+500g干冰\*1',
                    r'250g冰袋\*4',
                    r'XDJXN',
                    r'XDJLW'
                ]
                df['商家编码'] = df['商家编码'].astype(str)
                mask = ~df['商家编码'].str.contains('|'.join(exclude_patterns), case=False, na=False, regex=True)
                cleaned_df = df[mask]

                excluded_df = df[~mask]
                if not excluded_df.empty:
                    self.output_text.append("\n被筛掉的记录（包含指定商家编码模式）：")
                    self.output_text.append(excluded_df[['订单编号', '商家编码', '货品名称']].to_string())
                    logging.debug(f"Excluded records: {excluded_df[['订单编号', '商家编码', '货品名称']].to_string()}")

                output_file = os.path.join(self.output_dir, "中间过程处理_筛选商家编码.xlsx")
                try:
                    cleaned_df.to_excel(output_file, engine='openpyxl')
                    self.output_text.append(f"\n清洗后的数据已保存到: {os.path.basename(output_file)}")
                    logging.debug(f"Saved filtered data to: {output_file}")
                    return output_file
                except Exception as e:
                    self.output_text.append(f"\n保存 {os.path.basename(output_file)} 错误: {e}")
                    logging.error(f"Failed to save {output_file}: {e}")
                    QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                    return None
            except Exception as e:
                self.output_text.append(f"读取错误: {e}")
                logging.error(f"Failed to read {input_file}: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(input_file)} 失败: {e}")
                return None
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            logging.error(f"Input file not found: {input_file}")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")
            return None

    def append_shipping_data(self, input_file):
        input_file = resource_path(input_file)
        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            logging.debug(f"Reading file: {input_file}")
            try:
                abnormal_df = pd.read_excel(input_file)
                self.output_text.append(f"文件包含 {len(abnormal_df)} 条记录")
                self.output_text.append(f"前 5 行数据：\n{abnormal_df.head().to_string()}")
                logging.debug(f"File {input_file} contains {len(abnormal_df)} records")

                abnormal_df['子单原始单号'] = abnormal_df['子单原始单号'].astype(str)

                if os.path.exists(self.shipping_file):
                    self.output_text.append(f"\n正在读取发货数据文件: {os.path.basename(self.shipping_file)}")
                    logging.debug(f"Reading shipping file: {self.shipping_file}")
                    output_file = os.path.join(self.output_dir, "中间过程处理_追加发货数据字段.xlsx")

                    try:
                        writer = pd.ExcelWriter(output_file, engine='openpyxl')
                        if writer is None:
                            raise ValueError("Failed to initialize ExcelWriter with openpyxl engine")
                        
                        self.output_text.append(f"\n正在读取佛山发货数据（Sheet: {self.foshan_sheet}）")
                        logging.debug(f"Reading Foshan shipping data (Sheet: {self.foshan_sheet})")
                        foshan_df = pd.read_excel(self.shipping_file, sheet_name=self.foshan_sheet)
                        self.output_text.append(f"佛山发货数据前 5 行：\n{foshan_df.head().to_string()}")
                        logging.debug(f"Foshan data head: {foshan_df.head().to_string()}")

                        foshan_df['原始单号'] = foshan_df['原始单号'].astype(str)
                        foshan_columns = [col for col in foshan_df.columns if col != '原始单号']
                        foshan_columns_renamed = [f"佛山_{col}" for col in foshan_columns]
                        foshan_df = foshan_df.rename(columns=dict(zip(foshan_columns, foshan_columns_renamed)))

                        foshan_merged_df = abnormal_df.merge(
                            foshan_df[foshan_columns_renamed + ['原始单号']],
                            left_on='子单原始单号',
                            right_on='原始单号',
                            how='left'
                        )
                        foshan_merged_df = foshan_merged_df.drop(columns=['原始单号'], errors='ignore')

                        initial_count = len(foshan_merged_df)
                        foshan_merged_df = foshan_merged_df[
                            ~foshan_merged_df['仓库'].str.contains('济南', case=False, na=False)
                        ]
                        filtered_count = len(foshan_merged_df)
                        self.output_text.append(
                            f"\n佛山发货数据筛选后，剩余 {filtered_count} 条记录（原 {initial_count} 条，剔除了 {initial_count - filtered_count} 条包含‘济南’的记录）"
                        )
                        logging.debug(f"Foshan filtered: {filtered_count} records (from {initial_count})")

                        unmatched_foshan = foshan_merged_df[foshan_merged_df[foshan_columns_renamed].isna().all(axis=1)]
                        if not unmatched_foshan.empty:
                            self.output_text.append("\n警告：以下子单原始单号未在佛山发货数据中找到匹配：")
                            self.output_text.append(unmatched_foshan[['子单原始单号', '商家编码', '货品名称']].to_string())
                            logging.debug(f"Unmatched Foshan orders: {unmatched_foshan[['子单原始单号', '商家编码', '货品名称']].to_string()}")

                        foshan_merged_df.to_excel(writer, sheet_name='佛山发货数据', index=False)
                        self.output_text.append(f"\n佛山发货数据合并结果已保存到: {os.path.basename(output_file)}（Sheet: 佛山发货数据）")
                        logging.debug(f"Saved Foshan data to {output_file} (Sheet: 佛山发货数据)")

                        self.output_text.append(f"\n正在读取济南发货数据（Sheet: {self.jinan_sheet})")
                        logging.debug(f"Reading Jinan shipping data (Sheet: {self.jinan_sheet})")
                        jinan_df = pd.read_excel(self.shipping_file, sheet_name=self.jinan_sheet)
                        self.output_text.append(f"济南发货数据前 5 行：\n{jinan_df.head().to_string()}")
                        logging.debug(f"Jinan data head: {jinan_df.head().to_string()}")

                        jinan_df['原始单号'] = jinan_df['原始单号'].astype(str)
                        jinan_columns = [col for col in jinan_df.columns if col != '原始单号']
                        jinan_columns_renamed = [f"济南_{col}" for col in jinan_columns]
                        jinan_df = jinan_df.rename(columns=dict(zip(jinan_columns, jinan_columns_renamed)))

                        jinan_merged_df = abnormal_df.merge(
                            jinan_df[jinan_columns_renamed + ['原始单号']],
                            left_on='子单原始单号',
                            right_on='原始单号',
                            how='left'
                        )
                        jinan_merged_df = jinan_merged_df.drop(columns=['原始单号'], errors='ignore')

                        initial_count = len(jinan_merged_df)
                        jinan_merged_df = jinan_merged_df[
                            jinan_merged_df['仓库'].str.contains('济南', case=False, na=False)
                        ]
                        filtered_count = len(jinan_merged_df)
                        self.output_text.append(
                            f"\n济南发货数据筛选后，剩余 {filtered_count} 条记录（原 {initial_count} 条，剔除了 {initial_count - filtered_count} 条不包含‘济南’的记录）"
                        )
                        logging.debug(f"Jinan filtered: {filtered_count} records (from {initial_count})")

                        unmatched_jinan = jinan_merged_df[jinan_merged_df[jinan_columns_renamed].isna().all(axis=1)]
                        if not unmatched_jinan.empty:
                            self.output_text.append("\n警告：以下子单原始单号未在济南发货数据中找到匹配：")
                            self.output_text.append(unmatched_jinan[['子单原始单号', '商家编码', '货品名称']].to_string())
                            logging.debug(f"Unmatched Jinan orders: {unmatched_jinan[['子单原始单号', '商家编码', '货品名称']].to_string()}")

                        jinan_merged_df.to_excel(writer, sheet_name='济南发货数据', index=False)
                        self.output_text.append(f"\n济南发货数据合并结果已保存到: {os.path.basename(output_file)}（Sheet: 济南发货数据）")
                        logging.debug(f"Saved Jinan data to {output_file} (Sheet: 济南发货数据)")

                        writer.close()
                        self.output_text.append(f"\n追加发货数据后的结果已保存到: {os.path.basename(output_file)}")
                        logging.debug(f"Saved shipping data to: {output_file}")
                        return output_file

                    except Exception as e:
                        self.output_text.append(f"保存发货数据错误: {e}")
                        logging.error(f"Failed to save shipping data to {output_file}: {e}")
                        QMessageBox.critical(self, "错误", f"保存发货数据失败: {e}")
                        return None

                else:
                    self.output_text.append(f"\n发货数据文件 {os.path.basename(self.shipping_file)} 不存在！")
                    logging.error(f"Shipping file not found: {self.shipping_file}")
                    QMessageBox.critical(self, "错误", f"发货数据文件 {os.path.basename(self.shipping_file)} 不存在！")
                    return None

            except Exception as e:
                self.output_text.append(f"读取错误: {e}")
                logging.error(f"Failed to read {input_file}: {e}")
                QMessageBox.critical(self, "错误", f"读取 {os.path.basename(input_file)} 失败: {e}")
                return None
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            logging.error(f"Input file not found: {input_file}")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")
            return None

    def process_final_shipping_data(self, input_file):
        input_file = resource_path(input_file)
        if os.path.exists(input_file):
            self.output_text.append(f"\n正在读取文件: {os.path.basename(input_file)}")
            logging.debug(f"Reading file: {input_file}")
            try:
                xl = pd.ExcelFile(input_file)
                sheet_names = xl.sheet_names
                self.output_text.append(f"\n发现的 sheet 名称: {sheet_names}")
                logging.debug(f"Found sheets: {sheet_names}")
                
                output_file = os.path.join(self.output_dir, "最终结果_超区发货费用数据表.xlsx")
                try:
                    writer = pd.ExcelWriter(output_file, engine='openpyxl')
                    if writer is None:
                        raise ValueError("Failed to initialize ExcelWriter with openpyxl engine")
                    
                    for sheet in sheet_names:
                        self.output_text.append(f"\n=== 处理 Sheet: {sheet} ===")
                        logging.debug(f"Processing sheet: {sheet}")
                        
                        df = pd.read_excel(input_file, sheet_name=sheet)
                        self.output_text.append(f"{sheet} 原始记录数: {len(df)}")
                        self.output_text.append(f"前 5 行数据:\n{df.head().to_string()}")
                        logging.debug(f"{sheet} has {len(df)} records")
                        
                        initial_count = len(df)
                        df = df.drop_duplicates(subset=['订单编号'], keep='first')
                        self.output_text.append(f"\n去重后记录数: {len(df)}（原 {initial_count} 条，剔除了 {initial_count - len(df)} 条重复记录）")
                        logging.debug(f"Deduplicated {sheet}: {len(df)} records (from {initial_count})")
                        
                        columns_to_drop = [col for col in df.columns if '货品' in str(col) or '商家编码' in str(col)]
                        if columns_to_drop:
                            self.output_text.append(f"\n将删除的包含‘货品’或‘商家编码’的列: {columns_to_drop}")
                            logging.debug(f"Dropping columns: {columns_to_drop}")
                            df = df.drop(columns=columns_to_drop)
                        else:
                            self.output_text.append("\n未找到包含‘货品’或‘商家编码’的列")
                            logging.debug("No columns with '货品' or '商家编码' found")
                        
                        self.output_text.append(f"\n处理后的前 5 行数据:\n{df.head().to_string()}")
                        logging.debug(f"Processed {sheet} head: {df.head().to_string()}")
                        
                        df.to_excel(writer, sheet_name=sheet, index=False)
                        self.output_text.append(f"\n{sheet} 处理结果已保存到: {os.path.basename(output_file)}（Sheet: {sheet}）")
                        logging.debug(f"Saved {sheet} to {output_file}")
                    
                    writer.close()
                    self.output_text.append(f"\n去重及删除货品字段后的结果已保存到: {os.path.basename(output_file)}")
                    logging.debug(f"Saved final shipping data to: {output_file}")
                    return output_file
                
                except Exception as e:
                    self.output_text.append(f"\n保存最终结果错误: {e}")
                    logging.error(f"Failed to save final shipping data to {output_file}: {e}")
                    QMessageBox.critical(self, "错误", f"保存 {os.path.basename(output_file)} 失败: {e}")
                    return None

            except Exception as e:
                self.output_text.append(f"\n处理错误: {e}")
                logging.error(f"Failed to process {input_file}: {e}")
                QMessageBox.critical(self, "错误", f"处理 {os.path.basename(input_file)} 失败: {e}")
                return None
        else:
            self.output_text.append(f"\n文件 {os.path.basename(input_file)} 不存在！")
            logging.error(f"Input file not found: {input_file}")
            QMessageBox.critical(self, "错误", f"文件 {os.path.basename(input_file)} 不存在！")
            return None

def main():
    try:
        app = QApplication(sys.argv)
        window = OrderDataProcessor()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.error(f"Main function error: {e}")
        raise

if __name__ == "__main__":
    main()