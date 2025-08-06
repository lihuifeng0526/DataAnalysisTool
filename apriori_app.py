# apriori_app.py
import pandas as pd
import os
from mlxtend.frequent_patterns import apriori, association_rules
from sklearn.preprocessing import MultiLabelBinarizer
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QPushButton,
                             QTextEdit, QFileDialog, QLabel, QMessageBox, QLineEdit, QFormLayout,QApplication)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, pyqtSignal
import io
import contextlib

class AprioriApp(QMainWindow):
    closed = pyqtSignal()  # 自定义信号，用于窗口关闭时通知

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_paths = []
        self.data_dir = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('商品关联性分析工具')
        self.setGeometry(100, 100, 800, 600)

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

        # 文件选择按钮
        self.select_button = QPushButton('选择 Excel 文件', self)
        self.select_button.clicked.connect(self.select_files)
        layout.addWidget(self.select_button)

        # 文件路径显示
        self.file_label = QLabel('未选择文件', self)
        layout.addWidget(self.file_label)

        # 输入支持度和置信度
        form_layout = QFormLayout()
        self.support_input = QLineEdit(self)
        self.support_input.setPlaceholderText('请输入支持度（0到1，例如 0.01）')
        self.confidence_input = QLineEdit(self)
        self.confidence_input.setPlaceholderText('请输入置信度（0到1，例如 0.6）')
        form_layout.addRow('最小支持度：', self.support_input)
        form_layout.addRow('最小置信度：', self.confidence_input)
        layout.addLayout(form_layout)

        # 运行按钮
        self.run_button = QPushButton('运行分析', self)
        self.run_button.clicked.connect(self.run_analysis)
        layout.addWidget(self.run_button)

        # 返回按钮
        self.back_button = QPushButton('返回主菜单', self)
        self.back_button.clicked.connect(self.close)
        layout.addWidget(self.back_button)

        # 日志输出窗口
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        # 状态栏
        self.statusBar().showMessage('就绪')

    def closeEvent(self, event):
        self.closed.emit()  # 发出关闭信号
        event.accept()


    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, '选择 Excel 文件', '', 'Excel Files (*.xlsx *.xls)')
        if files:
            self.file_paths = files
            # 设置 data_dir 为第一个文件的父目录下的 Data 文件夹
            self.data_dir = os.path.join(os.path.dirname(files[0]), 'Data')
            os.makedirs(self.data_dir, exist_ok=True)  # 自动创建 Data 文件夹
            self.file_label.setText(f'已选择 {len(files)} 个文件: {", ".join([os.path.basename(f) for f in files])}')
            self.statusBar().showMessage('文件已选择，点击“运行分析”开始处理')
        else:
            self.file_label.setText('未选择文件')
            self.data_dir = None
            self.statusBar().showMessage('就绪')

    def log(self, message):
        self.log_text.append(message)
        QApplication.processEvents()  # 实时更新 GUI

    def validate_input(self, value, param_name):
        """验证输入是否为 0 到 1 之间的浮点数"""
        try:
            val = float(value)
            if 0 < val <= 1:
                return val
            else:
                raise ValueError(f"{param_name} 必须在 0 到 1 之间")
        except ValueError:
            raise ValueError(f"请输入有效的 {param_name}（0到1之间的数字，例如 0.01）")

    def run_analysis(self):
        if not self.file_paths or not self.data_dir:
            QMessageBox.warning(self, '错误', '请先选择 Excel 文件！')
            return

        self.log_text.clear()
        self.statusBar().showMessage('正在运行分析...')
        self.run_button.setEnabled(False)

        # 重定向 print 到 GUI 日志窗口
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            try:
                self.data_clean()
                self.data_clean2()
                # 获取用户输入的支持度和置信度
                support_text = self.support_input.text().strip()
                confidence_text = self.confidence_input.text().strip()
                min_support = None
                min_confidence = 0.6  # 默认置信度
                if support_text:
                    min_support = self.validate_input(support_text, "支持度")
                if confidence_text:
                    min_confidence = self.validate_input(confidence_text, "置信度")
                self.run_apriori(min_support, min_confidence)
                self.statusBar().showMessage('分析完成！')
            except Exception as e:
                self.log(f"错误：{str(e)}")
                self.statusBar().showMessage('分析失败！')
        self.log(output.getvalue())
        self.run_button.setEnabled(True)

    def data_clean(self):
        required_columns = ['订单编号', '店铺', '客户编号', '商家编码', '货品名称']
        invalid_codes = ['250g冰袋*2+500g干冰*1', '250g冰袋*4', 'XDJXN', 'XDJLW']
        all_valid_data = []

        # 确保 data_dir 存在
        os.makedirs(self.data_dir, exist_ok=True)
        self.log("\n=== 1. 合并和初步清洗数据 ===")
        self.log(f"说明：读取 {len(self.file_paths)} 个 Excel 文件，过滤无效商家编码，合并数据。")

        for file in self.file_paths:
            try:
                df = pd.read_excel(file, usecols=required_columns)
                df_valid = df[~df['商家编码'].isin(invalid_codes)]
                all_valid_data.append(df_valid)
                self.log(f"\n=== 正在处理文件: {file} ===")
                self.log("表头字段:")
                self.log(str(list(df_valid.columns)))
                self.log("\n数据内容:")
                self.log(str(df_valid))
                self.log("\n" + "="*50)
            except FileNotFoundError:
                self.log(f"错误：文件 {file} 未找到，请检查文件路径！")
            except ValueError as e:
                self.log(f"错误：文件 {file} 中缺少部分指定字段: {str(e)}")
            except Exception as e:
                self.log(f"读取文件 {file} 时发生错误: {str(e)}")

        if all_valid_data:
            try:
                merged_df = pd.concat(all_valid_data, ignore_index=True)
                order_counts = merged_df['订单编号'].value_counts()
                valid_orders = order_counts[order_counts > 1].index
                merged_df = merged_df[merged_df['订单编号'].isin(valid_orders)]
                if merged_df.empty:
                    self.log("错误：过滤后没有包含多件商品的订单数据！")
                    return
                output_file = os.path.join(self.data_dir, 'merged_data.csv')
                merged_df.to_csv(output_file, index=False, encoding='utf-8-sig')
                self.log(f"\n=== 合并后的数据已保存到: {output_file} ===")
                self.log("合并数据表头:")
                self.log(str(list(merged_df.columns)))
                self.log("\n合并数据内容:")
                self.log(str(merged_df))
                self.log("\n" + "="*50)
            except Exception as e:
                self.log(f"合并或保存CSV文件时发生错误: {str(e)}")
        else:
            self.log("错误：没有有效数据可合并！")

    def data_clean2(self):
        required_columns = ['订单编号', '店铺', '客户编号', '商家编码', '货品名称']
        merged_file = os.path.join(self.data_dir, 'merged_data.csv')

        # 确保 data_dir 存在
        os.makedirs(self.data_dir, exist_ok=True)
        try:
            self.log("\n=== 2. 读取合并数据并去重 ===")
            self.log(f"说明：读取 {merged_file}，包含订单编号、商家编码等字段。")
            df_merged = pd.read_csv(merged_file, usecols=required_columns, encoding='utf-8-sig')
            self.log("\n数据表头（字段名）：")
            self.log(str(list(df_merged.columns)))
            self.log("\n数据内容（前几行）：")
            self.log(str(df_merged.head()))
            self.log(f"\n总行数：{len(df_merged)}")
            self.log("\n" + "="*50)

            self.log("\n=== 清洗数据：对每个订单的商家编码去重 ===")
            self.log("说明：在每个订单编号内，移除重复的商家编码，保留第一条记录的完整信息。")

            def deduplicate_items(group):
                return group.drop_duplicates(subset=['商家编码'], keep='first')

            df_cleaned = df_merged.groupby('订单编号').apply(deduplicate_items).reset_index(drop=True)
            order_item_counts = df_cleaned.groupby('订单编号')['商家编码'].nunique()
            valid_orders = order_item_counts[order_item_counts > 1].index
            df_cleaned = df_cleaned[df_cleaned['订单编号'].isin(valid_orders)]

            if df_cleaned.empty:
                self.log("错误：去重并过滤后没有包含多种商品的订单数据！")
                return

            cleaned_output_file = os.path.join(self.data_dir, 'cleaned_merged_data.csv')
            df_cleaned.to_csv(cleaned_output_file, index=False, encoding='utf-8-sig')
            self.log(f"\n=== 清洗后的数据已保存到: {cleaned_output_file} ===")
            self.log("说明：仅保留去重后包含多种商品（商家编码数量 > 1）的订单。")
            self.log("\n清洗后数据表头：")
            self.log(str(list(df_cleaned.columns)))
            self.log("\n清洗后数据内容：")
            self.log(str(df_cleaned))
            self.log(f"\n总行数（清洗后）：{len(df_cleaned)}")
            self.log("\n" + "="*50)
            self.log("\n=== 验证去重效果 ===")
            self.log("说明：显示每个订单编号的商品种类数（去重后的商家编码数量）。")
            self.log("订单编号与商品种类数：")
            self.log(str(order_item_counts[order_item_counts > 1]))
            self.log("\n" + "="*50)

        except Exception as e:
            self.log(f"合并或保存CSV文件时发生错误: {str(e)}")

    def run_apriori(self, min_support=None, min_confidence=0.6):
        required_columns = ['订单编号', '店铺', '客户编号', '商家编码', '货品名称']
        merged_file = os.path.join(self.data_dir, 'cleaned_merged_data.csv')

        # 确保 data_dir 存在
        os.makedirs(self.data_dir, exist_ok=True)
        try:
            self.log("\n=== 3. 读取清洗后的数据并运行 Apriori 算法 ===")
            self.log(f"说明：从 {merged_file} 读取数据，包含订单编号、商家编码等字段。")
            df_merged = pd.read_csv(merged_file, usecols=required_columns, encoding='utf-8-sig')

            # 检查数据是否为空
            if df_merged.empty:
                self.log("错误：数据文件为空，请检查 'cleaned_merged_data.csv' 是否包含有效数据！")
                return

            self.log("\n=== 4. 创建商家编码到货品名称的映射 ===")
            self.log("说明：从数据中提取商家编码和货品名称的对应关系，用于后续显示商品名称。")
            item_name_mapping = df_merged[['商家编码', '货品名称']].drop_duplicates(subset=['商家编码']).set_index('商家编码')['货品名称'].to_dict()
            self.log("商家编码到货品名称的映射（部分）：")
            for code, name in list(item_name_mapping.items())[:5]:
                self.log(f"商家编码: {code}, 货品名称: {name}")
            self.log(f"总映射数：{len(item_name_mapping)}")
            self.log("\n" + "="*50)

            self.log("\n=== 5. 数据概览 ===")
            self.log("字段名称：")
            self.log(str(list(df_merged.columns)))
            self.log("\n数据内容（前几行）：")
            self.log(str(df_merged.head()))
            self.log(f"\n总记录数：{len(df_merged)}")
            self.log(f"总订单数（唯一订单编号）：{df_merged['订单编号'].nunique()}")
            self.log(f"商品种类数（唯一商家编码）：{df_merged['商家编码'].nunique()}")
            self.log("\n" + "="*50)

            self.log("\n=== 6. 生成事务数据 ===")
            self.log("说明：将相同订单编号的记录视为一个事务，事务内容为该订单购买的所有商品（商家编码）。")
            self.log("注意：仅保留包含多个商品（商家编码数≥2）的事务，用于商品关联性分析。")
            transactions = df_merged.groupby('订单编号')['商家编码'].apply(list).reset_index()
            transactions = transactions[transactions['商家编码'].map(len) >= 2]
            if transactions.empty:
                self.log("错误：没有包含多个商品（商家编码数≥2）的事务，无法进行关联性分析！")
                return
            self.log("\n事务数据内容（每个订单的商品列表，仅包含多个商品的订单）：")
            for _, row in transactions.iterrows():
                self.log(f"订单编号: {row['订单编号']}, 商品（商家编码）: {row['商家编码']}")
            self.log(f"\n总事务数（订单数，仅包含多个商品的订单）：{len(transactions)}")
            self.log("\n" + "="*50)

            self.log("\n=== 7. 转换为 one-hot 编码 ===")
            self.log("说明：将事务数据转换为矩阵，每列为一个商品（商家编码），1表示订单包含该商品，0表示不包含。")
            mlb = MultiLabelBinarizer()
            one_hot_data = mlb.fit_transform(transactions['商家编码'])
            one_hot_df = pd.DataFrame(one_hot_data, columns=mlb.classes_)
            self.log("\none-hot 编码数据（前几行）：")
            self.log("说明：每行为一个订单，每列为一个商品（商家编码），值为1表示订单包含该商品，值为0表示不包含。")
            self.log(str(one_hot_df.head()))
            self.log(f"\n商品种类数（唯一商家编码）：{len(mlb.classes_)}")
            self.log("\n" + "="*50)

            self.log("\n=== 8. 生成频繁项集 ===")
            self.log("说明：频繁项集是支持度≥最小支持度的商品组合，包含所有项集大小（包括单商品项集）。")
            self.log("支持度=包含该商品组合的订单数/总订单数，表示订单占比。")
            # 使用用户输入的支持度，若未输入则使用默认值
            if min_support is None:
                min_support = max(1 / len(transactions), 0.01)
            self.log(f"最小支持度设置为：{min_support:.4f}")
            frequent_itemsets = apriori(one_hot_df, min_support=min_support, use_colnames=True)
            frequent_itemsets['项集大小'] = frequent_itemsets['itemsets'].apply(len)
            frequent_itemsets['商品名称'] = frequent_itemsets['itemsets'].apply(
                lambda x: tuple(item_name_mapping.get(item, f"未知商品({item})") for item in x)
            )
            frequent_itemsets['单量'] = (frequent_itemsets['support'] * len(transactions)).round().astype(int)

            self.log("\n频繁项集结果（包含所有项集大小）：")
            self.log("字段说明：")
            self.log("- support: 支持度（该商品组合出现的订单占比，值为包含该组合的订单数/总订单数）")
            self.log("- itemsets: 商品组合（商家编码集合）")
            self.log("- 项集大小: 商品组合中的商品数量")
            self.log("- 商品名称: 商品组合的货品名称")
            self.log("- 单量: 购买该商品组合的订单数（支持度×总订单数）")
            if not frequent_itemsets.empty:
                self.log(str(frequent_itemsets[['support', 'itemsets', '项集大小', '商品名称', '单量']].sort_values(by='support', ascending=False)))
                output_file = os.path.join(self.data_dir, 'frequent_itemsets.xlsx')
                frequent_itemsets.to_excel(output_file, index=False, engine='openpyxl')
                self.log(f"\n频繁项集已保存到：{output_file}")
            else:
                self.log(f"没有找到满足最小支持度（{min_support:.4f}）的频繁项集，请尝试降低最小支持度或检查数据！")
                self.log("建议：检查事务数据是否包含足够的多商品订单，或降低 min_support（例如 0.005）。")
            self.log("\n" + "="*50)

            self.log("\n=== 9. 生成关联规则 ===")
            self.log("说明：关联规则表示商品组合间的关联关系，例如X→Y表示购买X后可能购买Y。")
            self.log("置信度=包含X和Y的订单数/包含X的订单数，表示规则的可靠性。")
            self.log(f"最小置信度设置为：{min_confidence:.2f}")
            if not frequent_itemsets.empty:
                try:
                    rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=min_confidence)
                    # 添加商品名称列到规则
                    rules['前件商品名称'] = rules['antecedents'].apply(
                        lambda x: tuple(item_name_mapping.get(item, f"未知商品({item})") for item in x)
                    )
                    rules['后件商品名称'] = rules['consequents'].apply(
                        lambda x: tuple(item_name_mapping.get(item, f"未知商品({item})") for item in x)
                    )
                    # 添加单量列
                    rules['单量'] = (rules['support'] * len(transactions)).round().astype(int)
                    # 打印关联规则
                    self.log("\n关联规则结果：")
                    self.log("字段说明：")
                    self.log("- antecedents: 前件（规则的X部分，商家编码）")
                    self.log("- consequents: 后件（规则的Y部分，商家编码）")
                    self.log("- support: 支持度（规则出现的订单占比）")
                    self.log("- confidence: 置信度（规则的可靠性）")
                    self.log("- lift: 提升度（规则的强度，>1表示正相关）")
                    self.log("- 前件商品名称: 前件的货品名称")
                    self.log("- 后件商品名称: 后件的货品名称")
                    self.log("- 单量: 购买该规则组合的订单数（支持度×总订单数）")
                    if not rules.empty:
                        self.log(str(rules[['antecedents', 'consequents', 'support', 'confidence', 'lift', '前件商品名称', '后件商品名称', '单量']].sort_values(by='confidence', ascending=False)))
                        # 保存关联规则到Excel
                        rules_output_file = os.path.join(self.data_dir, '最终结果_association_rules.xlsx')
                        rules.to_excel(rules_output_file, index=False, engine='openpyxl')
                        self.log(f"\n关联规则已保存到：{rules_output_file}")
                    else:
                        self.log(f"没有找到满足最小置信度（{min_confidence:.2f}）的关联规则，请尝试降低 min_confidence（例如 0.5）或检查频繁项集！")
                        self.log("建议：检查频繁项集是否包含足够的多商品组合（项集大小≥2）。")
                except Exception as e:
                    self.log(f"生成关联规则时发生错误: {str(e)}")
                    self.log("建议：检查频繁项集是否为空或只包含单商品项集。尝试降低 min_support（例如 0.005）或 min_confidence（例如 0.5）。")
            else:
                self.log("无法生成关联规则，因为没有频繁项集。")
                self.log("建议：降低 min_support（例如 0.005）或检查数据清洗步骤，确保事务包含多种商品。")
            self.log("\n" + "="*50)

        except FileNotFoundError:
            self.log(f"错误：文件 {merged_file} 未找到，请确保 'Data/cleaned_merged_data.csv' 存在！")
        except ValueError as e:
            self.log(f"错误：{str(e)}")
        except Exception as e:
            self.log(f"处理文件 {merged_file} 时发生错误: {str(e)}")