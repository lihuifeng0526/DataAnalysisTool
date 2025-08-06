# DataAnalysisTool

**English** | [中文](#中文)

## English

### Overview
`DataAnalysisTool` is a standalone Windows application designed for processing and analyzing e-commerce data, including cross-regional shipping cost data, abnormal order detection, and product association analysis. Built with Python and PyQt5, it is packaged as a single executable file using PyInstaller, requiring no Python environment or additional dependencies.

### Features
- **Cross-Regional Shipping Cost Analysis**: Processes order, inventory, and shipping data to generate cost reports, supporting Excel file input and output.
- **Abnormal Order Detection**: Identifies anomalies in order data based on warehouse and province matching.
- **Product Association Analysis**: Uses the Apriori algorithm to analyze product purchase patterns.
- **User-Friendly GUI**: Intuitive interface with progress bars and detailed logs for data processing.
- **Single Executable**: No installation required, runs directly on Windows 10/11 (64-bit).

### Installation
1. Download the latest release (`DataAnalysisTool.exe`) from the [GitHub Releases](https://github.com/your-username/DataAnalysisTool/releases) page.
2. Ensure your system has [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe) installed (required for some dependencies).
3. Place the `.exe` file in a folder with write permissions (e.g., `C:\Users\YourName\Documents`).
4. Double-click `DataAnalysisTool.exe` to launch the application.

### Usage
1. **Launch the Application**: Run `DataAnalysisTool.exe`. The GUI will display three main modules:
   - Cross-Regional Shipping Cost Data
   - Abnormal Order Data
   - Product Association Analysis
2. **Select Module**: Click on a module (e.g., "Cross-Regional Shipping Cost Data").
3. **Input Files**:
   - Select order data Excel files (`.xlsx` or `.xls`).
   - Choose an inventory data Excel file (e.g., `6.7-6.14库存数据.xlsx`).
   - Select a shipping data Excel file and specify sheet names for Foshan and Jinan data.
4. **Run Processing**: Click "Run Processing" to analyze data. Check the output directory for results (e.g., `最终结果_超区发货费用数据表.xlsx`).
5. **View Logs**: Errors and processing details are logged in `app.log` in the same directory as the `.exe`.

### System Requirements
- **Operating System**: Windows 10/11 (64-bit)
- **Disk Space**: ~200 MB for the executable and output files
- **Dependencies**: None (all included in the `.exe`)
- **Input Files**: Excel files (`.xlsx` or `.xls`) with specific column formats (refer to documentation for details)

### Contributing
Contributions are welcome! To contribute:
1. Fork the repository.
2. Create a new branch (`git checkout -b feature/your-feature`).
3. Commit your changes (`git commit -m "Add your feature"`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Open a Pull Request.

Please report issues or suggest features via the [Issues](https://github.com/your-username/DataAnalysisTool/issues) page.

### License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## 中文

### 概述
`DataAnalysisTool` 是一款独立的 Windows 应用程序，用于处理和分析电商数据，包括超区发货费用数据、异常订单检测和商品关联性分析。该工具使用 Python 和 PyQt5 开发，通过 PyInstaller 打包为单一可执行文件，无需 Python 环境或额外依赖。

### 功能
- **超区发货费用分析**：处理订单、库存和发货数据，生成费用报表，支持 Excel 文件输入和输出。
- **异常订单检测**：根据仓库和省份匹配规则识别订单数据中的异常。
- **商品关联性分析**：使用 Apriori 算法分析商品购买模式。
- **用户友好的界面**：直观的图形界面，包含进度条和详细的处理日志。
- **单一可执行文件**：无需安装，直接在 Windows 10/11（64 位）上运行。

### 安装
1. 从 [GitHub Releases](https://github.com/your-username/DataAnalysisTool/releases) 页面下载最新版本的 `DataAnalysisTool.exe`。
2. 确保系统已安装 [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)（部分依赖需要）。
3. 将 `.exe` 文件放置在具有写入权限的文件夹（例如 `C:\Users\YourName\Documents`）。
4. 双击 `DataAnalysisTool.exe` 启动应用程序。

### 使用方法
1. **启动应用程序**：运行 `DataAnalysisTool.exe`，界面将显示三个主要模块：
   - 超区发货费用数据
   - 异常订单数据
   - 商品关联性分析
2. **选择模块**：点击所需模块（例如“超区发货费用数据”）。
3. **输入文件**：
   - 选择订单数据的 Excel 文件（`.xlsx` 或 `.xls`）。
   - 选择库存数据的 Excel 文件（例如 `6.7-6.14库存数据.xlsx`）。
   - 选择发货数据的 Excel 文件，并指定佛山和济南数据的 sheet 名称。
4. **运行处理**：点击“运行处理”分析数据。结果将保存在输出目录（例如 `最终结果_超区发货费用数据表.xlsx`）。
5. **查看日志**：错误和处理详情记录在 `.exe` 所在目录的 `app.log` 文件中。

### 系统要求
- **操作系统**：Windows 10/11（64 位）
- **磁盘空间**：约 200 MB 用于可执行文件和输出文件
- **依赖项**：无（所有依赖已包含在 `.exe` 中）
- **输入文件**：Excel 文件（`.xlsx` 或 `.xls`），需符合特定列格式（详见文档）

### 贡献
欢迎贡献代码！贡献步骤如下：
1. Fork 本仓库。
2. 创建新分支（`git checkout -b feature/your-feature`）。
3. 提交更改（`git commit -m "Add your feature"`）。
4. 推送分支（`git push origin feature/your-feature`）。
5. 提交 Pull Request。

请通过 [Issues](https://github.com/your-username/DataAnalysisTool/issues) 页面报告问题或建议功能。

### 许可证
本项目采用 MIT 许可证，详情见 [LICENSE](LICENSE) 文件。

## Version History
- **v1.0.0** (2025-08-06): Initial release with cross-regional shipping cost analysis, abnormal order detection, and product association analysis.
