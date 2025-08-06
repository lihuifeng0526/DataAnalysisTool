DataAnalysisToolEnglish | 中文 (#中文)EnglishOverviewDataAnalysisTool is a standalone Windows application designed for processing and analyzing e-commerce data, including cross-regional shipping cost data, abnormal order detection, and product association analysis. Built with Python and PyQt5, it is packaged as a single executable file using PyInstaller, requiring no Python environment or additional dependencies.FeaturesCross-Regional Shipping Cost Analysis: Processes order, inventory, and shipping data to generate cost reports, supporting Excel file input and output.
Abnormal Order Detection: Identifies anomalies in order data based on warehouse and province matching.
Product Association Analysis: Uses the Apriori algorithm to analyze product purchase patterns.
User-Friendly GUI: Intuitive interface with progress bars and detailed logs for data processing.
Single Executable: No installation required, runs directly on Windows 10/11 (64-bit).

InstallationDownload the latest release (DataAnalysisTool.exe) from the GitHub Releases page.
Ensure your system has Microsoft Visual C++ Redistributable installed (required for some dependencies).
Place the .exe file in a folder with write permissions (e.g., C:\Users\YourName\Documents).
Double-click DataAnalysisTool.exe to launch the application.

UsageLaunch the Application: Run DataAnalysisTool.exe. The GUI will display three main modules:Cross-Regional Shipping Cost Data
Abnormal Order Data
Product Association Analysis

Select Module: Click on a module (e.g., "Cross-Regional Shipping Cost Data").
Input Files:Select order data Excel files (.xlsx or .xls).
Choose an inventory data Excel file.
Select a shipping data Excel file and specify sheet names for Foshan and Jinan data.

Run Processing: Click "Run Processing" to analyze data. Check the output directory for results (e.g., 最终结果_超区发货费用数据表.xlsx).
View Logs: Errors and processing details are logged in app.log in the same directory as the .exe.

System RequirementsOperating System: Windows 10/11 (64-bit)
Disk Space: ~200 MB for the executable and output files
Dependencies: None (all included in the .exe)
Input Files: Excel files (.xlsx or .xls) with specific column formats (refer to documentation for details)

ContributingContributions are welcome! To contribute:Fork the repository.
Create a new branch (git checkout -b feature/your-feature).
Commit your changes (git commit -m "Add your feature").
Push to the branch (git push origin feature/your-feature).
Open a Pull Request.

Please report issues or suggest features via the Issues page.LicenseThis project is licensed under the MIT License. See the LICENSE file for details.中文概述DataAnalysisTool 是一款独立的 Windows 应用程序，用于处理和分析电商数据，包括超区发货费用数据、异常订单检测和商品关联性分析。该工具使用 Python 和 PyQt5 开发，通过 PyInstaller 打包为单一可执行文件，无需 Python 环境或额外依赖。功能超区发货费用分析：处理订单、库存和发货数据，生成费用报表，支持 Excel 文件输入和输出。
异常订单检测：根据仓库和省份匹配规则识别订单数据中的异常。
商品关联性分析：使用 Apriori 算法分析商品购买模式。
用户友好的界面：直观的图形界面，包含进度条和详细的处理日志。
单一可执行文件：无需安装，直接在 Windows 10/11（64 位）上运行。

安装从 GitHub Releases 页面下载最新版本的 DataAnalysisTool.exe。
确保系统已安装 Microsoft Visual C++ Redistributable（部分依赖需要）。
将 .exe 文件放置在具有写入权限的文件夹（例如 C:\Users\YourName\Documents）。
双击 DataAnalysisTool.exe 启动应用程序。

使用方法启动应用程序：运行 DataAnalysisTool.exe，界面将显示三个主要模块：超区发货费用数据
异常订单数据
商品关联性分析

选择模块：点击所需模块（例如“超区发货费用数据”）。
输入文件：选择订单数据的 Excel 文件（.xlsx 或 .xls）。
选择库存数据的 Excel 文件。
选择发货数据的 Excel 文件，并指定佛山和济南数据的 sheet 名称。

运行处理：点击“运行处理”分析数据。结果将保存在输出目录（例如 最终结果_超区发货费用数据表.xlsx）。
查看日志：错误和处理详情记录在 .exe 所在目录的 app.log 文件中。

系统要求操作系统：Windows 10/11（64 位）
磁盘空间：约 200 MB 用于可执行文件和输出文件
依赖项：无（所有依赖已包含在 .exe 中）
输入文件：Excel 文件（.xlsx 或 .xls），需符合特定列格式（详见文档）

贡献欢迎贡献代码！贡献步骤如下：Fork 本仓库。
创建新分支（git checkout -b feature/your-feature）。
提交更改（git commit -m "Add your feature"）。
推送分支（git push origin feature/your-feature）。
提交 Pull Request。

请通过 Issues 页面报告问题或建议功能。许可证本项目采用 MIT 许可证，详情见 LICENSE 文件。

