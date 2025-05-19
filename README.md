# ZenExporter 🐛📁

自动化导出禅道Bug关联图片和附件的Python工具

## ✨ 核心功能
- 自动认证获取禅道API Token
- 智能解析Bug步骤中的图片链接
- 批量下载附件并保留原始名称
- 自动生成带超链接的Excel报告
- 进度可视化和完成提示

## 📦 依赖环境
```bash
Python >= 3.8
altgraph==0.17.4
openpyxl==3.1.5
requests==2.32.3
filetype==1.2.0
```

## ⚙️ 配置说明
```ini
[zentao]
url = http://your.zentao.server:port
username = your_account
password = your_password

[excel]
file_path = bugs.xlsx
bug_id_column = A
start_row = 2

[logs]
log_file = export.log
```

## 🚀 快速开始
```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python export_bug_images.py

# 打包为EXE
pyinstaller export_bug_images.spec