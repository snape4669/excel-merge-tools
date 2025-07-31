# Windows构建说明

## 方法一：使用GitHub Actions (推荐)

1. 将此文件夹上传到GitHub仓库
2. 在仓库中创建 `.github/workflows/build.yml` 文件
3. 推送代码，GitHub Actions会自动构建Windows可执行文件

## 方法二：使用本地Windows环境

1. 在Windows上安装Python 3.8+
2. 安装依赖：pip install -r requirements.txt
3. 安装PyInstaller：pip install pyinstaller
4. 运行构建：pyinstaller --onefile --windowed --name="Excel文件合并工具" merge_excel_file.py

## 方法三：使用在线构建服务

- Replit: https://replit.com
- Gitpod: https://gitpod.io
- GitHub Codespaces: https://github.com/features/codespaces
