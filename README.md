# Excel文件合并工具

一个用于合并多个Excel文件的图形界面工具，支持.xls和.xlsx格式。

## 功能特点

- 🖥️ 图形用户界面，操作简单直观
- 📁 通过对话框选择文件夹
- 📊 自动检测并合并所有Excel文件
- 📈 实时显示处理进度
- ✅ 支持.xls和.xlsx格式
- 🎯 输出文件保存到选择的文件夹中
- 🏷️ 支持自定义标题行设置
- 🔍 自动检查标题行一致性

## 打包成可执行文件

### macOS 构建

#### 方法一：使用打包脚本（推荐）

```bash
# 构建可执行文件
python3 build.py

# 清理构建文件
python3 build.py clean
```

#### 方法二：直接使用PyInstaller

```bash
# macOS (当前架构)
pyinstaller --onefile --windowed --name="Excel文件合并工具" merge_excel_file.py

# 如果需要universal2支持（Intel + Apple Silicon），需要先安装universal2版本的pandas
# pip install pandas --upgrade --force-reinstall --no-binary :all:
# pyinstaller --onefile --windowed --name="Excel文件合并工具" --target-architecture=universal2 merge_excel_file.py
```

### Windows 构建

#### 方法一：使用批处理文件（推荐）

```cmd
# 双击运行或在命令提示符中执行
build_windows.bat
```

#### 方法二：直接使用PyInstaller

```cmd
# 在命令提示符中执行
pyinstaller --onefile --windowed --name="Excel文件合并工具" merge_excel_file.py
```

### Linux 构建

```bash
pyinstaller --onefile --windowed --name="Excel文件合并工具" merge_excel_file.py
```

### 方法三：使用配置文件

```bash
pyinstaller build_config.spec
```

## 依赖要求

在打包之前，确保安装了以下依赖：

```bash
pip3 install pyinstaller pandas openpyxl
```

## 使用说明

### 运行Python脚本

```bash
python3 merge_excel_file.py
```

### 运行可执行文件

1. 双击生成的可执行文件
2. 点击"选择文件夹"按钮
3. 选择包含Excel文件的文件夹
4. 点击"开始合并"按钮
5. 等待处理完成

## 输出文件

合并后的文件将保存为 `merged_file.xlsx`，位于选择的文件夹中。

## 自动构建和发布

本项目使用 GitHub Actions 实现自动构建和发布。当您创建新的版本标签时，系统会自动：

1. 构建 Windows 和 macOS 版本的可执行文件
2. 创建包含可执行文件和说明文档的安装包
3. 在 GitHub Releases 页面发布新版本
4. 提供下载链接

### 创建新版本

```bash
# 1. 提交代码更改
git add .
git commit -m "添加新功能"
git push origin main

# 2. 创建版本标签
git tag v1.0.1
git push origin v1.0.1

# 3. 查看构建进度
# 访问 GitHub Actions 页面查看构建状态
```

### 下载最新版本

访问 [GitHub Releases](https://github.com/your-username/excel-merge-tools/releases) 页面下载最新版本。

## 系统要求

- **macOS**: 10.14+ (支持Intel和Apple Silicon)
- **Windows**: Windows 7+

## 注意事项

1. 确保选择的文件夹中包含Excel文件（.xls或.xlsx格式）
2. 程序会自动跳过无法读取的文件并显示警告
3. 如果输出文件已存在，会被覆盖
4. 处理大文件时可能需要较长时间

## 故障排除

### 常见问题

1. **"未找到Excel文件"错误**
   - 确保选择的文件夹中包含.xls或.xlsx文件
   - 检查文件是否被其他程序占用

2. **"读取文件失败"错误**
   - 检查文件是否损坏
   - 确保文件格式正确

3. **可执行文件无法运行**
   - 在macOS上，可能需要在"系统偏好设置 > 安全性与隐私"中允许运行
   - 在Windows上，可能需要右键选择"以管理员身份运行"

### 获取帮助

如果遇到问题，请检查：
- Python版本（推荐3.8+）
- 依赖包是否正确安装
- 文件权限是否正确 