# Excel文件合并工具

一个用于合并多个Excel文件的图形界面工具，支持.xls和.xlsx格式。

## 功能特点

- 🖥️ 图形用户界面，操作简单直观
- 📁 通过对话框选择文件夹
- 📊 自动检测并合并所有Excel文件
- 📈 实时显示处理进度
- ✅ 支持.xls和.xlsx格式
- 🎯 输出文件保存到选择的文件夹中

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

## 系统要求

- **macOS**: 10.14+ (支持Intel和Apple Silicon)
- **Windows**: Windows 7+
- **Linux**: 大多数现代发行版

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