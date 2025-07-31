# Windows可执行文件构建指南

## 概述

由于您当前在macOS环境下，无法直接生成Windows的.exe文件。我们为您提供了几种解决方案来生成Windows可执行文件。

## 已生成的文件

✅ **excel_merger_windows_build.zip** - 包含所有必要的源代码和构建配置

## 构建方法

### 方法一：使用GitHub Actions (推荐)

1. **创建GitHub仓库**
   - 访问 https://github.com/new
   - 创建新仓库，例如：`excel-merger-tool`

2. **上传源代码**
   - 解压 `excel_merger_windows_build.zip`
   - 将解压后的文件上传到GitHub仓库

3. **自动构建**
   - 推送代码后，GitHub Actions会自动构建Windows可执行文件
   - 在仓库的Actions标签页下载生成的.exe文件

### 方法二：使用在线构建服务

#### Replit (免费)
1. 访问 https://replit.com
2. 创建新的Python项目
3. 上传源代码文件
4. 在终端中运行：
   ```bash
   pip install -r requirements.txt
   pyinstaller --onefile --windowed --name="Excel文件合并工具" merge_excel_file.py
   ```

#### Gitpod (免费)
1. 访问 https://gitpod.io
2. 连接GitHub仓库
3. 在终端中运行构建命令

### 方法三：本地Windows环境

如果您有Windows电脑：

1. **安装Python**
   - 下载并安装Python 3.8+：https://www.python.org/downloads/

2. **安装依赖**
   ```cmd
   pip install -r requirements.txt
   ```

3. **构建可执行文件**
   ```cmd
   pyinstaller --onefile --windowed --name="Excel文件合并工具" merge_excel_file.py
   ```

4. **查找生成的文件**
   - 生成的文件在 `dist` 目录中
   - 文件名为：`Excel文件合并工具.exe`

## 文件说明

- `merge_excel_file.py` - 主程序源代码
- `requirements.txt` - Python依赖包列表
- `BUILD_INSTRUCTIONS.md` - 详细构建说明
- `.github/workflows/build.yml` - GitHub Actions自动构建配置

## 使用生成的.exe文件

1. 将生成的 `Excel文件合并工具.exe` 复制到Windows电脑
2. 双击运行程序
3. 按照界面提示选择包含Excel文件的文件夹
4. 点击"开始合并"按钮
5. 合并后的文件将保存为 `merged_file.xlsx`

## 注意事项

- 生成的.exe文件可能比较大（约20-30MB），这是正常的
- 首次运行时，Windows可能会显示安全警告，选择"仍要运行"
- 确保目标文件夹中有Excel文件（.xls或.xlsx格式）

## 故障排除

### 常见问题

1. **"无法运行此应用"错误**
   - 右键点击.exe文件 → 属性 → 勾选"解除锁定"

2. **缺少DLL文件**
   - 确保在Windows 7+系统上运行
   - 安装Microsoft Visual C++ Redistributable

3. **程序无法启动**
   - 检查是否有杀毒软件阻止
   - 尝试以管理员身份运行

## 技术支持

如果遇到问题，请检查：
- Python版本兼容性
- 依赖包是否正确安装
- 系统权限设置 