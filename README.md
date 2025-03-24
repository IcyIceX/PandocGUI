# PandocGUI

找了半天没找到mac版本带GUI的pandoc,就用AI做了一个

## 说明：
- `pypandoc` 是唯一需要通过 `pip` 安装的外部依赖。 

## 功能
- 批量将指定文件夹中的 Markdown 文件转换为 Docx 格式。
- 提供直观的 GUI，支持选择输入和输出文件夹。
- 支持 PyInstaller 打包为独立应用程序（需内置 Pandoc 可执行文件）。
- 跨平台兼容（Windows、macOS、Linux）。

## 安装

### 前置条件
- **Python 3.6+**: 确保您的系统已安装 Python。
- **Pandoc**: 需要在系统 PATH 中安装 Pandoc，或者在打包时提供 Pandoc 二进制文件。

### 安装步骤
1. 克隆仓库：
   ```bash
   git clone https://github.com/IcyIceX/PandocGUI.git
   cd pandocGUI
   ```

2. 创建并激活虚拟环境（可选但推荐）：
   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/macOS
   venv\Scripts\activate     # Windows
   ```

3. 安装依赖：
   ```bash
   pip install pypandoc
   ```

4. 运行脚本：
   ```bash
   python pandocGUI.py
   ```

## 使用方法
1. 启动程序后，界面会显示“输入文件夹”和“输出文件夹”两个选项。
2. 点击“浏览”按钮选择包含 Markdown 文件的输入文件夹。
3. 点击“浏览”按钮选择保存 Docx 文件的输出文件夹。
4. 点击“开始转换”按钮，程序将批量转换所有 `.md` 文件并显示状态。