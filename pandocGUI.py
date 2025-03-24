import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import pypandoc
from tkinter import ttk

def get_pandoc_path():
    """
    获取 Pandoc 二进制文件的路径。
    如果在 PyInstaller 打包的环境中运行，则从临时目录中获取 Pandoc 路径。
    否则，默认使用系统中安装的 Pandoc。
    
    返回:
    str: Pandoc 可执行文件的路径。
    """
    if getattr(sys, 'frozen', False):
        # 在 PyInstaller 打包的环境中运行
        base_path = sys._MEIPASS  # PyInstaller 的临时解压目录
        pandoc_path = os.path.join(base_path, 'pandoc', 'bin', 'pandoc')
        if os.name == 'nt':  # Windows 系统需要添加 .exe 后缀
            pandoc_path += '.exe'
    else:
        # 在开发环境中运行
        pandoc_path = 'pandoc'  # 假设系统中已安装 Pandoc 并在 PATH 中
    return pandoc_path

def convert_md_to_docx(md_folder_path, docx_folder_path, pandoc_path):
    """
    将指定文件夹内的所有 Markdown 文件转换为 Docx 格式，并保存到目标文件夹。

    参数:
    md_folder_path: str, 包含 Markdown 文件的源文件夹路径。
    docx_folder_path: str, 转换后的 Docx 文件保存的目标文件夹路径。
    pandoc_path: str, Pandoc 二进制文件的路径。
    
    返回:
    tuple: (bool, str)，表示是否成功及相关消息。
    """
    # 如果目标文件夹不存在，则创建
    if not os.path.exists(docx_folder_path):
        os.makedirs(docx_folder_path)

    # 检查源文件夹中是否有 .md 文件
    md_files = [f for f in os.listdir(md_folder_path) if f.endswith('.md')]
    if not md_files:
        return False, "没有找到 Markdown 文件"

    try:
        converted_count = 0
        # 遍历所有 Markdown 文件并转换
        for filename in md_files:
            md_file_path = os.path.join(md_folder_path, filename)
            docx_file_path = os.path.join(docx_folder_path, os.path.splitext(filename)[0] + ".docx")
            # 直接调用 Pandoc 的命令行工具
            cmd = [
                pandoc_path,
                md_file_path,
                '-f', 'markdown',
                '-t', 'docx',
                '--standalone',
                '-o', docx_file_path
            ]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                raise Exception(f"Pandoc 错误: {result.stderr}")
            converted_count += 1
        return True, f"已成功转换 {converted_count} 个文件"
    except Exception as e:
        return False, f"转换过程中出错: {str(e)}"

class PandocGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Markdown 转 Word 工具")
        self.root.geometry("600x300")
        self.root.resizable(False, False)

        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TButton", font=("微软雅黑", 10))
        self.style.configure("TLabel", font=("微软雅黑", 10))
        self.style.configure("TEntry", font=("微软雅黑", 10))

        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 输入文件夹选择
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=10)
        ttk.Label(input_frame, text="输入文件夹:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.input_path_var = tk.StringVar()
        input_entry = ttk.Entry(input_frame, textvariable=self.input_path_var, width=40)
        input_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        browse_input_btn = ttk.Button(input_frame, text="浏览...", command=self.browse_input_folder)
        browse_input_btn.grid(row=0, column=2, padx=5, pady=5)
        input_frame.columnconfigure(1, weight=1)

        # 输出文件夹选择
        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=10)
        ttk.Label(output_frame, text="输出文件夹:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.output_path_var = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path_var, width=40)
        output_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        browse_output_btn = ttk.Button(output_frame, text="浏览...", command=self.browse_output_folder)
        browse_output_btn.grid(row=0, column=2, padx=5, pady=5)
        output_frame.columnconfigure(1, weight=1)

        # 状态显示
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=20)
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, foreground="gray", anchor="center")
        status_label.pack(fill=tk.X)

        # 转换按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        convert_button = ttk.Button(button_frame, text="开始转换", command=self.start_conversion, width=15)
        convert_button.pack(side=tk.TOP, anchor=tk.CENTER)

        # 设置默认路径
        current_dir = os.path.dirname(os.path.abspath(__file__))
        input_dir = os.path.join(current_dir, "input")
        output_dir = os.path.join(current_dir, "output")
        if os.path.exists(input_dir):
            self.input_path_var.set(input_dir)
        if os.path.exists(output_dir):
            self.output_path_var.set(output_dir)

    def browse_input_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含 Markdown 文件的文件夹")
        if folder_path:
            self.input_path_var.set(folder_path)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(title="选择保存 Word 文件的文件夹")
        if folder_path:
            self.output_path_var.set(folder_path)

    def start_conversion(self):
        input_path = self.input_path_var.get()
        output_path = self.output_path_var.get()

        # 检查路径是否有效
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("错误", "请选择有效的输入文件夹")
            return
        if not output_path:
            messagebox.showerror("错误", "请选择输出文件夹")
            return

        # 获取 Pandoc 路径
        pandoc_path = get_pandoc_path()

        # 检查 Pandoc 是否存在
        if not os.path.exists(pandoc_path):
            messagebox.showerror("错误", f"Pandoc 未找到: {pandoc_path}")
            return

        # 更新状态
        self.status_var.set("转换中...")
        self.root.update_idletasks()

        # 执行转换
        success, message = convert_md_to_docx(input_path, output_path, pandoc_path)

        # 更新状态并显示结果
        if success:
            self.status_var.set(message)
            messagebox.showinfo("成功", message)
        else:
            self.status_var.set(f"错误: {message}")
            messagebox.showerror("错误", message)

if __name__ == "__main__":
    root = tk.Tk()
    app = PandocGUI(root)
    root.mainloop()
