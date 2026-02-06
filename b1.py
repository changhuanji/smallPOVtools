import tkinter as tk
from tkinter import messagebox
import os
import subprocess
import platform

# ==========================================
#  配置区域 (请修改这里)
# ==========================================
# 在这里填入你的 exe 绝对路径
# 注意：
# 1. 路径两边加上引号
# 2. 路径前面加上 r，防止反斜杠转义
# 3. 示例: r"D:\Games\Genshin Impact\Genshin Impact Game\YuanShen.exe"

TARGET_EXE_PATH = r"E:\povtools\7.fps计算.exe"


# ==========================================
#  核心逻辑
# ==========================================
def _launch_program():
    path = TARGET_EXE_PATH

    # 1. 检查文件是否存在
    if not os.path.exists(path):
        messagebox.showerror("错误", f"找不到文件：\n{path}\n\n请检查代码中的 TARGET_EXE_PATH 配置。")
        return

    try:
        # 2. 获取目标文件夹路径
        # 很多软件依赖同目录下的 dll 或 config 文件，必须设置 cwd (当前工作目录)
        work_dir = os.path.dirname(path)

        # 3. 启动进程
        # subprocess.Popen 是非阻塞的，主程序不会卡死
        if platform.system() == 'Windows':
            subprocess.Popen(path, cwd=work_dir, shell=False)
        else:
            # Linux/Mac 兼容写法 (虽然 exe 通常是 Windows 的)
            subprocess.Popen([path], cwd=work_dir)

        # 启动成功提示 (可选，不想弹窗可以注释掉)
        # messagebox.showinfo("成功", "程序已启动！")

    except Exception as e:
        messagebox.showerror("启动失败", f"无法启动程序：\n{str(e)}")


# ==========================================
#  UI 界面
# ==========================================
def show_ui(parent):
    top = tk.Toplevel(parent)
    top.title("快速启动器")
    top.geometry("400x250")
    top.transient(parent) # 修改点
    top.grab_set()        # 修改点

    # 居中
    top.update_idletasks()
    x = (top.winfo_screenwidth() - top.winfo_width()) // 2
    y = (top.winfo_screenheight() - top.winfo_height()) // 2
    top.geometry(f"+{x}+{y}")

    # 显示当前的配置路径
    tk.Label(top, text="当前配置的目标程序:", font=("Arial", 10)).pack(pady=(20, 5))

    # 使用 Text 组件显示路径，方便长路径自动换行
    txt_path = tk.Text(top, height=3, width=40, bg="#f0f0f0", relief="flat", font=("Consolas", 9))
    txt_path.insert("1.0", TARGET_EXE_PATH)
    txt_path.configure(state="disabled")  # 只读
    txt_path.pack(padx=20, pady=5)

    tk.Label(top, text="(如需修改，请打开 launcher_tools.py 编辑第12行)", fg="gray", font=("Arial", 8)).pack(pady=5)

    # 启动按钮
    def run():
        _launch_program()
        # 可选：启动后自动关闭这个小窗口
        # top.destroy()

    tk.Button(top, text="立即启动", bg="#E91E63", fg="white", font=("Arial", 14, "bold"),
              command=run).pack(pady=20, fill="x", padx=40)