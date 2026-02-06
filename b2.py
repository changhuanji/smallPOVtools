import tkinter as tk
from tkinter import messagebox
import subprocess
import platform
import webbrowser


# ===========================
# 核心功能逻辑
# ===========================
def _open_url_in_edge(url):
    """
    尝试使用 Microsoft Edge 打开指定 URL
    """
    current_os = platform.system()

    try:
        if current_os == 'Windows':
            # Windows 下使用 start 命令调用 msedge
            # shell=True 是必须的，因为 start 是 shell 命令
            subprocess.run(f'start msedge "{url}"', shell=True, check=True)

        elif current_os == 'Darwin':  # macOS
            # Mac 下尝试调用 Edge (如果安装了的话)
            try:
                subprocess.run(['open', '-a', 'Microsoft Edge', url], check=True)
            except:
                # 如果没装 Edge，回退到默认浏览器
                webbrowser.open(url)
        else:
            # Linux 或其他系统，直接使用默认浏览器
            webbrowser.open(url)

    except Exception as e:
        # 如果调用 Edge 失败（比如没安装），回退到系统默认浏览器
        print(f"Edge 启动失败，尝试默认浏览器: {e}")
        webbrowser.open(url)


def _open_map(service_name):
    urls = {
        "gaode": "https://www.amap.com",
        "baidu": "https://map.baidu.com",
        "tencent": "https://map.qq.com"
    }

    target_url = urls.get(service_name)
    if target_url:
        _open_url_in_edge(target_url)


def _api_placeholder():
    messagebox.showinfo("API 查询",
                        "此功能为预留接口。\n\n未来可接入：\n1. 天气查询 API\n2. 汇率换算 API\n3. 快递查询 API")


# ===========================
# UI 界面逻辑
# ===========================
def show_ui(parent):
    top = tk.Toplevel(parent)
    top.title("网络工具箱")
    top.geometry("400x350")
    top.transient(parent) # 修改点
    top.grab_set()        # 修改点

    # 居中
    top.update_idletasks()
    x = (top.winfo_screenwidth() - top.winfo_width()) // 2
    y = (top.winfo_screenheight() - top.winfo_height()) // 2
    top.geometry(f"+{x}+{y}")

    pad_opts = {'padx': 20, 'pady': 10}

    # === 模块 1: 网页地图 (2选1 的第一个选项) ===
    # 使用 LabelFrame 框起来，视觉上区分
    frame_map = tk.LabelFrame(top, text="选项 1: 在 Edge 中打开地图", font=("Arial", 10, "bold"), fg="#2196F3")
    frame_map.pack(fill="x", **pad_opts)

    # 次级界面：3个具体地图选项
    btn_style = {"font": ("Arial", 10), "width": 12, "cursor": "hand2"}

    # 高德
    tk.Button(frame_map, text="高德地图", bg="#E3F2FD",
              command=lambda: _open_map("gaode"), **btn_style).pack(pady=5)

    # 百度
    tk.Button(frame_map, text="百度地图", bg="#E3F2FD",
              command=lambda: _open_map("baidu"), **btn_style).pack(pady=5)

    # 腾讯
    tk.Button(frame_map, text="腾讯地图", bg="#E3F2FD",
              command=lambda: _open_map("tencent"), **btn_style).pack(pady=5)

    # === 模块 2: API 查询 (2选1 的第二个选项) ===
    frame_api = tk.LabelFrame(top, text="选项 2: API 数据服务", font=("Arial", 10, "bold"), fg="#4CAF50")
    frame_api.pack(fill="x", **pad_opts)

    tk.Button(frame_api, text="API 查询 (预留)", bg="#E8F5E9", width=20, height=2,
              command=_api_placeholder).pack(pady=10)