import tkinter as tk
from tkinter import messagebox

# === 导入功能模块 ===
try:
    import a0
    import a1
    import a2
    import b0
    import b1
    import b2
except ImportError as e:
    print(f"严重错误：缺少模块文件 - {e}")
    print("请确保所有工具脚本都在同一目录下。")


# === 主程序逻辑 ===
def create_main_interface():
    root = tk.Tk()
    root.title("多功能工具箱 (v2.7)")

    w, h = 650, 300
    ws, hs = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{int((ws / 2) - (w / 2))}+{int((hs / 2) - (h / 2))}")

    frame_container = tk.Frame(root)
    frame_container.place(relx=0.5, rely=0.5, anchor="center")

    btn_config = {
        "font": ("Arial", 14, "bold"),
        "width": 12,
        "height": 2
    }
    grid_config = {"padx": 15, "pady": 15}

    # === 按钮映射表 ===
    # 这里的 lambda p=root: module.show_ui(p) 是关键
    # 它将主窗口 root 传递给子模块
    buttons_map = [
        # 第一行
        ("PPT生成器", lambda: a0.show_ui(root)),
        ("PPT改色/去空", lambda: a1.show_ui(root)),
        ("导出透明PNG", lambda: a2.show_ui(root)),

        # 第二行
        ("图片转动画", lambda: b0.show_ui(root)),
        ("启动外部程序", lambda: b1.show_ui(root)),
        ("地图/API工具", lambda: b2.show_ui(root)),
    ]

    idx = 0
    for row in range(2):
        for col in range(3):
            if idx < len(buttons_map):
                text, command_func = buttons_map[idx]
                btn = tk.Button(frame_container, text=text, command=command_func, **btn_config)
                btn.grid(row=row, column=col, **grid_config)
                idx += 1

    root.mainloop()


if __name__ == "__main__":
    create_main_interface()