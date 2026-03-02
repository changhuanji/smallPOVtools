import tkinter as tk
from tkinter import messagebox, ttk


# ==============================================================================
#  辅助类：时间码输入组 (H:M:S:F)
# ==============================================================================
class TimeCodeInput(tk.Frame):
    def __init__(self, parent, label_text, next_widget_group=None):
        super().__init__(parent)
        self.entries = []
        self.next_group = next_widget_group  # 指向下一个输入组，用于跨组跳转

        # 标签
        tk.Label(self, text=label_text, width=12, anchor="e").pack(side="left", padx=(0, 10))

        # 创建 4 个输入框
        labels = [":", ":", ":", ""]  # 分隔符
        self.vars = []

        for i in range(4):
            # 变量追踪用于限制长度
            var = tk.StringVar()
            var.trace("w", lambda name, index, mode, v=var, idx=i: self._on_type(v, idx))
            self.vars.append(var)

            # 输入框
            entry = tk.Entry(self, textvariable=var, width=3, justify="center", font=("Arial", 11))
            entry.pack(side="left")
            self.entries.append(entry)

            # 绑定焦点跳转
            # KeyRelease 用于检测输入完成，FocusIn 用于全选
            entry.bind("<KeyRelease>", lambda e, idx=i: self._check_jump(e, idx))
            entry.bind("<FocusIn>", lambda e: e.widget.select_range(0, tk.END))

            # 分隔符
            if labels[i]:
                tk.Label(self, text=labels[i], font=("Arial", 10, "bold")).pack(side="left")

        # 提示文字 (H M S F)
        tk.Label(self, text="(时:分:秒:帧)", fg="gray", font=("Arial", 8)).pack(side="left", padx=5)

    def _on_type(self, var, idx):
        """限制只能输入数字且最多2位"""
        val = var.get()
        if not val.isdigit():
            # 移除非数字字符
            new_val = "".join([c for c in val if c.isdigit()])
            var.set(new_val)
            val = new_val

        if len(val) > 2:
            var.set(val[:2])

    def _check_jump(self, event, idx):
        """输入满2位自动跳到下一个框"""
        # 忽略 Backspace, Tab 等控制键
        if event.keysym in ["BackSpace", "Tab", "Shift_L", "Shift_R", "Delete", "Left", "Right"]:
            return

        current_val = self.vars[idx].get()
        if len(current_val) >= 2:
            if idx < 3:
                # 组内跳转
                self.entries[idx + 1].focus_set()
            elif self.next_group:
                # 跳到下一个输入组的第一个框
                self.next_group.entries[0].focus_set()

    def set_next_group(self, group):
        """设置下一个跳转的目标组"""
        self.next_group = group

    def get_total_frames(self, fps):
        """计算总帧数"""
        try:
            h = int(self.vars[0].get() or 0)
            m = int(self.vars[1].get() or 0)
            s = int(self.vars[2].get() or 0)
            f = int(self.vars[3].get() or 0)

            total_seconds = h * 3600 + m * 60 + s
            return total_seconds * fps + f
        except ValueError:
            return 0


# ==============================================================================
#  核心逻辑与 UI
# ==============================================================================
class SpeedCalculatorApp:
    def __init__(self, root):
        self.top = tk.Toplevel(root)
        self.top.title("视频变速计算器")
        self.top.geometry("520x550")

        # 窗口层级处理
        self.top.transient(root)

        # 居中
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() - 520) // 2
        y = (self.top.winfo_screenheight() - 550) // 2
        self.top.geometry(f"+{x}+{y}")

        self.build_ui()

    def build_ui(self):
        # === 1. 顶部设置 ===
        frame_top = tk.Frame(self.top)
        frame_top.pack(fill="x", padx=20, pady=10)

        # 置顶选项
        self.var_topmost = tk.BooleanVar(value=False)
        chk_top = tk.Checkbutton(frame_top, text="窗口始终置顶", variable=self.var_topmost,
                                 command=self.toggle_topmost)
        chk_top.pack(side="left")

        # === 2. 帧率设置 ===
        frame_fps = tk.LabelFrame(self.top, text="帧率设置 (FPS)", fg="#2196F3")
        frame_fps.pack(fill="x", padx=20, pady=5)

        tk.Label(frame_fps, text="时间线/目标帧率:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_tl_fps = tk.Entry(frame_fps, width=8)
        self.entry_tl_fps.insert(0, "30")
        self.entry_tl_fps.grid(row=0, column=1, sticky="w")

        tk.Label(frame_fps, text="原素材帧率:").grid(row=0, column=2, padx=10, pady=10, sticky="e")
        self.entry_src_fps = tk.Entry(frame_fps, width=8)
        self.entry_src_fps.insert(0, "30")
        self.entry_src_fps.grid(row=0, column=3, sticky="w")

        # === 3. 目标时间段 (Timeline) ===
        frame_target = tk.LabelFrame(self.top, text="时间线目标范围 (Target)", fg="#FF9800")
        frame_target.pack(fill="x", padx=20, pady=10)

        # 实例化输入组件
        self.input_start = TimeCodeInput(frame_target, "起始时间:")
        self.input_start.pack(fill="x", pady=5)

        self.input_end = TimeCodeInput(frame_target, "结束时间:")
        self.input_end.pack(fill="x", pady=5)

        # === 4. 原素材长度 (Source) ===
        frame_source = tk.LabelFrame(self.top, text="原素材长度 (Source)", fg="#4CAF50")
        frame_source.pack(fill="x", padx=20, pady=10)

        self.input_src_len = TimeCodeInput(frame_source, "素材时长:")
        self.input_src_len.pack(fill="x", pady=5)

        # === 设置Tab跳转顺序 ===
        # 起始 -> 结束 -> 素材时长
        self.input_start.set_next_group(self.input_end)
        self.input_end.set_next_group(self.input_src_len)

        # === 5. 结果显示与操作 ===
        frame_action = tk.Frame(self.top)
        frame_action.pack(fill="both", expand=True, padx=20, pady=10)

        btn_calc = tk.Button(frame_action, text="计算并复制 (Calculate)",
                             bg="#673AB7", fg="white", font=("Arial", 12, "bold"),
                             command=self.calculate)
        btn_calc.pack(fill="x", pady=5)

        # 结果展示区
        frame_res = tk.Frame(frame_action, bg="#f0f0f0", bd=1, relief="sunken")
        frame_res.pack(fill="both", expand=True, pady=10)

        self.lbl_speed = tk.Label(frame_res, text="0.00 %", font=("Arial", 24, "bold"), bg="#f0f0f0", fg="#E91E63")
        self.lbl_speed.pack(pady=(20, 5))

        self.lbl_info = tk.Label(frame_res, text="目标时长: 0帧 | 原素材: 0帧", bg="#f0f0f0", fg="gray")
        self.lbl_info.pack(pady=(0, 20))

        btn_copy = tk.Button(frame_action, text="仅复制结果", command=self.copy_result)
        btn_copy.pack(fill="x")

    def toggle_topmost(self):
        self.top.attributes('-topmost', self.var_topmost.get())

    def calculate(self):
        try:
            # 获取 FPS
            try:
                tl_fps = float(self.entry_tl_fps.get())
                src_fps = float(self.entry_src_fps.get())
                if tl_fps <= 0 or src_fps <= 0: raise ValueError
            except:
                messagebox.showerror("错误", "帧率必须为正数", parent=self.top)
                return

            # 获取帧数
            # 起始和结束使用 时间线帧率
            f_start = self.input_start.get_total_frames(tl_fps)
            f_end = self.input_end.get_total_frames(tl_fps)

            # 原素材长度使用 原素材帧率 (因为通常查看素材属性时是基于其自身帧率的)
            # 但这里有歧义：如果用户是在时间线上看原素材长度，应该用时间线帧率。
            # 为了通用性，通常"素材时长"是指文件本身的属性，所以用 src_fps。
            f_src_len = self.input_src_len.get_total_frames(src_fps)

            target_duration = f_end - f_start

            if target_duration <= 0:
                messagebox.showerror("错误", "结束时间必须大于起始时间", parent=self.top)
                return

            if f_src_len <= 0:
                messagebox.showerror("错误", "原素材时长必须大于0", parent=self.top)
                return

            # 计算逻辑
            # 变速原理：
            # 速度 % = (原素材实际时长 / 目标坑位时长) * 100
            # 注意：这里要统一到绝对时间（秒）或者统一到同一帧率下的帧数。
            # 方法：统一转换为秒

            sec_target = target_duration / tl_fps
            sec_source = f_src_len / src_fps

            speed_ratio = (sec_source / sec_target) * 100

            # 格式化输出
            result_str = f"{speed_ratio:.2f}"

            # 更新界面
            self.lbl_speed.config(text=f"{result_str} %")
            self.lbl_info.config(
                text=f"目标时长: {sec_target:.2f}s ({int(target_duration)}帧)\n原素材: {sec_source:.2f}s ({int(f_src_len)}帧)")

            # 自动复制
            self.top.clipboard_clear()
            self.top.clipboard_append(result_str)
            self.top.update()  # 保持剪贴板内容

            # 闪烁提示
            original_bg = self.lbl_speed.cget("bg")
            self.lbl_speed.config(bg="#C8E6C9")  # 浅绿
            self.top.after(200, lambda: self.lbl_speed.config(bg=original_bg))

        except Exception as e:
            messagebox.showerror("计算错误", str(e), parent=self.top)

    def copy_result(self):
        val = self.lbl_speed.cget("text").replace(" %", "")
        self.top.clipboard_clear()
        self.top.clipboard_append(val)
        messagebox.showinfo("提示", "已复制到剪贴板", parent=self.top)


# ==============================================================================
#  模块入口
# ==============================================================================
def show_ui(parent):
    SpeedCalculatorApp(parent)