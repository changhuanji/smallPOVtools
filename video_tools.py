import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import cv2
import numpy as np
import math
import threading
import os
import subprocess
import shutil


# ===========================
# 辅助：检查 FFmpeg
# ===========================
def check_ffmpeg():
    return shutil.which("ffmpeg") is not None


# ===========================
# 核心渲染逻辑
# ===========================
def _render_video_thread(img_path, output_path, res_mode, res_w, res_h, fps,
                         angle, distance, speed, pos_mode, custom_pos_x, custom_pos_y,
                         use_gpu, progress_callback, done_callback):
    pipe = None
    try:
        # 1. 检查 FFmpeg
        if not check_ffmpeg():
            raise RuntimeError("未检测到 ffmpeg.exe！\n请下载 ffmpeg.exe 并放到本软件同级目录下。")

        # 2. 读取图片
        img = cv2.imdecode(np.fromfile(img_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
        if img is None:
            raise ValueError("无法读取图片。")

        if len(img.shape) == 2:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
        elif img.shape[2] == 3:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)

        img_h, img_w = img.shape[:2]

        # 3. 解析分辨率
        canvas_w, canvas_h = 1920, 1080
        if res_mode == "1080P":
            canvas_w, canvas_h = 1920, 1080
        elif res_mode == "4K":
            canvas_w, canvas_h = 3840, 2160
        elif res_mode == "自定义":
            try:
                canvas_w, canvas_h = int(res_w), int(res_h)
                if canvas_w <= 0 or canvas_h <= 0: raise ValueError
            except:
                raise ValueError("自定义分辨率宽高必须为正整数")

        # 4. 解析起始位置 (start_x, start_y)
        # 目标：计算图片左上角在画布上的坐标
        start_x, start_y = 0, 0

        if pos_mode == "正中":
            start_x = (canvas_w - img_w) / 2.0
            start_y = (canvas_h - img_h) / 2.0
        elif pos_mode == "左上":
            start_x, start_y = 0, 0
        elif pos_mode == "左下":
            start_x, start_y = 0, canvas_h - img_h
        elif pos_mode == "右上":
            start_x, start_y = canvas_w - img_w, 0
        elif pos_mode == "右下":
            start_x, start_y = canvas_w - img_w, canvas_h - img_h
        elif pos_mode == "上边正中":
            start_x = (canvas_w - img_w) / 2.0
            start_y = 0
        elif pos_mode == "下边正中":
            start_x = (canvas_w - img_w) / 2.0
            start_y = canvas_h - img_h
        elif pos_mode == "左边正中":
            start_x = 0
            start_y = (canvas_h - img_h) / 2.0
        elif pos_mode == "右边正中":
            start_x = canvas_w - img_w
            start_y = (canvas_h - img_h) / 2.0
        elif pos_mode == "自定义(中心点坐标)":
            try:
                cx, cy = float(custom_pos_x), float(custom_pos_y)
                start_x = cx - img_w / 2.0
                start_y = cy - img_h / 2.0
            except:
                raise ValueError("自定义坐标 X/Y 必须为数字")

        # 5. 运动参数
        total_frames = int((distance / speed) * fps)
        if total_frames <= 0: total_frames = fps

        rad = math.radians(angle)
        vel_x = speed * math.cos(rad)
        vel_y = speed * math.sin(rad)
        dx_per_frame = vel_x / fps
        dy_per_frame = vel_y / fps

        # 6. FFmpeg 命令
        if not output_path.lower().endswith('.mov'):
            output_path = os.path.splitext(output_path)[0] + ".mov"

        common_input = [
            'ffmpeg', '-y', '-f', 'rawvideo', '-vcodec', 'rawvideo',
            '-s', f'{canvas_w}x{canvas_h}', '-pix_fmt', 'bgra',
            '-r', str(fps), '-i', '-'
        ]

        if use_gpu:
            print("Trying NVIDIA GPU (HEVC)...")
            enc_flags = ['-c:v', 'hevc_nvenc', '-pix_fmt', 'yuva444p', '-preset', 'p7', '-rc', 'vbr', '-b:v', '20M',
                         output_path]
        else:
            print("Using CPU (ProRes)...")
            enc_flags = ['-c:v', 'prores_ks', '-profile:v', '4', '-pix_fmt', 'yuva444p10le', '-vendor', 'apl0',
                         output_path]

        command = common_input + enc_flags

        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        pipe = subprocess.Popen(command, stdin=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo)

        # 7. 渲染
        for i in range(total_frames):
            curr_x = int(start_x + dx_per_frame * i)
            curr_y = int(start_y + dy_per_frame * i)

            canvas = np.zeros((canvas_h, canvas_w, 4), dtype=np.uint8)

            x1_c, y1_c = max(0, curr_x), max(0, curr_y)
            x2_c, y2_c = min(canvas_w, curr_x + img_w), min(canvas_h, curr_y + img_h)
            x1_i, y1_i = max(0, -curr_x), max(0, -curr_y)

            w_slice = x2_c - x1_c
            h_slice = y2_c - y1_c

            if w_slice > 0 and h_slice > 0:
                canvas[y1_c:y2_c, x1_c:x2_c] = img[y1_i:y1_i + h_slice, x1_i:x1_i + w_slice]

            try:
                pipe.stdin.write(canvas.tobytes())
            except Exception as e:
                _, stderr = pipe.communicate()
                err_msg = stderr.decode('utf-8', errors='ignore')
                if "hevc_nvenc" in err_msg:
                    raise RuntimeError(f"GPU编码失败: {err_msg}")
                else:
                    raise RuntimeError(f"FFmpeg写入错误: {err_msg}")

            if i % 10 == 0:
                progress_callback((i + 1) / total_frames * 100)

        pipe.stdin.close()
        _, stderr = pipe.communicate()
        if pipe.returncode != 0:
            raise RuntimeError(f"FFmpeg异常退出:\n{stderr.decode('utf-8', errors='ignore')}")

        progress_callback(100)
        done_callback(None, output_path)

    except Exception as e:
        if pipe:
            try:
                pipe.kill()
            except:
                pass
        done_callback(str(e), None)


# ===========================
# UI 界面逻辑
# ===========================
def show_ui(parent):
    top = tk.Toplevel(parent)
    top.title("透明MOV生成器 (ProRes/GPU)")
    top.geometry("700x900")  # 加高以容纳新选项
    top.update_idletasks()
    x = (top.winfo_screenwidth() - top.winfo_width()) // 2
    y = (top.winfo_screenheight() - top.winfo_height()) // 2
    top.geometry(f"+{x}+{y}")

    pad_opts = {'padx': 10, 'pady': 8}

    has_ffmpeg = check_ffmpeg()
    status_bg = "#4CAF50" if has_ffmpeg else "#FF5722"
    status_txt = "FFmpeg 就绪" if has_ffmpeg else "未检测到 ffmpeg.exe"
    tk.Label(top, text=status_txt, bg=status_bg, fg="white", font=("Arial", 10, "bold")).pack(fill="x")

    # 1. 文件
    tk.Label(top, text="源图片:").pack(anchor="w", **pad_opts)
    entry_in = tk.Entry(top);
    entry_in.pack(fill="x", padx=10)
    tk.Button(top, text="浏览", command=lambda: (entry_in.delete(0, tk.END), entry_in.insert(0,
                                                                                             filedialog.askopenfilename(
                                                                                                 parent=top, filetypes=[
                                                                                                     ("Images",
                                                                                                      "*.png")])))).pack(
        anchor="e", padx=10)

    tk.Label(top, text="输出视频 (.mov):").pack(anchor="w", **pad_opts)
    entry_out = tk.Entry(top);
    entry_out.pack(fill="x", padx=10)

    def sel_out():
        f = filedialog.asksaveasfilename(defaultextension=".mov", filetypes=[("Video", "*.mov")], parent=top)
        if f: entry_out.delete(0, tk.END); entry_out.insert(0, f)

    tk.Button(top, text="保存", command=sel_out).pack(anchor="e", padx=10)

    tk.Frame(top, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    # 2. 画布设置 (Grid)
    frame_res = tk.Frame(top);
    frame_res.pack(fill="x", padx=10)
    tk.Label(frame_res, text="分辨率:").grid(row=0, column=0, sticky="w")

    var_res = tk.StringVar(value="1080P")

    def toggle_custom_res():
        state = "normal" if var_res.get() == "自定义" else "disabled"
        e_cw.config(state=state)
        e_ch.config(state=state)

    tk.Radiobutton(frame_res, text="1080P", variable=var_res, value="1080P", command=toggle_custom_res).grid(row=0,
                                                                                                             column=1)
    tk.Radiobutton(frame_res, text="4K", variable=var_res, value="4K", command=toggle_custom_res).grid(row=0, column=2)
    tk.Radiobutton(frame_res, text="自定义", variable=var_res, value="自定义", command=toggle_custom_res).grid(row=0,
                                                                                                               column=3)

    e_cw = tk.Entry(frame_res, width=6);
    e_cw.insert(0, "1920");
    e_cw.grid(row=0, column=4, padx=2)
    tk.Label(frame_res, text="x").grid(row=0, column=5)
    e_ch = tk.Entry(frame_res, width=6);
    e_ch.insert(0, "1080");
    e_ch.grid(row=0, column=6, padx=2)
    toggle_custom_res()

    tk.Label(frame_res, text="帧率:").grid(row=1, column=0, sticky="w", pady=5)
    var_fps = tk.IntVar(value=60)
    tk.Radiobutton(frame_res, text="60", variable=var_fps, value=60).grid(row=1, column=1)
    tk.Radiobutton(frame_res, text="120", variable=var_fps, value=120).grid(row=1, column=2)

    tk.Label(frame_res, text="编码:").grid(row=2, column=0, sticky="w", pady=5)
    var_gpu = tk.BooleanVar(value=False)
    tk.Checkbutton(frame_res, text="NVIDIA GPU (HEVC)", variable=var_gpu, fg="#E91E63").grid(row=2, column=1,
                                                                                             columnspan=3, sticky="w")

    tk.Frame(top, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    # 3. 起始位置设置
    frame_pos = tk.Frame(top);
    frame_pos.pack(fill="x", padx=10)
    tk.Label(frame_pos, text="起始位置:").grid(row=0, column=0, sticky="w")

    pos_options = ["正中", "左上", "左下", "右上", "右下", "上边正中", "下边正中", "左边正中", "右边正中",
                   "自定义(中心点坐标)"]
    combo_pos = ttk.Combobox(frame_pos, values=pos_options, width=15, state="readonly")
    combo_pos.current(0)
    combo_pos.grid(row=0, column=1, columnspan=2, sticky="w")

    tk.Label(frame_pos, text="X:").grid(row=0, column=3, padx=2)
    e_cx = tk.Entry(frame_pos, width=6);
    e_cx.insert(0, "960");
    e_cx.grid(row=0, column=4)
    tk.Label(frame_pos, text="Y:").grid(row=0, column=5, padx=2)
    e_cy = tk.Entry(frame_pos, width=6);
    e_cy.insert(0, "540");
    e_cy.grid(row=0, column=6)

    def on_pos_change(event):
        state = "normal" if combo_pos.get() == "自定义(中心点坐标)" else "disabled"
        e_cx.config(state=state)
        e_cy.config(state=state)

    combo_pos.bind("<<ComboboxSelected>>", on_pos_change)
    on_pos_change(None)

    tk.Frame(top, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    # 4. 运动参数
    frame_move = tk.Frame(top);
    frame_move.pack(fill="x", padx=10)
    tk.Label(frame_move, text="方向(角度):").grid(row=0, column=0, sticky="w");
    e_ang = tk.Entry(frame_move, width=8);
    e_ang.insert(0, "0");
    e_ang.grid(row=0, column=1)

    tk.Label(frame_move, text="距离(px):").grid(row=0, column=2, padx=5)
    e_dist = tk.Entry(frame_move, width=8);
    e_dist.insert(0, "500");
    e_dist.grid(row=0, column=3)

    tk.Label(frame_move, text="速度(px/s):").grid(row=0, column=4, padx=5)
    e_spd = tk.Entry(frame_move, width=8);
    e_spd.insert(0, "100");
    e_spd.grid(row=0, column=5)

    progress_bar = ttk.Progressbar(top, length=400, mode="determinate");
    progress_bar.pack(pady=10)
    lbl_stat = tk.Label(top, text="Ready", fg="gray");
    lbl_stat.pack()

    def run():
        f_in = entry_in.get();
        f_out = entry_out.get()
        if not f_in or not f_out: return messagebox.showwarning("提示", "路径为空", parent=top)
        try:
            ang, dist, spd = float(e_ang.get()), float(e_dist.get()), float(e_spd.get())
            if spd <= 0: raise ValueError
        except:
            return messagebox.showerror("错误", "运动参数错误", parent=top)

        btn_run.config(state="disabled", text="Rendering...")

        def on_done(err, path):
            btn_run.config(state="normal", text="开始生成")
            if err:
                messagebox.showerror("失败", err, parent=top)
            else:
                messagebox.showinfo("成功", f"保存至:\n{path}", parent=top)

        t = threading.Thread(target=_render_video_thread, args=(
            f_in, f_out, var_res.get(), e_cw.get(), e_ch.get(), var_fps.get(),
            ang, dist, spd, combo_pos.get(), e_cx.get(), e_cy.get(),
            var_gpu.get(), lambda v: progress_bar.configure(value=v), on_done
        ))
        t.daemon = True;
        t.start()

    btn_run = tk.Button(top, text="开始生成", bg="#2196F3", fg="white", font=("Arial", 12, "bold"), command=run)
    btn_run.pack(pady=10, fill="x", padx=20)