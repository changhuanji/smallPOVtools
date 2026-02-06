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
    """检查系统路径中是否有 ffmpeg"""
    return shutil.which("ffmpeg") is not None


# ===========================
# 核心渲染逻辑 (FFmpeg 管道模式)
# ===========================
def _render_video_thread(img_path, output_path, resolution, fps, angle, distance, speed, use_gpu, progress_callback,
                         done_callback):
    pipe = None
    try:
        # 1. 强制检查 FFmpeg
        if not check_ffmpeg():
            raise RuntimeError(
                "未检测到 ffmpeg.exe！\n\n必须安装 FFmpeg 才能生成视频。\n请下载 ffmpeg.exe 并放到本软件同级目录下。")

        # 2. 读取源图片 (强制保留 Alpha)
        img = cv2.imdecode(np.fromfile(img_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
        if img is None:
            raise ValueError("无法读取图片，请检查文件。")

        # 3. 确保图片是 4 通道 (BGRA)
        if len(img.shape) == 2:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
        elif img.shape[2] == 3:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)

        img_h, img_w = img.shape[:2]

        # 4. 参数计算
        canvas_w, canvas_h = resolution
        total_frames = int((distance / speed) * fps)
        if total_frames <= 0: total_frames = fps

        rad = math.radians(angle)
        vel_x = speed * math.cos(rad)
        vel_y = speed * math.sin(rad)
        dx_per_frame = vel_x / fps
        dy_per_frame = vel_y / fps

        # 5. 构建 FFmpeg 命令
        if not output_path.lower().endswith('.mov'):
            output_path = os.path.splitext(output_path)[0] + ".mov"

        common_input_flags = [
            'ffmpeg',
            '-y',
            '-f', 'rawvideo',
            '-vcodec', 'rawvideo',
            '-s', f'{canvas_w}x{canvas_h}',
            '-pix_fmt', 'bgra',
            '-r', str(fps),
            '-i', '-',
        ]

        if use_gpu:
            # === GPU 模式 (NVIDIA HEVC) ===
            print("正在尝试使用 NVIDIA GPU 编码 (HEVC)...")
            encoder_flags = [
                '-c:v', 'hevc_nvenc',
                '-pix_fmt', 'yuva444p',
                '-preset', 'p7',
                '-tune', 'hq',
                '-rc', 'vbr',
                '-b:v', '20M',
                output_path
            ]
        else:
            # === CPU 模式 (Apple ProRes 4444) ===
            print("正在使用 CPU 编码 (ProRes 4444)...")
            encoder_flags = [
                '-c:v', 'prores_ks',
                '-profile:v', '4',
                '-pix_fmt', 'yuva444p10le',
                '-vendor', 'apl0',
                output_path
            ]

        command = common_input_flags + encoder_flags

        # 打开管道
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        pipe = subprocess.Popen(command, stdin=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo)

        # 6. 逐帧渲染
        start_x = (canvas_w - img_w) / 2.0
        start_y = (canvas_h - img_h) / 2.0

        for i in range(total_frames):
            curr_x = int(start_x + dx_per_frame * i)
            curr_y = int(start_y + dy_per_frame * i)

            # 创建透明画布 (BGRA)
            canvas = np.zeros((canvas_h, canvas_w, 4), dtype=np.uint8)

            # 计算 ROI
            x1_c, y1_c = max(0, curr_x), max(0, curr_y)
            x2_c, y2_c = min(canvas_w, curr_x + img_w), min(canvas_h, curr_y + img_h)
            x1_i, y1_i = max(0, -curr_x), max(0, -curr_y)

            w_slice = x2_c - x1_c
            h_slice = y2_c - y1_c

            if w_slice > 0 and h_slice > 0:
                canvas[y1_c:y2_c, x1_c:x2_c] = img[y1_i:y1_i + h_slice, x1_i:x1_i + w_slice]

            # 写入管道
            try:
                pipe.stdin.write(canvas.tobytes())
            except Exception as e:
                # 发生写入错误时，立即读取 stderr 查明原因
                _, stderr = pipe.communicate()
                err_msg = stderr.decode('utf-8', errors='ignore')
                if "hevc_nvenc" in err_msg:
                    raise RuntimeError(f"GPU 编码失败：未检测到支持的 NVIDIA 显卡。\n\n详细错误: {err_msg}")
                else:
                    raise RuntimeError(f"FFmpeg 写入错误: {err_msg}")

            # 进度
            if i % 10 == 0:
                prog = (i + 1) / total_frames * 100
                progress_callback(prog)

        # === 关键修复点：防止 99% 卡死 ===

        # 1. 先关闭输入流，告诉 FFmpeg 数据发完了
        pipe.stdin.close()

        # 2. 使用 communicate() 代替 wait()
        # 这会读取并清空 stderr 缓冲区，防止死锁
        _, stderr = pipe.communicate()

        # 3. 检查返回值
        if pipe.returncode != 0:
            err_msg = stderr.decode('utf-8', errors='ignore')
            raise RuntimeError(f"FFmpeg 异常退出 (Code {pipe.returncode}):\n{err_msg}")

        # 强制进度条走完
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
    top.geometry("550x750")

    top.transient(parent)
    top.grab_set()

    top.update_idletasks()
    x = (top.winfo_screenwidth() - top.winfo_width()) // 2
    y = (top.winfo_screenheight() - top.winfo_height()) // 2
    top.geometry(f"+{x}+{y}")

    pad_opts = {'padx': 10, 'pady': 8}

    # === FFmpeg 状态检测 ===
    has_ffmpeg = check_ffmpeg()
    if has_ffmpeg:
        status_bg, status_fg, status_txt = "#4CAF50", "white", "环境正常: 已检测到 FFmpeg"
    else:
        status_bg, status_fg, status_txt = "#FF5722", "white", "环境错误: 未检测到 ffmpeg.exe！"

    tk.Label(top, text=status_txt, bg=status_bg, fg=status_fg, font=("Arial", 10, "bold")).pack(fill="x")

    # === 1. 文件选择 ===
    tk.Label(top, text="源图片路径 (必须是透明 PNG):").pack(anchor="w", **pad_opts)
    entry_in = tk.Entry(top)
    entry_in.pack(fill="x", padx=10)
    tk.Button(top, text="浏览图片",
              command=lambda: (entry_in.delete(0, tk.END), entry_in.insert(0, filedialog.askopenfilename(parent=top,
                                                                                                         filetypes=[
                                                                                                             ("Images",
                                                                                                              "*.png")])))).pack(
        anchor="e", padx=10)

    tk.Label(top, text="输出视频路径 (.mov):").pack(anchor="w", **pad_opts)
    entry_out = tk.Entry(top)
    entry_out.pack(fill="x", padx=10)

    def sel_out():
        path = filedialog.asksaveasfilename(defaultextension=".mov", filetypes=[("Video", "*.mov")], parent=top)
        if path:
            entry_out.delete(0, tk.END)
            entry_out.insert(0, path)

    tk.Button(top, text="浏览保存", command=sel_out).pack(anchor="e", padx=10)

    tk.Frame(top, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    # === 2. 视频参数 ===
    frame_param = tk.Frame(top)
    frame_param.pack(fill="x", padx=10)

    tk.Label(frame_param, text="画布尺寸:").grid(row=0, column=0, sticky="w", pady=5)
    var_res = tk.StringVar(value="1080p")
    tk.Radiobutton(frame_param, text="1080P", variable=var_res, value="1080p").grid(row=0, column=1, sticky="w")
    tk.Radiobutton(frame_param, text="4K", variable=var_res, value="4k").grid(row=0, column=2, sticky="w")

    tk.Label(frame_param, text="帧率 (FPS):").grid(row=1, column=0, sticky="w", pady=5)
    var_fps = tk.IntVar(value=60)
    tk.Radiobutton(frame_param, text="60 FPS", variable=var_fps, value=60).grid(row=1, column=1, sticky="w")
    tk.Radiobutton(frame_param, text="120 FPS", variable=var_fps, value=120).grid(row=1, column=2, sticky="w")

    # === GPU 加速选项 ===
    tk.Label(frame_param, text="编码模式:").grid(row=2, column=0, sticky="w", pady=5)
    var_gpu = tk.BooleanVar(value=False)
    chk_gpu = tk.Checkbutton(frame_param, text="尝试使用 NVIDIA 显卡加速 (HEVC)", variable=var_gpu, fg="#E91E63")
    chk_gpu.grid(row=2, column=1, columnspan=2, sticky="w")

    tk.Frame(top, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    # === 3. 运动参数 ===
    frame_move = tk.Frame(top)
    frame_move.pack(fill="x", padx=10)

    tk.Label(frame_move, text="运动方向 (角度):").grid(row=0, column=0, sticky="w", pady=5)
    entry_angle = tk.Entry(frame_move, width=10);
    entry_angle.insert(0, "0");
    entry_angle.grid(row=0, column=1, sticky="w")
    tk.Label(frame_move, text="(0=右, 90=下)", fg="gray").grid(row=0, column=2, sticky="w", padx=5)

    tk.Label(frame_move, text="移动总距离 (px):").grid(row=1, column=0, sticky="w", pady=5)
    entry_dist = tk.Entry(frame_move, width=10);
    entry_dist.insert(0, "500");
    entry_dist.grid(row=1, column=1, sticky="w")

    tk.Label(frame_move, text="速度 (px/秒):").grid(row=2, column=0, sticky="w", pady=5)
    entry_speed = tk.Entry(frame_move, width=10);
    entry_speed.insert(0, "100");
    entry_speed.grid(row=2, column=1, sticky="w")

    # === 说明 ===
    tk.Label(top,
             text="说明：\n1. CPU模式: ProRes 4444 (兼容性最佳)。\n2. 显卡模式: HEVC (速度快，需N卡)。\n3. 如果卡在99%，请耐心等待文件封包完成。",
             fg="#2196F3", bg="#E3F2FD", justify="left", bd=1, relief="groove").pack(fill="x", padx=10, pady=10,
                                                                                     ipady=5)

    # === 进度条 ===
    progress_bar = ttk.Progressbar(top, orient="horizontal", length=400, mode="determinate")
    progress_bar.pack(fill="x", padx=10, pady=5)
    lbl_status = tk.Label(top, text="准备就绪", fg="gray")
    lbl_status.pack()

    # === 执行按钮 ===
    def run():
        f_in = entry_in.get()
        f_out = entry_out.get()
        if not f_in or not f_out:
            messagebox.showwarning("提示", "请选择输入和输出路径", parent=top)
            return

        try:
            ang = float(entry_angle.get())
            dist = float(entry_dist.get())
            spd = float(entry_speed.get())
            if spd <= 0: raise ValueError
        except:
            messagebox.showerror("错误", "参数输入有误", parent=top)
            return

        res_map = {"1080p": (1920, 1080), "4k": (3840, 2160)}
        target_res = res_map[var_res.get()]
        target_fps = var_fps.get()
        use_gpu_mode = var_gpu.get()

        mode_text = "GPU加速 (HEVC)" if use_gpu_mode else "CPU (ProRes)"
        btn_run.config(state="disabled", text=f"正在渲染 [{mode_text}]...")
        progress_bar['value'] = 0
        lbl_status.config(text="初始化 FFmpeg...", fg="blue")

        def update_prog(val):
            top.after(0, lambda: progress_bar.configure(value=val))
            top.after(0, lambda: lbl_status.config(text=f"渲染中: {val:.1f}%"))

        def on_done(error_msg, final_path):
            top.after(0, lambda: _finish_ui(error_msg, final_path))

        def _finish_ui(error_msg, final_path):
            btn_run.config(state="normal", text="开始生成 MOV")
            if error_msg:
                lbl_status.config(text="失败", fg="red")
                messagebox.showerror("渲染错误", error_msg, parent=top)
            else:
                lbl_status.config(text="完成！", fg="green")
                messagebox.showinfo("成功", f"视频生成成功！\n路径: {final_path}\n模式: {mode_text}", parent=top)
                try:
                    os.startfile(final_path)
                except:
                    pass

        t = threading.Thread(target=_render_video_thread, args=(
            f_in, f_out, target_res, target_fps, ang, dist, spd, use_gpu_mode, update_prog, on_done
        ))
        t.daemon = True
        t.start()

    btn_run = tk.Button(top, text="开始生成 MOV", bg="#9C27B0", fg="white", font=("Arial", 12, "bold"),
                        command=run)
    btn_run.pack(pady=10, fill="x", padx=20)