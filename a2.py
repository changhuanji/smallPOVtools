import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import os
import platform
import re
import cv2
import numpy as np

try:
    import win32com.client

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ==============================================================================
# 功能模块 1: PPT 逐页导出 (全新 Slide.Export 方式)
# ==============================================================================
def _process_export_transparent_png(input_path, output_dir, dpi_str, target_ratio_mode,
                                    enable_exp, exp_w, exp_h, exp_anchor):
    # 1. 环境检查
    if platform.system() != 'Windows':
        messagebox.showerror("系统不支持", "PPT 导出功能仅支持 Windows 系统。")
        return
    if not HAS_WIN32:
        messagebox.showerror("缺少依赖", "请安装: pip install pywin32")
        return
    if not input_path or not output_dir:
        messagebox.showwarning("提示", "路径不能为空！")
        return

    # 2. 解析 DPI
    target_dpi = 216  # 默认值
    try:
        nums = re.findall(r"\d+", str(dpi_str))
        if nums:
            target_dpi = int(nums[0])
        if target_dpi < 30 or target_dpi > 3000:
            raise ValueError
    except:
        messagebox.showerror("错误", "DPI 必须是 30 到 3000 之间的整数。")
        return

    # 3. 检查实验性参数
    final_target_w = 0
    final_target_h = 0
    if enable_exp:
        try:
            final_target_w = int(exp_w)
            final_target_h = int(exp_h)
            if final_target_w <= 0 or final_target_h <= 0: raise ValueError
        except:
            messagebox.showerror("输入错误", "实验性功能的宽/高必须为正整数。")
            return

    abs_output = os.path.abspath(output_dir)
    if not os.path.exists(abs_output):
        try:
            os.makedirs(abs_output)
        except Exception as e:
            messagebox.showerror("路径错误", f"无法创建输出文件夹:\n{e}")
            return

    abs_input = os.path.abspath(input_path)
    ppt_app = None
    pres = None

    try:
        ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
        ppt_app.Visible = True
        ppt_app.WindowState = 2  # 最小化

        pres = ppt_app.Presentations.Open(abs_input, WithWindow=False)

        slide_w_points = pres.PageSetup.SlideWidth
        slide_h_points = pres.PageSetup.SlideHeight

        # === 计算目标分辨率 ===
        scale_factor = target_dpi / 72.0
        base_px_w = int(slide_w_points * scale_factor)
        base_px_h = int(slide_h_points * scale_factor)

        # 处理比例
        export_h = base_px_h
        if target_ratio_mode == "16:9":
            export_h = int(base_px_w * 9 / 16)
        elif target_ratio_mode == "4:3":
            export_h = int(base_px_w * 3 / 4)
        elif target_ratio_mode == "1:1":
            export_h = base_px_w

        total_slides = pres.Slides.Count
        success_count = 0

        for i in range(1, total_slides + 1):
            slide = pres.Slides(i)
            save_name = f"slide_{i}.png"
            save_full_path = os.path.join(abs_output, save_name)

            # 记录原始背景状态 (以便恢复)
            orig_follow = slide.FollowMasterBackground
            orig_visible = slide.Background.Fill.Visible

            try:
                # === 核心修改：使用 Slide.Export 替代 ShapeRange ===
                # 这种方法分辨率极其精准，不会出现偏差

                # 1. 尝试强制设置背景透明
                # 注意：这需要 PPT 设置支持，部分版本可能依旧输出白底
                # 如果用户 PPT 母版有图片背景，这里可能无法去除，建议用户在 PPT 里删掉背景图
                slide.FollowMasterBackground = 0  # 不跟随母版
                slide.Background.Fill.Visible = 0  # 背景不可见 (透明)
                slide.Background.Fill.Transparency = 1.0  # 100% 透明

                # 2. 原生导出
                # Export(FileName, FilterName, ScaleWidth, ScaleHeight)
                # 直接传入计算好的整数宽高
                slide.Export(save_full_path, "PNG", int(base_px_w), int(export_h))

                # 3. === 实验性功能：OpenCV 二次处理 ===
                if enable_exp:
                    img = cv2.imdecode(np.fromfile(save_full_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
                    if img is not None:
                        if len(img.shape) == 2:
                            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
                        elif img.shape[2] == 3:
                            img = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)

                        h, w = img.shape[:2]

                        if w != final_target_w or h != final_target_h:
                            canvas = np.zeros((final_target_h, final_target_w, 4), dtype=np.uint8)

                            x_offset = 0
                            y_offset = 0

                            if "左" in exp_anchor:
                                x_offset = 0
                            elif "右" in exp_anchor:
                                x_offset = final_target_w - w
                            else:
                                x_offset = (final_target_w - w) // 2

                            if "上" in exp_anchor:
                                y_offset = 0
                            elif "下" in exp_anchor:
                                y_offset = final_target_h - h
                            else:
                                y_offset = (final_target_h - h) // 2

                            x1_c = max(0, x_offset)
                            y1_c = max(0, y_offset)
                            x2_c = min(final_target_w, x_offset + w)
                            y2_c = min(final_target_h, y_offset + h)

                            x1_img = max(0, -x_offset)
                            y1_img = max(0, -y_offset)

                            w_slice = x2_c - x1_c
                            h_slice = y2_c - y1_c

                            if w_slice > 0 and h_slice > 0:
                                canvas[y1_c:y2_c, x1_c:x2_c] = img[y1_img:y1_img + h_slice, x1_img:x1_img + w_slice]

                            is_success, buffer = cv2.imencode(".png", canvas)
                            if is_success:
                                buffer.tofile(save_full_path)

                success_count += 1

            except Exception as e:
                print(f"Page {i} error: {e}")

            finally:
                # 恢复背景设置 (以免用户保存 PPT 后发现背景没了)
                try:
                    slide.FollowMasterBackground = orig_follow
                    slide.Background.Fill.Visible = orig_visible
                except:
                    pass

        msg = f"导出成功: {success_count}/{total_slides} 页\n保存位置: {abs_output}"
        if enable_exp:
            msg += f"\n\n已强制调整为: {final_target_w}x{final_target_h}"
        else:
            msg += f"\n\n当前DPI: {target_dpi} (尺寸 {base_px_w}x{export_h})"

        messagebox.showinfo("完成", msg)
        try:
            os.startfile(abs_output)
        except:
            pass

    except Exception as e:
        messagebox.showerror("错误", f"发生错误：\n{str(e)}")
    finally:
        if pres: pres.Close()
        if ppt_app:
            try:
                ppt_app.Quit()
            except:
                pass


# ==============================================================================
# 功能模块 2: 批量裁切/扩展
# ==============================================================================
def _process_batch_crop_extend(folder_path, val_top, val_bottom, val_left, val_right):
    if not folder_path or not os.path.exists(folder_path):
        messagebox.showerror("错误", "请选择有效的文件夹！")
        return
    try:
        v_t, v_b, v_l, v_r = int(val_top), int(val_bottom), int(val_left), int(val_right)
    except:
        messagebox.showerror("错误", "裁切数值必须是整数。")
        return

    valid_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff')
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(valid_exts)]
    if not files:
        messagebox.showwarning("提示", "没有找到图片。")
        return

    success_count = 0
    for filename in files:
        file_path = os.path.join(folder_path, filename)
        try:
            img = cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
            if img is None: continue
            if len(img.shape) == 2:
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
            elif img.shape[2] == 3:
                img = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)

            h, w = img.shape[:2]
            crop_t, crop_b = max(0, v_t), max(0, v_b)
            crop_l, crop_r = max(0, v_l), max(0, v_r)

            if (crop_t + crop_b >= h) or (crop_l + crop_r >= w): continue

            if crop_t > 0 or crop_b > 0 or crop_l > 0 or crop_r > 0:
                end_y = -crop_b if crop_b > 0 else None
                end_x = -crop_r if crop_r > 0 else None
                img = img[crop_t: end_y, crop_l: end_x]

            pad_t, pad_b = abs(min(0, v_t)), abs(min(0, v_b))
            pad_l, pad_r = abs(min(0, v_l)), abs(min(0, v_r))

            if pad_t > 0 or pad_b > 0 or pad_l > 0 or pad_r > 0:
                img = cv2.copyMakeBorder(img, pad_t, pad_b, pad_l, pad_r, cv2.BORDER_CONSTANT, value=(0, 0, 0, 0))

            name_part, ext_part = os.path.splitext(filename)
            if ext_part.lower() == '.png':
                is_success, buffer = cv2.imencode(".png", img)
                if is_success: buffer.tofile(file_path); success_count += 1
            else:
                new_path = os.path.join(folder_path, name_part + ".png")
                is_success, buffer = cv2.imencode(".png", img)
                if is_success: buffer.tofile(new_path); os.remove(file_path); success_count += 1
        except Exception as e:
            print(f"Error {filename}: {e}")
    messagebox.showinfo("完成", f"批量处理完成！共 {success_count} 张。")


# ==============================================================================
# 功能模块 3: 图片批量去底
# ==============================================================================
def _process_batch_remove_bg(folder_path, rgb_str, tolerance):
    if not folder_path or not os.path.exists(folder_path):
        messagebox.showerror("错误", "请选择有效的文件夹！")
        return
    try:
        rgb_clean = rgb_str.replace("，", ",").replace(" ", "")
        r, g, b = map(int, rgb_clean.split(','))
    except:
        messagebox.showerror("错误", "RGB 格式错误。")
        return

    files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.jpg', '.png', '.jpeg', '.bmp'))]
    if not files:
        messagebox.showwarning("提示", "没有找到图片。")
        return

    success_count = 0
    for filename in files:
        file_path = os.path.join(folder_path, filename)
        try:
            img = cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
            if img is None: continue
            if len(img.shape) == 2:
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
            elif img.shape[2] == 3:
                img = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)

            lower = np.array([max(0, b - tolerance), max(0, g - tolerance), max(0, r - tolerance), 0])
            upper = np.array([min(255, b + tolerance), min(255, g + tolerance), min(255, r + tolerance), 255])

            bgr = img[:, :, :3]
            lower_bgr = np.array([max(0, b - tolerance), max(0, g - tolerance), max(0, r - tolerance)])
            upper_bgr = np.array([min(255, b + tolerance), min(255, g + tolerance), min(255, r + tolerance)])
            mask = cv2.inRange(bgr, lower_bgr, upper_bgr)

            img[:, :, 3] = np.where(mask == 255, 0, img[:, :, 3])

            name_part, ext_part = os.path.splitext(filename)
            if ext_part.lower() == '.png':
                is_success, buffer = cv2.imencode(".png", img)
                if is_success: buffer.tofile(file_path); success_count += 1
            else:
                new_path = os.path.join(folder_path, name_part + ".png")
                is_success, buffer = cv2.imencode(".png", img)
                if is_success: buffer.tofile(new_path); os.remove(file_path); success_count += 1
        except Exception as e:
            print(f"Error {filename}: {e}")
    messagebox.showinfo("完成", f"去底完成！共 {success_count} 张。")


# ==============================================================================
# UI 界面逻辑
# ==============================================================================
def show_ui(parent):
    top = tk.Toplevel(parent)
    top.title("图片处理工具箱")
    top.geometry("520x620")
    top.transient(parent)
    top.grab_set()

    top.update_idletasks()
    x = (top.winfo_screenwidth() - top.winfo_width()) // 2
    y = (top.winfo_screenheight() - top.winfo_height()) // 2
    top.geometry(f"+{x}+{y}")

    notebook = ttk.Notebook(top)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    # Tab 1
    tab1 = tk.Frame(notebook)
    notebook.add(tab1, text="1. PPT 逐页导出")
    _init_ppt_ui(tab1, top)

    # Tab 2
    tab2 = tk.Frame(notebook)
    notebook.add(tab2, text="2. 批量裁切/扩展")
    _init_crop_extend_ui(tab2, top)

    # Tab 3
    tab3 = tk.Frame(notebook)
    notebook.add(tab3, text="3. 批量图片去背景")
    _init_bg_remove_ui(tab3, top)


# -----------------------------------------------------------
# Tab 1 UI
# -----------------------------------------------------------
def _init_ppt_ui(frame, parent_win):
    pad_opts = {'padx': 10, 'pady': 5}

    def select_file(entry):
        path = filedialog.askopenfilename(filetypes=[("PPT", "*.pptx;*.ppt")], parent=parent_win)
        if path: entry.delete(0, tk.END); entry.insert(0, path)

    def select_dir(entry):
        path = filedialog.askdirectory(parent=parent_win)
        if path: entry.delete(0, tk.END); entry.insert(0, path)

    tk.Label(frame, text="选择 PPT 文件:").pack(anchor="w", **pad_opts)
    entry_in = tk.Entry(frame);
    entry_in.pack(fill="x", **pad_opts)
    tk.Button(frame, text="浏览", command=lambda: select_file(entry_in)).pack(anchor="e", padx=10)

    tk.Label(frame, text="输出文件夹:").pack(anchor="w", **pad_opts)
    entry_out = tk.Entry(frame);
    entry_out.pack(fill="x", **pad_opts)
    tk.Button(frame, text="浏览", command=lambda: select_dir(entry_out)).pack(anchor="e", padx=10)

    tk.Frame(frame, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    tk.Label(frame, text="DPI 设置 (清晰度):").pack(anchor="w", padx=10, pady=(5, 0))
    dpi_values = ["72 (屏幕)", "96", "150", "288", "300 (打印)", "600"]
    combo_dpi = ttk.Combobox(frame, values=dpi_values, width=15)
    combo_dpi.current(3)  # 默认 216
    combo_dpi.pack(anchor="w", padx=10, pady=2)

    tk.Label(frame, text="强制比例 (PPT导出阶段):").pack(anchor="w", padx=10, pady=(5, 0))
    combo_ratio = ttk.Combobox(frame, values=["原比例", "16:9", "4:3"], width=15, state="readonly")
    combo_ratio.current(1);
    combo_ratio.pack(anchor="w", padx=10, pady=2)

    tk.Frame(frame, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    # === 实验性功能 ===
    exp_frame = tk.LabelFrame(frame, text="实验性功能：强制裁切/填充", fg="#E91E63")
    exp_frame.pack(fill="x", padx=10, pady=5)

    var_enable = tk.BooleanVar(value=False)

    def toggle_exp():
        state = "normal" if var_enable.get() else "disabled"
        e_w.config(state=state)
        e_h.config(state=state)
        cb_anchor.config(state=state)

    chk_exp = tk.Checkbutton(exp_frame, text="启用强制分辨率修正 (二次处理)", variable=var_enable, command=toggle_exp)
    chk_exp.pack(anchor="w", padx=5)

    grid_f = tk.Frame(exp_frame)
    grid_f.pack(fill="x", padx=10, pady=5)

    tk.Label(grid_f, text="目标宽:").grid(row=0, column=0)
    e_w = tk.Entry(grid_f, width=8);
    e_w.insert(0, "3840");
    e_w.grid(row=0, column=1, padx=5)

    tk.Label(grid_f, text="目标高:").grid(row=0, column=2)
    e_h = tk.Entry(grid_f, width=8);
    e_h.insert(0, "2160");
    e_h.grid(row=0, column=3, padx=5)

    tk.Label(grid_f, text="锚点(保留):").grid(row=1, column=0, pady=5)
    anchor_vals = ["左上 (裁剪右/下)", "居中 (裁剪四周)", "左下 (裁剪右/上)", "右上 (裁剪左/下)", "右下 (裁剪左/上)"]
    cb_anchor = ttk.Combobox(grid_f, values=anchor_vals, state="readonly", width=18)
    cb_anchor.current(0)
    cb_anchor.grid(row=1, column=1, columnspan=3, sticky="w", padx=5)

    toggle_exp()

    def run():
        _process_export_transparent_png(
            entry_in.get(), entry_out.get(), combo_dpi.get(), combo_ratio.get(),
            var_enable.get(), e_w.get(), e_h.get(), cb_anchor.get()
        )

    tk.Button(frame, text="开始导出", bg="#FF9800", fg="white", font=("Arial", 12, "bold"), command=run).pack(pady=20,
                                                                                                              fill="x",
                                                                                                              padx=20)


# -----------------------------------------------------------
# Tab 2 UI (批量裁切/扩展)
# -----------------------------------------------------------
def _init_crop_extend_ui(frame, parent_win):
    pad_opts = {'padx': 10, 'pady': 8}

    def select_dir(entry):
        path = filedialog.askdirectory(parent=parent_win)
        if path: entry.delete(0, tk.END); entry.insert(0, path)

    tk.Label(frame, text="选择图片文件夹:").pack(anchor="w", **pad_opts)
    entry_folder = tk.Entry(frame);
    entry_folder.pack(fill="x", **pad_opts)
    tk.Button(frame, text="浏览文件夹", command=lambda: select_dir(entry_folder)).pack(anchor="e", padx=10)

    tk.Frame(frame, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    tk.Label(frame, text="输入像素值 (上 / 下 / 左 / 右):", font=("Arial", 10, "bold")).pack(anchor="w", padx=10)

    hint_frame = tk.LabelFrame(frame, text="规则说明", fg="gray")
    hint_frame.pack(fill="x", padx=10, pady=5)
    tk.Label(hint_frame, text="正数 (+): 向内裁切 (图片变小)\n负数 (-): 向外扩展 (图片变大，填充透明像素)",
             justify="left", fg="#2196F3").pack(anchor="w", padx=5, pady=5)

    grid_frame = tk.Frame(frame)
    grid_frame.pack(fill="x", padx=20, pady=10)

    tk.Label(grid_frame, text="Top (上):").grid(row=0, column=0, sticky="e", pady=5)
    e_top = tk.Entry(grid_frame, width=8);
    e_top.insert(0, "0");
    e_top.grid(row=0, column=1, padx=5)
    tk.Label(grid_frame, text="Bottom (下):").grid(row=0, column=2, sticky="e", pady=5)
    e_btm = tk.Entry(grid_frame, width=8);
    e_btm.insert(0, "0");
    e_btm.grid(row=0, column=3, padx=5)
    tk.Label(grid_frame, text="Left (左):").grid(row=1, column=0, sticky="e", pady=5)
    e_lft = tk.Entry(grid_frame, width=8);
    e_lft.insert(0, "0");
    e_lft.grid(row=1, column=1, padx=5)
    tk.Label(grid_frame, text="Right (右):").grid(row=1, column=2, sticky="e", pady=5)
    e_rgt = tk.Entry(grid_frame, width=8);
    e_rgt.insert(0, "0");
    e_rgt.grid(row=1, column=3, padx=5)

    tk.Label(frame, text="警告：直接覆盖源文件！\n若扩展，JPG将转为PNG。", fg="#E91E63", bg="#FCE4EC", justify="left", bd=1,
             relief="groove").pack(fill="x", padx=10, pady=20, ipady=5)

    def run():
        _process_batch_crop_extend(entry_folder.get(), e_top.get(), e_btm.get(), e_lft.get(), e_rgt.get())

    tk.Button(frame, text="开始批量处理", bg="#673AB7", fg="white", font=("Arial", 12, "bold"), command=run).pack(
        side="bottom", pady=20, fill="x", padx=20)


# -----------------------------------------------------------
# Tab 3 UI (去底)
# -----------------------------------------------------------
def _init_bg_remove_ui(frame, parent_win):
    pad_opts = {'padx': 10, 'pady': 8}

    def select_dir(entry):
        path = filedialog.askdirectory(parent=parent_win)
        if path: entry.delete(0, tk.END); entry.insert(0, path)

    tk.Label(frame, text="选择图片文件夹:").pack(anchor="w", **pad_opts)
    entry_folder = tk.Entry(frame);
    entry_folder.pack(fill="x", **pad_opts)
    tk.Button(frame, text="浏览文件夹", command=lambda: select_dir(entry_folder)).pack(anchor="e", padx=10)

    tk.Frame(frame, height=2, bd=1, relief="sunken").pack(fill="x", padx=10, pady=10)

    tk.Label(frame, text="背景色 (RGB):").pack(anchor="w", padx=10)
    frame_c = tk.Frame(frame);
    frame_c.pack(fill="x", padx=10)
    entry_rgb = tk.Entry(frame_c, width=15);
    entry_rgb.insert(0, "255,255,255");
    entry_rgb.pack(side="left")
    tk.Label(frame_c, text="(如 0,0,0 黑色)", fg="gray").pack(side="left", padx=10)

    tk.Label(frame, text="容差 (0-100):").pack(anchor="w", padx=10, pady=(10, 0))
    scale_tol = tk.Scale(frame, from_=0, to=100, orient="horizontal");
    scale_tol.set(10);
    scale_tol.pack(fill="x", padx=20, pady=5)

    tk.Label(frame, text="警告：覆盖源文件！JPG转PNG。", fg="#E91E63", bg="#FCE4EC", justify="left", bd=1,
             relief="groove").pack(fill="x", padx=10, pady=20)

    def run():
        _process_batch_remove_bg(entry_folder.get(), entry_rgb.get(), scale_tol.get())

    tk.Button(frame, text="开始去底", bg="#2196F3", fg="white", font=("Arial", 12, "bold"), command=run).pack(
        side="bottom", pady=20, fill="x", padx=20)