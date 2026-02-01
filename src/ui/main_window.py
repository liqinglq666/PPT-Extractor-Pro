import os
import sys
import threading
import time
import re
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
import cv2
import numpy as np

# Internal utility imports
from src.utils.file_ops import cv2_imread_safe, cv2_imwrite_safe
from src.utils.time_ops import parse_time, format_time
from src.ui.dialogs import VideoCutterDialog


class PPTExtractorEngine(tb.Window):
    """
    PPT Extractor Professional - Main GUI Engine.
    Fixes:
    1. ROI SYNC: Monitor and Capture UIs now show the CROPPED image, not the full frame.
    2. UI LAYOUT: Increased dimensions and adjusted splitter for better visibility.
    """

    def __init__(self):
        super().__init__(themename="cosmo")

        self.title("PPT Extractor Professional")
        # [UI Fix] 增加默认宽高，适配更多内容
        self.geometry("1500x950")
        self.minsize(1200, 800)
        self._init_assets()

        # --- Configuration Variables ---
        self.video_path = tb.StringVar()
        self.output_path = tb.StringVar()
        self.project_name = tb.StringVar(value=f"Lecture_{int(time.time())}")

        # Algorithmic Parameters
        self.diff_threshold = tb.IntVar(value=10)
        self.stability_frames = tb.IntVar(value=5)
        self.check_interval = tb.DoubleVar(value=0.5)

        # State Vectors
        self.status_msg = tb.StringVar(value="SYSTEM READY / 待命")
        self.log_msg = tb.StringVar(value="Core services initialized.")
        self.var_processed = tb.StringVar(value="00:00:00")
        self.var_captured = tb.StringVar(value="0")
        self.progress_var = tb.IntVar(value=0)

        # Feature Flags
        self.monitor_on = tb.BooleanVar(value=True)
        self.make_pdf = tb.BooleanVar(value=True)
        self.remove_borders = tb.BooleanVar(value=True)
        self.high_precision = tb.BooleanVar(value=False)

        # Runtime State
        self.roi_rect = None
        self.is_time_locked = False
        self.is_running = False
        self._thread_lock = threading.Lock()

        self._init_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _init_assets(self):
        try:
            base_path = sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.abspath(".")
            icon_path = os.path.join(base_path, "assets", "app_main.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

    def log(self, text):
        ts = time.strftime("%H:%M:%S")
        self.after(0, lambda: self.log_msg.set(f"[{ts}] {text}"))

    def set_status(self, status_text, color="white"):
        def _update():
            self.status_msg.set(status_text)

        self.after(0, _update)

    def on_closing(self):
        if self.is_running:
            if messagebox.askokcancel("Quit", "Engine is running. Force quit?"):
                self.is_running = False
                self.destroy()
        else:
            self.destroy()

    def _init_ui(self):
        st = tb.Style()
        st.configure("Custom.Horizontal.TProgressbar", thickness=10)

        # Header
        header = tb.Frame(self, bootstyle="primary")
        header.pack(side=TOP, fill=X)
        h_content = tb.Frame(header, padding=(15, 10), bootstyle="primary")
        h_content.pack(fill=X)

        tb.Label(h_content, text="PPT EXTRACTOR PRO", font=("Impact", 24), foreground="white",
                 bootstyle="inverse-primary").pack(side=LEFT)
        tb.Label(h_content, text="Automated Slide Extraction System", font=("Segoe UI", 10, "italic"),
                 foreground="#e0e0e0", bootstyle="inverse-primary").pack(side=RIGHT, pady=5)

        # Footer
        footer = tb.Frame(self, bootstyle="secondary", padding=(10, 5))
        footer.pack(side=BOTTOM, fill=X)
        footer.columnconfigure(1, weight=1)

        f_stat = tb.Frame(footer, bootstyle="secondary")
        f_stat.grid(row=0, column=0, sticky="w")
        tb.Label(f_stat, text="STATUS:", font=("Arial", 8, "bold"), bootstyle="inverse-secondary").pack(side=LEFT)
        tb.Label(f_stat, textvariable=self.status_msg, font=("Segoe UI", 9, "bold"), foreground="#00ff00",
                 bootstyle="inverse-secondary").pack(side=LEFT, padx=5)

        f_center = tb.Frame(footer, bootstyle="secondary")
        f_center.grid(row=0, column=1, padx=20, sticky="ew")
        self.progress = tb.Progressbar(f_center, bootstyle="primary-striped", variable=self.progress_var,
                                       mode='determinate', maximum=100, style="Custom.Horizontal.TProgressbar")
        self.progress.pack(fill=X, pady=(2, 2))
        tb.Label(f_center, textvariable=self.log_msg, font=("Consolas", 8), foreground="#dcdcdc",
                 bootstyle="inverse-secondary").pack(anchor=W)

        f_count = tb.Frame(footer, bootstyle="secondary")
        f_count.grid(row=0, column=2, sticky="e")
        tb.Label(f_count, text="CAPTURED:", font=("Arial", 7), bootstyle="inverse-secondary").pack(side=LEFT)
        tb.Label(f_count, textvariable=self.var_captured, font=("Consolas", 11, "bold"), foreground="white",
                 bootstyle="inverse-secondary").pack(side=LEFT, padx=(2, 0))

        # Splitter
        splitter = ttk.PanedWindow(self, orient=HORIZONTAL)
        splitter.pack(side=TOP, fill=BOTH, expand=True, padx=10, pady=10)

        left_container = tb.Frame(splitter)
        right_panel = tb.Frame(splitter)
        splitter.add(left_container, weight=0)
        splitter.add(right_panel, weight=1)

        # [UI Fix] Give sidebar more space by default (520px)
        self.after(200, lambda: splitter.sashpos(0, 520))

        # Left Panel (Fixed Button + Scrollable Settings)
        self.btn_run = tb.Button(left_container, text="INITIALIZE ENGINE / 启动抽取引擎", command=self.toggle_run,
                                 bootstyle="primary", padding=15)
        self.btn_run.pack(side=BOTTOM, fill=X, pady=(10, 0), padx=5)

        # Scrollable Sidebar Setup
        canvas = tk.Canvas(left_container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tb.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def _configure_canvas(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", _configure_canvas)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        left_container.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        left_container.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self._build_sidebar(scrollable_frame)
        self._build_viewports(right_panel)

    def _build_sidebar(self, parent):
        # Reduced padding to fit more content
        c1 = tb.Labelframe(parent, text=" [1] Data Source / 数据配置 ", padding=8, bootstyle="primary")
        c1.pack(fill=X, pady=(0, 8), padx=5)

        r1 = tb.Frame(c1);
        r1.pack(fill=X, pady=2)
        tb.Label(r1, text="视频路径:").pack(side=LEFT)
        tb.Entry(r1, textvariable=self.video_path).pack(side=LEFT, fill=X, expand=True, padx=5)
        tb.Button(r1, text="📂", command=self.select_video, bootstyle="primary-outline", width=4).pack(side=LEFT)

        r2 = tb.Frame(c1);
        r2.pack(fill=X, pady=2)
        tb.Label(r2, text="输出目录:").pack(side=LEFT)
        tb.Entry(r2, textvariable=self.output_path).pack(side=LEFT, fill=X, expand=True, padx=5)
        tb.Button(r2, text="📂", command=self.select_output, bootstyle="primary-outline", width=4).pack(side=LEFT)

        r3 = tb.Frame(c1);
        r3.pack(fill=X, pady=2)
        tb.Label(r3, text="项目命名:", foreground="#007bff").pack(side=LEFT)
        tb.Entry(r3, textvariable=self.project_name).pack(side=LEFT, fill=X, expand=True, padx=5)
        tb.Label(r3, text="(自动建文件夹)", font=("Arial", 7), foreground="#999").pack(side=RIGHT)

        c2 = tb.Labelframe(parent, text=" [2] Timeline / 时域控制 ", padding=8, bootstyle="primary")
        c2.pack(fill=X, pady=(0, 8), padx=5)

        tb.Button(c2, text="✂️ 启动可视化时间轴裁剪 (Visual Cutter)", command=self.open_video_cutter,
                  bootstyle="outline-primary").pack(fill=X, pady=(0, 5))

        t_head = tb.Frame(c2);
        t_head.pack(fill=X)
        tb.Label(t_head, text="Start / 起点", font=("Segoe UI", 8), foreground="#999").pack(side=LEFT, expand=True)
        tb.Label(t_head, text="End / 终点", font=("Segoe UI", 8), foreground="#999").pack(side=LEFT, expand=True)

        t_grid = tb.Frame(c2);
        t_grid.pack(fill=X)
        self.ent_start = tb.Entry(t_grid, justify="center", font=("Consolas", 11, "bold"),
                                  state="readonly", foreground="#28a745")
        self.ent_start.pack(side=LEFT, fill=X, expand=True, padx=(0, 2))

        self.ent_end = tb.Entry(t_grid, justify="center", font=("Consolas", 11, "bold"),
                                state="readonly", foreground="#fd7e14")
        self.ent_end.pack(side=LEFT, fill=X, expand=True, padx=(2, 0))

        sw_f = tb.Frame(parent, padding=5);
        sw_f.pack(fill=X, pady=(0, 5), padx=5)
        tb.Checkbutton(sw_f, text="自动生成 PDF", variable=self.make_pdf, bootstyle="primary-round-toggle").pack(
            side=LEFT, padx=5)
        tb.Checkbutton(sw_f, text="智能去黑边", variable=self.remove_borders, bootstyle="primary-round-toggle").pack(
            side=RIGHT, padx=5)

        c3 = tb.Labelframe(parent, text=" [3] Visual Kernel / 视觉算子 ", padding=8, bootstyle="primary")
        c3.pack(fill=X, padx=5)

        row_roi = tb.Frame(c3);
        row_roi.pack(fill=X, pady=(0, 10))
        tb.Button(row_roi, text="[+] 设定 ROI 扫描区域", command=self.set_roi, bootstyle="primary-outline",
                  width=20).pack(side=LEFT)
        self.lbl_roi_status = tb.Label(row_roi, text="全屏扫描", font=("Segoe UI", 8), foreground="#999")
        self.lbl_roi_status.pack(side=RIGHT, padx=5)

        tb.Separator(c3, orient=HORIZONTAL, bootstyle="light").pack(fill=X, pady=5)

        self._add_param_row(c3, "判定阈值:", self.diff_threshold, "12", "越小越灵敏 (5-20)", 0, min_val=2, max_val=50)
        self._add_param_row(c3, "扫描步长:", self.check_interval, "0.5", "步长越短越精准 (0.3-1.0)", 1, is_spin=True)
        self._add_param_row(c3, "防抖等级:", self.stability_frames, "5", "连续稳定多少帧才抓取 (3-10)", 2,
                            is_spin=False, min_val=2, max_val=20)

    def _add_param_row(self, master, label, var, suggest, note, row, is_spin=False, min_val=5, max_val=30):
        f = tb.Frame(master);
        f.pack(fill=X, pady=4)
        top = tb.Frame(f);
        top.pack(fill=X)
        tb.Label(top, text=label, width=12, font=("Segoe UI", 9, "bold")).pack(side=LEFT)
        val_label = tb.Label(top, text=f"[{var.get()}]", font=("Arial", 9, "bold"), foreground="#007bff")
        val_label.pack(side=RIGHT, padx=5)

        def _on_scale_change(v):
            val = float(v)
            if not is_spin:
                val = int(val)
            else:
                val = round(val, 1)
            var.set(val)
            val_label.config(text=f"[{val}]")

        if is_spin:
            s = tb.Spinbox(top, textvariable=var, from_=0.1, to=10.0, increment=0.1, width=8,
                           command=lambda: val_label.config(text=f"[{round(var.get(), 1)}]"))
            s.pack(side=LEFT, padx=5)
        else:
            s = tb.Scale(top, variable=var, from_=min_val, to=max_val, bootstyle="primary", command=_on_scale_change)
            s.pack(side=LEFT, fill=X, expand=True, padx=5)
        tb.Label(f, text=f"说明: {note}", font=("Segoe UI", 8), foreground="#999").pack(anchor=W)

    def _build_viewports(self, parent):
        m_frame = tb.Frame(parent, bootstyle="secondary", padding=1)
        m_frame.pack(side=TOP, fill=BOTH, expand=True, pady=(0, 5))
        m_head = tb.Frame(m_frame, bootstyle="secondary", height=25);
        m_head.pack(fill=X)
        tb.Label(m_head, text=" ● SOURCE MONITOR / 实时预览", foreground="white", font=("Segoe UI", 8, "bold"),
                 bootstyle="inverse-secondary").pack(side=LEFT, padx=10)
        tb.Checkbutton(m_head, text="LIVE", variable=self.monitor_on,
                       bootstyle="primary-round-toggle", command=self.on_monitor_toggle).pack(side=RIGHT, padx=5)
        self.preview_container = tb.Frame(m_frame, bootstyle="light")
        self.preview_container.pack(fill=BOTH, expand=True)
        self.lbl_preview = tb.Label(self.preview_container, text="[ STANDBY ]", anchor="center", font=("Consolas", 14),
                                    foreground="#999")
        self.lbl_preview.pack(fill=BOTH, expand=True)

        c_frame = tb.Frame(parent, bootstyle="secondary", padding=1)
        c_frame.pack(side=BOTTOM, fill=BOTH, expand=True, pady=(5, 0))
        c_head = tb.Frame(c_frame, bootstyle="secondary", height=25);
        c_head.pack(fill=X)
        tb.Label(c_head, text=" ■ LATEST CAPTURE / 最近捕获", foreground="white", font=("Segoe UI", 8, "bold"),
                 bootstyle="inverse-secondary").pack(side=LEFT, padx=10)
        self.capture_container = tb.Frame(c_frame, bootstyle="light")
        self.capture_container.pack(fill=BOTH, expand=True)
        self.lbl_capture = tb.Label(self.capture_container, text="[ READY ]", anchor="center", font=("Consolas", 14),
                                    foreground="#999")
        self.lbl_capture.pack(fill=BOTH, expand=True)

    def on_monitor_toggle(self):
        if not self.monitor_on.get():
            self.lbl_preview.config(image='', text="[ MONITOR DISABLED / 预览已关闭 ]")
        else:
            self.lbl_preview.config(image='', text="[ INITIALIZING STREAM... ]")

    def sanitize_filename(self, name):
        return re.sub(r'[\\/*?:"<>|]', "", name)

    def open_folder(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            self.log(f"Auto-open failed: {e}")

    def set_roi(self):
        video = self.video_path.get()
        if not video or not os.path.exists(video):
            return messagebox.showwarning("Warning", "请先加载有效的视频文件。")
        self.log("Initializing ROI Selector...")
        self.set_status("SETTING ROI")
        messagebox.showinfo("ROI Guide", "操作提示：\n1. 拖动鼠标框选区域\n2. 按 ENTER 键确认\n3. 按 C 键取消")
        try:
            cap = cv2.VideoCapture(video)
            start_txt = self.ent_start.get()
            if start_txt:
                try:
                    h, m, s = map(int, start_txt.split(':'))
                    start_ms = (h * 3600 + m * 60 + s) * 1000
                    cap.set(cv2.CAP_PROP_POS_MSEC, start_ms)
                except ValueError:
                    pass
            ret, frame = cap.read()
            cap.release()
            if not ret: raise ValueError("Cannot read video stream.")

            win_name = "ROI Selector (Enter=Confirm, C=Cancel)"
            cv2.namedWindow(win_name, cv2.WINDOW_NORMAL)
            cv2.resizeWindow(win_name, 1280, 720)
            roi = cv2.selectROI(win_name, frame, showCrosshair=True, fromCenter=False)
            cv2.destroyWindow(win_name)

            x, y, w, h = roi
            if w > 0 and h > 0:
                self.roi_rect = (x, y, w, h)
                self.lbl_roi_status.config(text=f"已锁定: {w}x{h}", foreground="#28a745")
                self.log(f"ROI locked: x={x}, y={y}, w={w}, h={h}")
            else:
                self.roi_rect = None
                self.lbl_roi_status.config(text="全屏扫描", foreground="#999")
                self.log("ROI selection cancelled.")
            self.set_status("READY")
        except Exception as e:
            self.log(f"ROI Error: {e}")
            self.set_status("ERROR", "red")

    def run_logic(self):
        video_path = self.video_path.get()
        base_output = self.output_path.get()
        raw_name = self.project_name.get().strip()
        project_name = self.sanitize_filename(raw_name) or f"Lecture_{int(time.time())}"

        project_dir = os.path.join(base_output, project_name)
        images_dir = os.path.join(project_dir, "Runs")
        pdf_dir = os.path.join(project_dir, "PDFs")

        if not base_output or not os.path.exists(video_path):
            self.log("Error: Invalid paths.")
            self.is_running = False
            self.set_status("CONFIG ERROR", "red")
            self.after(0, lambda: self.btn_run.config(text="INITIALIZE ENGINE / 启动抽取引擎", bootstyle="primary"))
            return

        try:
            os.makedirs(images_dir, exist_ok=True)
            os.makedirs(pdf_dir, exist_ok=True)
        except Exception as e:
            self.log(f"Error creating directories: {e}")
            return

        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            self.log("Error: Cannot open video source.")
            self.is_running = False
            self.set_status("ERROR", "red")
            self.after(0, lambda: self.btn_run.config(text="INITIALIZE ENGINE", bootstyle="primary"))
            return

        start_sec = 0.0
        end_sec = float(cap.get(cv2.CAP_PROP_FRAME_COUNT)) / cap.get(cv2.CAP_PROP_FPS)

        try:
            s_txt = self.ent_start.get()
            e_txt = self.ent_end.get()
            p_start = parse_time(s_txt)
            if p_start >= 0: start_sec = p_start
            p_end = parse_time(e_txt)
            if p_end > 0: end_sec = p_end

            cap.set(cv2.CAP_PROP_POS_MSEC, start_sec * 1000)
            self.log(f"Range set: {format_time(start_sec)} -> {format_time(end_sec)}")
        except Exception as e:
            self.log(f"Time Warning: {e}. Using defaults.")

        fps = cap.get(cv2.CAP_PROP_FPS)
        total_duration = end_sec - start_sec
        if total_duration <= 0: total_duration = 1

        prev_frame_gray = None
        last_captured_hash = None
        stable_counter = 0
        captured_count = 0
        captured_image_paths = []

        self.log(f"Running... Target: {project_name}")
        self.set_status("RUNNING / 运行中", "#00ff00")

        try:
            while self.is_running:
                thresh = self.diff_threshold.get()
                interval = self.check_interval.get()
                stability = self.stability_frames.get()

                frames_to_skip = int(fps * interval)
                if frames_to_skip < 1: frames_to_skip = 1

                for _ in range(frames_to_skip):
                    cap.grab()

                ret, frame = cap.read()
                if not ret: break

                current_pos_sec = cap.get(cv2.CAP_PROP_POS_MSEC) / 1000

                if current_pos_sec > end_sec:
                    self.log(f"Reached end time: {format_time(end_sec)}")
                    self.progress_var.set(100)
                    break

                elapsed = current_pos_sec - start_sec
                percent = int((elapsed / total_duration) * 100)
                percent = max(0, min(100, percent))
                self.progress_var.set(percent)

                # --- [FIX: ROI CROPPING FIRST] ---
                # Apply ROI BEFORE updating UI or Processing
                if self.roi_rect:
                    x, y, w, h = self.roi_rect
                    h_frame, w_frame = frame.shape[:2]
                    x = max(0, x);
                    y = max(0, y)
                    w = min(w, w_frame - x);
                    h = min(h, h_frame - y)
                    # This is the actual image we process and see
                    process_frame = frame[y:y + h, x:x + w]
                else:
                    process_frame = frame

                # Update Monitor with the CROPPED frame
                if self.monitor_on.get():
                    self._update_monitor_ui(process_frame)

                # Algorithm uses CROPPED frame
                gray = cv2.cvtColor(process_frame, cv2.COLOR_BGR2GRAY)
                gray_small = cv2.resize(gray, (64, 64))

                is_static = False
                if prev_frame_gray is not None:
                    err = np.sum((gray_small.astype("float") - prev_frame_gray.astype("float")) ** 2)
                    err /= float(gray_small.shape[0] * gray_small.shape[1])
                    metric = err / 100
                    if metric < thresh:
                        is_static = True
                        stable_counter += 1
                    else:
                        stable_counter = 0
                prev_frame_gray = gray_small

                if is_static and stable_counter == stability:
                    is_unique = True
                    if last_captured_hash is not None:
                        dup_err = np.sum((gray_small.astype("float") - last_captured_hash.astype("float")) ** 2)
                        dup_err /= float(gray_small.shape[0] * gray_small.shape[1])
                        dup_metric = dup_err / 100
                        if dup_metric < (thresh * 1.5):
                            is_unique = False

                    if is_unique:
                        last_captured_hash = gray_small
                        captured_count += 1

                        filename = os.path.join(images_dir, f"slide_{captured_count:04d}.jpg")
                        # Save the CROPPED frame
                        cv2_imwrite_safe(filename, process_frame)
                        captured_image_paths.append(filename)

                        # Update Capture UI with the CROPPED frame
                        self._update_capture_ui(process_frame, captured_count)
                        self.log(f"Saved: slide_{captured_count:04d}.jpg")

                self.var_processed.set(format_time(current_pos_sec))
                time.sleep(0.002)

            if self.make_pdf.get() and captured_image_paths:
                self.log("Generating PDF...")
                self.set_status("GENERATING PDF", "cyan")
                try:
                    pdf_path = os.path.join(pdf_dir, f"{project_name}_Full.pdf")
                    img1 = Image.open(captured_image_paths[0]).convert('RGB')
                    img_list = [Image.open(p).convert('RGB') for p in captured_image_paths[1:]]
                    img1.save(pdf_path, save_all=True, append_images=img_list)
                    self.log("PDF Generated.")
                    messagebox.showinfo("Success", f"Extraction Complete!\nPDF Saved to:\n{pdf_path}")
                except Exception as e:
                    self.log(f"PDF Gen Error: {e}")

            self.open_folder(project_dir)

        except Exception as e:
            self.log(f"Runtime Exception: {e}")
            self.set_status("CRASHED", "red")
        finally:
            cap.release()
            self.is_running = False
            self.set_status("FINISHED / 完成", "cyan")
            self.after(0, lambda: self.btn_run.config(text="INITIALIZE ENGINE / 启动抽取引擎", bootstyle="primary"))
            self.log("Job Done.")

    def _update_monitor_ui(self, frame_cv2):
        color_frame = cv2.cvtColor(frame_cv2, cv2.COLOR_BGR2RGB)
        image = Image.fromarray(color_frame)
        w = self.preview_container.winfo_width() or 400
        h = self.preview_container.winfo_height() or 300
        img_w, img_h = image.size
        ratio = min(w / img_w, h / img_h)
        new_w, new_h = int(img_w * ratio), int(img_h * ratio)
        image = image.resize((new_w, new_h), Image.Resampling.BILINEAR)
        photo = ImageTk.PhotoImage(image)

        def _refresh():
            if self.monitor_on.get():
                self.lbl_preview.config(image=photo, text="")
                self.lbl_preview.image = photo

        self.after(0, _refresh)

    def _update_capture_ui(self, frame_cv2, count):
        color_frame = cv2.cvtColor(frame_cv2, cv2.COLOR_BGR2RGB)
        image = Image.fromarray(color_frame)
        w = self.capture_container.winfo_width() or 400
        h = self.capture_container.winfo_height() or 300
        img_w, img_h = image.size
        ratio = min(w / img_w, h / img_h)
        new_w, new_h = int(img_w * ratio), int(img_h * ratio)
        image = image.resize((new_w, new_h), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(image)

        def _refresh():
            self.lbl_capture.config(image=photo, text="")
            self.lbl_capture.image = photo
            self.var_captured.set(str(count))

        self.after(0, _refresh)

    def select_video(self):
        f = filedialog.askopenfilename(filetypes=[("Video Files", "*.mp4 *.avi *.mkv")])
        if f:
            self.video_path.set(f)
            self.roi_rect = None
            self.lbl_roi_status.config(text="全屏扫描", foreground="#999")
            self.set_status("READY")
            self.log(f"Source loaded: {os.path.basename(f)}")

    def select_output(self):
        d = filedialog.askdirectory()
        if d:
            self.output_path.set(d)
            self.log(f"Output directory set: {d}")

    def open_video_cutter(self):
        video = self.video_path.get()
        if not video or not os.path.exists(video):
            return messagebox.showwarning("Warning", "请先加载有效的视频文件。")

        def sync(s, e):
            self._update_time_ui(s, e)
            self.is_time_locked = True
            self.log(f"Timeline synchronized: {s} - {e}")

        VideoCutterDialog(self, video, parse_time(self.ent_start.get()), parse_time(self.ent_end.get()), sync)

    def _update_time_ui(self, s, e):
        def _exec():
            self.ent_start.config(state="normal")
            self.ent_start.delete(0, tk.END);
            self.ent_start.insert(0, s);
            self.ent_start.config(state="readonly")
            self.ent_end.config(state="normal")
            self.ent_end.delete(0, tk.END);
            self.ent_end.insert(0, e);
            self.ent_end.config(state="readonly")

        self.after(0, _exec)

    def toggle_run(self):
        if self.is_running:
            self.is_running = False
            self.set_status("STOPPING...")
            self.log("Stopping engine...")
            return
        if not self.video_path.get():
            return messagebox.showerror("Error", "请先选择视频文件。")
        self.is_running = True
        self.btn_run.config(text="ABORT ENGINE / 停止运行", bootstyle="danger")
        threading.Thread(target=self.run_logic, daemon=True).start()