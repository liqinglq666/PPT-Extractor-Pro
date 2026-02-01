import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox
from PIL import Image, ImageTk
import cv2
import math

# Note: 确保 time_ops.py 路径正确
from src.utils.time_ops import format_time


class VideoCutterDialog(tb.Toplevel):
    """
    Timeline Slicing Interface.
    Fix: Locked preview frame size to prevent window expansion covering controls.
    """

    def __init__(self, parent, video_path, initial_start_sec, initial_end_sec, callback):
        super().__init__(parent)

        self.title("Timeline Slicer - Video Processing Unit")
        self.geometry("1200x850")
        self.minsize(1000, 750)

        self.callback = callback
        self.video_path = video_path

        # --- 1. Robust Initialization ---
        self.cap = cv2.VideoCapture(video_path)
        if not self.cap.isOpened():
            messagebox.showerror("Error", "无法打开视频文件，请检查路径或文件完整性。")
            self.destroy()
            return

        self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
        self.fps = self.cap.get(cv2.CAP_PROP_FPS)
        if self.fps <= 0: self.fps = 25.0

        # Boundary logic
        duration = self.total_frames / self.fps
        initial_start_sec = max(0, initial_start_sec)
        if initial_end_sec <= 0 or initial_end_sec > duration:
            initial_end_sec = duration

        self.start_frame = int(initial_start_sec * self.fps)
        self.end_frame = int(initial_end_sec * self.fps)
        self.current_frame_idx = self.start_frame

        self.create_ui()

        # Initial state sync
        self.seek_to(self.start_frame)
        self.update_boundary_labels()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def create_ui(self):
        """Layout: 模块化 UI 布局构建"""
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 1. Preview Area [FIX: FIXED SIZE CONTAINER]
        # pack_propagate(False) prevents the frame from shrinking/growing to fit the image
        self.preview_frame = tb.Frame(self, bootstyle="dark", width=1150, height=550)
        self.preview_frame.pack_propagate(False)
        self.preview_frame.pack(fill=BOTH, expand=True, padx=10, pady=(10, 5))

        self.lbl_image = tb.Label(self.preview_frame, text="SIGNAL STANDBY", anchor="center")
        self.lbl_image.pack(fill=BOTH, expand=True)

        # 2. Control Area
        control_container = tb.Frame(self)
        control_container.pack(fill=X, side=BOTTOM, padx=10, pady=10)

        # --- Slider ---
        slider_frame = tb.Frame(control_container, padding=(5, 0, 5, 5))
        slider_frame.pack(fill=X)

        self.slider_var = tb.DoubleVar()
        self.slider = tb.Scale(slider_frame, from_=0, to=self.total_frames,
                               variable=self.slider_var, command=self.on_slide, bootstyle="info")
        self.slider.pack(fill=X)

        # --- Fine Tuning Controls ---
        fine_tune_frame = tb.Frame(control_container, padding=5)
        fine_tune_frame.pack(fill=X, pady=5)

        btn_box = tb.Frame(fine_tune_frame)
        btn_box.pack(anchor=CENTER)

        tb.Button(btn_box, text="<< -1s", bootstyle="outline-secondary",
                  command=lambda: self.step_frame(-int(self.fps))).pack(side=LEFT, padx=2)
        tb.Button(btn_box, text="< Frame", bootstyle="secondary", command=lambda: self.step_frame(-1)).pack(side=LEFT,
                                                                                                            padx=2)
        tb.Button(btn_box, text="Frame >", bootstyle="secondary", command=lambda: self.step_frame(1)).pack(side=LEFT,
                                                                                                           padx=2)
        tb.Button(btn_box, text="+1s >>", bootstyle="outline-secondary",
                  command=lambda: self.step_frame(int(self.fps))).pack(side=LEFT, padx=2)

        # --- Dashboard ---
        dashboard = tb.Frame(control_container, padding=5)
        dashboard.pack(fill=X)

        f1 = tb.Labelframe(dashboard, text=" CURSOR / 指针 ", bootstyle="info", padding=5)
        f1.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))
        self.lbl_current_disp = tb.Label(f1, text="00:00:00", font=("Impact", 18), bootstyle="info")
        self.lbl_current_disp.pack(anchor=CENTER)

        f2 = tb.Labelframe(dashboard, text=" IN / 起点 ", bootstyle="warning", padding=5)
        f2.pack(side=LEFT, fill=BOTH, expand=True, padx=5)
        self.lbl_start_disp = tb.Label(f2, text="00:00:00", font=("Impact", 18), bootstyle="warning")
        self.lbl_start_disp.pack(anchor=CENTER)
        tb.Button(f2, text="SET IN [ 当前帧 ]", bootstyle="warning-outline", command=self.set_start, width=15).pack()

        f3 = tb.Labelframe(dashboard, text=" OUT / 终点 ", bootstyle="danger", padding=5)
        f3.pack(side=LEFT, fill=BOTH, expand=True, padx=(5, 0))
        self.lbl_end_disp = tb.Label(f3, text="00:00:00", font=("Impact", 18), bootstyle="danger")
        self.lbl_end_disp.pack(anchor=CENTER)
        tb.Button(f3, text="SET OUT [ 当前帧 ]", bootstyle="danger-outline", command=self.set_end, width=15).pack()

        btn_confirm = tb.Button(control_container, text="APPLY AND SYNC / 同步配置", bootstyle="success",
                                command=self.confirm, width=30)
        btn_confirm.pack(fill=X, pady=(15, 0))

    # ================= Logic Methods =================

    def seek_to(self, frame_idx):
        frame_idx = max(0, min(frame_idx, self.total_frames - 1))
        self.current_frame_idx = frame_idx
        self.slider.set(frame_idx)
        self.update_preview(frame_idx)
        self.lbl_current_disp.config(text=format_time(frame_idx / self.fps))

    def on_slide(self, val):
        frame_idx = int(float(val))
        if abs(frame_idx - self.current_frame_idx) > 0:
            self.current_frame_idx = frame_idx
            self.update_preview(frame_idx)
            self.lbl_current_disp.config(text=format_time(frame_idx / self.fps))

    def step_frame(self, delta):
        new_frame = self.current_frame_idx + delta
        self.seek_to(new_frame)

    def update_preview(self, frame_idx):
        if not self.cap.isOpened(): return

        self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
        ret, frame = self.cap.read()

        if ret:
            h, w = frame.shape[:2]

            # [FIX: Calculate resize ratio based on fixed container size]
            # Use slightly less than 1150/550 to ensure margins
            max_w, max_h = 1100, 530

            ratio = min(max_w / w, max_h / h)
            new_w, new_h = int(w * ratio), int(h * ratio)

            frame = cv2.resize(frame, (new_w, new_h), interpolation=cv2.INTER_LINEAR)
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb)
            tk_img = ImageTk.PhotoImage(img)

            self.lbl_image.config(image=tk_img, text="")
            self.lbl_image.image = tk_img
        else:
            self.lbl_image.config(image="", text="[ End of Stream ]")

    def set_start(self):
        self.start_frame = self.current_frame_idx
        self.update_boundary_labels()

    def set_end(self):
        self.end_frame = self.current_frame_idx
        self.update_boundary_labels()

    def update_boundary_labels(self):
        self.lbl_start_disp.config(text=format_time(self.start_frame / self.fps))
        self.lbl_end_disp.config(text=format_time(self.end_frame / self.fps))

    def _on_close(self):
        if self.cap.isOpened():
            self.cap.release()
        self.destroy()

    def confirm(self):
        if self.start_frame >= self.end_frame:
            messagebox.showerror("Validation Error", "In-point (起点) must be before Out-point (终点).")
            return

        s_time = format_time(self.start_frame / self.fps)
        e_time = format_time(self.end_frame / self.fps)

        self.callback(s_time, e_time)
        self._on_close()