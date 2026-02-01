import sys
import os
import multiprocessing  # Safety: 引入多进程支持

# Refactor: 运行时路径解析与打包环境适配
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def configure_runtime_path():
    """Note: 确保项目根目录与 src 模块在 sys.path 中"""
    if BASE_DIR not in sys.path:
        sys.path.insert(0, BASE_DIR)

    # 显式添加 src 目录，防止相对导入在某些环境下失效
    src_path = os.path.join(BASE_DIR, 'src')
    if src_path not in sys.path:
        sys.path.append(src_path)


def initialize_high_dpi_awareness():
    """Fix: 解决 Windows 高分屏缩放导致的文字模糊问题"""
    if sys.platform == 'win32':
        try:
            from ctypes import windll
            # 1 = System DPI Aware, 2 = Per Monitor DPI Aware (更优，但需系统支持)
            windll.shcore.SetProcessDpiAwareness(1)
        except (ImportError, AttributeError, OSError):
            pass


def bootstrap():
    """Application Entry Point"""
    # Safety: 防止 Windows 下 PyInstaller 打包后的多进程无限递归炸弹
    multiprocessing.freeze_support()

    configure_runtime_path()
    initialize_high_dpi_awareness()

    try:
        # Note: 延迟导入，确保环境配置完成后再加载 UI 依赖
        from src.ui.main_window import PPTExtractorEngine

        app = PPTExtractorEngine()
        # Refactor: 直接运行，进入主事件循环
        app.mainloop()

    except Exception as e:
        # Safety: 捕获所有启动期异常（不仅是 ImportError），防止程序静默闪退
        err_msg = f"Critical Startup Error:\n{str(e)}"
        sys.stderr.write(err_msg + "\n")

        # 尝试弹出原生对话框提示用户，即使 UI 库加载失败
        try:
            import tkinter.messagebox as mb
            # 创建临时 root 窗口以显示弹窗，随后立即销毁
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            mb.showerror("System Error", f"Application failed to start.\n\nLog:\n{e}")
            root.destroy()
        except:
            # 最后的兜底，如果 tkinter 都坏了，只能写标准错误流
            pass

        sys.exit(1)


if __name__ == "__main__":
    bootstrap()