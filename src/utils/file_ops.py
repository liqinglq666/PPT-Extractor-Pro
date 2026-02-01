import cv2
import numpy as np
import os


def cv2_imread_safe(file_path, flags=cv2.IMREAD_UNCHANGED):
    """
    Fix: 解决 OpenCV 无法原生读取含中文路径图片的问题
    Args:
        file_path: 图片路径
        flags: OpenCV 读取标志 (如 cv2.IMREAD_COLOR)
    """
    if not os.path.exists(file_path):
        return None

    try:
        # np.fromfile 读取为字节流，再解码
        img_array = np.fromfile(file_path, dtype=np.uint8)
        return cv2.imdecode(img_array, flags)
    except Exception as e:
        # print(f"Read Error: {e}") # Debug only
        return None


def cv2_imwrite_safe(file_path, img, params=None):
    """
    Fix: 支持将图片保存至含中文名称的路径，并自动适配后缀名
    Args:
        file_path: 保存路径 (如 'D:/测试/image.png')
        img: OpenCV 图像矩阵
        params: 编码参数 (如 [cv2.IMWRITE_JPEG_QUALITY, 90])
    """
    if img is None:
        return False

    try:
        # 1. 获取文件后缀名 (如 .jpg, .png)
        ext = os.path.splitext(file_path)[1]
        if not ext:
            ext = ".jpg"  # 默认回退到 jpg

        # 2. 根据后缀名编码
        valid, buf = cv2.imencode(ext, img, params)

        if valid:
            # 3. 写入文件
            buf.tofile(file_path)
            return True
        return False
    except Exception as e:
        # print(f"Write Error: {e}") # Debug only
        return False