import re


def format_time(seconds):
    """
    Refactor: 将秒数格式化为 HH:MM:SS。
    Fix: 弃用 timedelta，改用数学取模，解决 >24小时显示为 '1 day...' 导致 UI 错位的问题。
    """
    try:
        seconds = int(float(seconds))  # 兼容 float 输入
        if seconds < 0: seconds = 0

        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)

        # 强制格式化为 00:00:00，即使超过24小时也能正确显示 (e.g. 25:00:00)
        return f"{h:02d}:{m:02d}:{s:02d}"
    except (ValueError, TypeError):
        return "00:00:00"


def parse_time(time_str):
    """
    Fix: 解析 H:M:S 格式，增强鲁棒性。
    Note: 使用正则提取纯数字部分，解耦具体的 UI 占位符（如 'Waiting', '🚫'）。
    """
    if not time_str or not isinstance(time_str, str):
        return -1

    try:
        # 1. 正则清洗：只保留数字和冒号，过滤掉所有中文、字母和特殊符号
        # 例如: "Waiting..." -> "", "00:12:30 (设定)" -> "00:12:30"
        clean_str = re.sub(r'[^\d:]', '', time_str)

        if not clean_str:
            return -1

        # 2. 拆分并计算
        parts = list(map(int, clean_str.split(':')))
        n = len(parts)

        if n == 3:  # HH:MM:SS
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
        elif n == 2:  # MM:SS
            return parts[0] * 60 + parts[1]
        elif n == 1:  # SS
            return parts[0]

        return -1  # 格式怪异 (e.g. "12:30:40:50")

    except Exception:
        return -1