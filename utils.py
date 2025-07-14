import os
import re
from datetime import datetime
import pandas as pd


def check_file_exists(file_path):
    """检查文件是否存在"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")


def check_file_format(file_path, supported_formats):
    """检查文件格式是否支持"""
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in supported_formats:
        raise ValueError(f"不支持的文件格式: {file_ext}，支持的格式: {', '.join(supported_formats)}")


def get_default_title(file_path):
    """获取默认标题"""
    return os.path.basename(file_path)


def determine_data_type(value):
    """确定单元格数据类型"""
    # 处理pandas/numpy的数值类型
    if isinstance(value, (int, float)):
        return 'numeric'

    # 处理日期类型
    if isinstance(value, (datetime, pd.Timestamp)):
        return 'date'

    # 处理布尔类型
    if isinstance(value, bool):
        return 'boolean'

    # 尝试将字符串解析为数值
    try:
        float(value)
        return 'numeric'
    except (ValueError, TypeError):
        pass

    # 尝试将字符串解析为日期
    date_formats = ['%Y-%m-%d', '%m/%d/%Y', '%d-%b-%Y', '%Y-%m-%d %H:%M:%S']
    for fmt in date_formats:
        try:
            datetime.strptime(str(value), fmt)
            return 'date'
        except (ValueError, TypeError):
            pass

    # 尝试将字符串解析为布尔值
    if str(value).lower() in ['true', 'false']:
        return 'boolean'

    # 默认视为字符串
    return 'string'
