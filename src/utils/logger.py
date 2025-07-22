"""
日志配置模块
"""
import logging
import sys
from pathlib import Path
from datetime import datetime


def setup_logger(log_level: str = "INFO") -> logging.Logger:
    """
    设置日志记录器
    
    Args:
        log_level: 日志级别
        
    Returns:
        Logger: 配置好的日志记录器
    """
    # 创建logs目录
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # 配置日志格式
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建根日志记录器
    logger = logging.getLogger()
    logger.setLevel(getattr(logging, log_level.upper()))
    
    # 清除现有处理器
    logger.handlers.clear()
    
    # 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 文件处理器
    log_file = log_dir / f"birt_generator_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger