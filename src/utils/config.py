"""
配置管理模块
"""
import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv


class Config:
    """配置管理类"""
    
    def __init__(self, config_path: str = ".env"):
        self.config_path = config_path
        self._load_config()
    
    def _load_config(self):
        """加载配置"""
        # 加载.env文件
        if Path(self.config_path).exists():
            load_dotenv(self.config_path)
        
        # Gemini配置
        self.gemini_api_key = os.getenv("GEMINI_API_KEY")
        
        # 数据库配置
        self.db_driver = os.getenv("QS_DB_DRIVER", "com.jeedsoft.jeedsql.jdbc.Driver")
        self.db_url = os.getenv("QS_DB_URL", "jdbc:jeedsql:jtds:sqlserver://localhost:1433;DatabaseName=QS;")
        self.db_user = os.getenv("QS_DB_USER", "")
        self.db_password = os.getenv("QS_DB_PASSWORD", "")
        
        # BIRT配置
        self.birt_path = os.getenv("BIRT_PATH", "D:/BIRT")
        self.qs_app_name = os.getenv("QS_APP_NAME", "")
        self.qs_extension_path = os.getenv("QS_EXTENSION_PATH", "extension/app/report/birt/working")
        
        # 服务配置
        self.server_host = os.getenv("SERVER_HOST", "0.0.0.0")
        self.server_port = int(os.getenv("SERVER_PORT", "8000"))
        self.log_level = os.getenv("LOG_LEVEL", "INFO")
    
    def get_db_config(self) -> dict:
        """获取数据库配置"""
        return {
            'driver': self.db_driver,
            'url': self.db_url,
            'user': self.db_user,
            'password': self.db_password
        }
    
    def get_gemini_config(self) -> dict:
        """获取Gemini配置"""
        return {
            'api_key': self.gemini_api_key
        }