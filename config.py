import os
from dotenv import load_dotenv

class Config:
    """Config class to load environment variables"""

    def __init__(self):
        load_dotenv()
        self.ZabbixAPI = os.getenv("ZABBIX_API")
        self.ZabbixUser = os.getenv("ZABBIX_USER")
        self.ZabbixPass = os.getenv("ZABBIX_PASSWORD")