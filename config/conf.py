#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""参数提取 用于从环境变量/.env等信息中提取可配置参数"""
import os
import sys
from dotenv import load_dotenv

# 将 BASE_DIR 加入python搜索路径
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)
TEMP_PATH = os.path.join(BASE_DIR, "temp")
os.makedirs(TEMP_PATH, exist_ok=True)
# [读取.env文件，转化为环境变量]
load_dotenv()

# [BASE]
ENV = os.getenv("ENV", "dev")
DEBUG = ENV == "dev"

# [PDF_SERVER]
PDF_SERVER = os.getenv("PDF_SERVER", "")

# [SERVER]
WORD_PARSE = os.getenv("WORD_PARSE", "")

FS_TYPE = os.getenv("FS_TYPE")
S3_KEY_ID = os.getenv("S3_KEY_ID")
S3_KEY_SECRET = os.getenv("S3_KEY_SECRET")
S3_ENDPOINT = os.getenv("S3_ENDPOINT")
