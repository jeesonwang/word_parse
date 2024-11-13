#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import time
import os
import traceback
import uuid

from loguru import logger
from flask_restful import Resource
from app.common.response import ResUtil, fields, args_parser
from werkzeug.datastructures import FileStorage

from app.controller.s3fs_file import s3_controller
from app.controller.word_parse import word
from app.common.log import logger
from config.const import DOWNLOAD_PATH, COLOR_PATH
from engine.parse_results import WordProcessingResults

class ParagraphColoringResource(Resource, ResUtil):
    @args_parser({
        "cloud_path": fields(type=str, required=True),
        "pdf_filepath": fields(type=str, required=True),
        "coloring_pdf_filepath": fields(type=str, required=True)
    })
    def post(self, cloud_path: str, pdf_filepath: str, coloring_pdf_filepath: str):
        if cloud_path.split(".")[-1] not in ("doc", "docx"):
            return self.message(code=1102, message=f"只支持doc或docx的文档")
        
        file_name = os.path.basename(cloud_path)
        # download from s3 to folder
        file_path = os.path.join(DOWNLOAD_PATH, file_name)
        s3_controller.download_file(cloud_path=cloud_path, local_path=file_path)

        if not os.path.exists(file_path):
            return self.message(code=-1, message=f"file:{file_path} not exist")

        results: WordProcessingResults
        results = word.set_paragraph_colors(file_path)
        logger.info(f"原pdf文件: {results.origin_pdf_path}, 染色后的pdf文件: {results.colored_pdf_path}")
        # 上传pdf文件和染色后的pdf文件
        s3_controller.upload(results.origin_pdf_path, pdf_filepath)
        s3_controller.upload(results.colored_pdf_path, coloring_pdf_filepath)
        
        return self.message(data=results.data, message="success")

