#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask import send_file
from flask_restful import Resource
from app.common.response import ResUtil, fields, args_parser

from app.controller.generate_word import generate_word


class GenerateWordFile(Resource, ResUtil):
    @args_parser({
        "file_name": fields(type=str, required=True),
        "page_style": fields(type=dict, required=True, children={
            "page_margin": fields(type=list, required=True),
            "page_size": fields(type=list, default=[595.3, 841.9]),
        }),
        "content": fields(type=list, required=True),
        "return_file_type": fields(type=str)
    })
    def post(self, file_name: str, page_style: dict, content: list, return_file_type: str):
        word_path = generate_word.generate_word_file(file_name=file_name, page_style=page_style, content=content)
        if return_file_type == "file_stream":
            return send_file(
                word_path,
                download_name=file_name,
                as_attachment=True)
        return self.message({'word_path': word_path}, message="success")
