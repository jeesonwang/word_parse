#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json

from flask import Flask
from flask_restful import request
from loguru import logger
from werkzeug.exceptions import HTTPException

from app.config.error_code import MESSAGE, UnknownError, ParamCheckError, DataBaseError, MethodNotAllowed
from app.common.response import ResUtil


class ApiError(Exception):
    def __init__(self, code: int, message: str = None):
        self.code = code
        self.message = message if message is not None else MESSAGE[code]


def record_error_log(func):
    """错误日志记录装饰器"""
    def log_func(err: Exception):
        r = func(err)
        if isinstance(r, tuple):
            # string response
            response_data = r[0]
            status_code = r[1]
        else:
            # json response
            response_data = r.json
            status_code = r.status_code
        err_data = json.dumps({
            "url": request.url,
            "remote_addr": request.remote_addr,
            "request_headers": dict(request.headers),
            "request_data": request.args if request.method == 'GET' else request.get_json(silent=True),
            "response_data": response_data,
            "status_code": status_code
        }, ensure_ascii=False)
        logger.exception(err_data)
        return r

    return log_func


def register_error_handler(app: Flask):
    @app.errorhandler(400)
    @record_error_log
    def catch_400_error(err: Exception):
        """捕捉400错误，常见于flask_restful.reparser参数解析失败"""
        if isinstance(err, HTTPException):
            if getattr(err, "data", None) is None:
                return ResUtil.message(code=ParamCheckError, message="数据解析有误，请检查传入值")
        message = str(err) + ":" + " ".join([str(arg) for arg in err.args])
        return ResUtil.message(code=ParamCheckError, message=message)

    @app.errorhandler(404)
    @record_error_log
    def catch_404_error(err: Exception):
        """捕捉404错误"""
        return "<h2>{code} {name}</h2>\n{desc}".format(code=err.code, name=err.name, desc=err.description), 404

    @app.errorhandler(405)
    @record_error_log
    def catch_405_error(err: Exception):
        """捕捉405错误，常见于url无该方法"""
        return ResUtil.message(code=MethodNotAllowed, message=str(err))

    @app.errorhandler(Exception)
    @record_error_log
    def catch_other_error(err: Exception):
        """捕捉其他类型错误，返回相关错误信息"""
        if getattr(err, "code", None):
            return ResUtil.message(code=getattr(err, "code"), message=getattr(err, "message", "未知错误"))
        return ResUtil.message(code=UnknownError, message=str(err))
