#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from app.exception.base import ApiError
from app.config.error_code import ParamCheckError, DataNotFound
from app.config.http_status import HTTP_400_BAD_REQUEST, HTTP_404_NOT_FOUND


class NotFoundError(ApiError):
    default_code = DataNotFound
    default_message = "数据不存在"
    default_http_code = HTTP_404_NOT_FOUND


class ParamsCheckError(ApiError):
    default_code = ParamCheckError
    default_message = "参数校验失败"
    default_http_code = HTTP_400_BAD_REQUEST
