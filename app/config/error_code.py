#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from app.config.http_status import *

# API内部错误，以1101开头
RequestSuccess = 0
UnknownError = 1001
ParamCheckError = 1101
ParamTypeError = 1102
DataBaseError = 1103
DataExistsError = 1104
AccessDenied = 1105
RequestTimeout = 1106
ExternalServerError = 1107
InsideServerError = 1108
ServiceUnavailable = 1109
MethodNotAllowed = 1110
DataNotFound = 1111
DataUpdateError = 1112
DataChangeError = 1113
DataDelError = 1114


MESSAGE = {
    RequestSuccess: {"message": "请求成功", "http_code": HTTP_200_OK},
    UnknownError: {"message": "未知错误", "http_code": HTTP_500_INTERNAL_SERVER_ERROR},
    # API内部错误，以1101开头
    ParamCheckError: {"message": "参数错误", "http_code": HTTP_400_BAD_REQUEST},
    ParamTypeError: {"message": "数据格式不正确", "http_code": HTTP_400_BAD_REQUEST},
    DataBaseError: {"message": "数据库错误", "http_code": HTTP_400_BAD_REQUEST},
    DataExistsError: {"message": "数据已存在", "http_code": HTTP_400_BAD_REQUEST},
    AccessDenied: {"message": "请求被拒绝", "http_code": HTTP_403_FORBIDDEN},
    RequestTimeout: {"message": "等待超时", "http_code": HTTP_408_REQUEST_TIMEOUT},
    ExternalServerError: {"message": "外部服务异常", "http_code": HTTP_500_INTERNAL_SERVER_ERROR},
    InsideServerError: {"message": "内部服务异常", "http_code": HTTP_500_INTERNAL_SERVER_ERROR},
    ServiceUnavailable: {"message": "接口异常，请稍后再试", "http_code": HTTP_503_SERVICE_UNAVAILABLE},
    MethodNotAllowed: {"message": "方法不允许", "http_code": HTTP_405_METHOD_NOT_ALLOWED},
    DataNotFound: {"message": "资源不存在", "http_code": HTTP_403_FORBIDDEN},
    DataUpdateError: {"message": "资源更新失败", "http_code": HTTP_400_BAD_REQUEST},
    DataChangeError: {"message": "资源修改失败", "http_code": HTTP_400_BAD_REQUEST},
    DataDelError: {"message": "资源删除失败", "http_code": HTTP_400_BAD_REQUEST},
}
