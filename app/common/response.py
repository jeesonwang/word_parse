#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import types
import functools
from typing import Any, Type, Tuple, List, Callable, Union
from copy import deepcopy

from flask_restful import request
from flask import jsonify, g

from app.config.error_code import MESSAGE


def fields(type: Union[Type, Tuple[Type]], required: Union[bool, Callable] = False, choices: Union[tuple, list] = None,
           dest: str = None, location=None, default: Any = None, children: dict = None,
           process: Union[Tuple[Any], List[Any]] = None):
    if required is True and default:
        raise KeyError("The parameter [default] will not run when 'required = True'")
    if default is not None and not isinstance(default, type):
        raise TypeError("The parameter [{}] does not belong to {}".format(default, type))
    if dest is not None and not dest:
        raise KeyError("The parameter [dest] not allowed to be empty")
    if choices is not None and (not choices or not isinstance(choices, (tuple, list))):
        raise KeyError("The parameter [choices] must have a value when used and belong to the '(list,tuple)'")

    location_enum = ("json", "form", "args", "values", "headers", "cookies", "file", "files")
    if location is not None:
        if location not in location_enum:
            raise KeyError("The parameter [location] just allowed {}".format(location_enum))
        if location not in ("json", "args", "headers", "form", "file", "files"):
            raise KeyError("The parameter [location]-[{}] not supported yet, please add it yourself~".format(location))

    return dict(type=type, required=required, choices=choices, dest=dest, location=location, default=default,
                children=children, process=process)


def parser_input(args_map: dict, *args, **kwargs):
    default_location = "args" if request.method == "GET" else "json"
    input_data = kwargs if kwargs else {}
    for arg_key, arg_value in args_map.items():
        location = arg_value["location"] if arg_value.get("location") is not None else default_location
        if arg_key in input_data:
            input_data[arg_key] = arg_value.get("type")(input_data[arg_key])
        if location == "json":
            v = request.get_json().get(arg_key)
        elif location == "args":
            v = request.args.get(arg_key, type=arg_value.get("type"))
        elif location == "headers":
            v = request.headers.get(arg_key)
        elif location == "form":
            v = request.form.get(arg_key, type=arg_value.get("type"))
        elif location == "file":
            v = request.files.get(arg_key)
        elif location == "files":
            v = request.files.getlist(arg_key)
        else:
            raise KeyError("Unsupported location '{}'".format(location))

        if v is not None:
            input_data[arg_key] = v
    return input_data


def args_process(args_map: dict, input_data: dict):
    for arg_key, arg_value in args_map.items():
        if arg_key not in input_data:
            if arg_value.get("required") is True:
                raise KeyError("Missing required parameters [{}]".format(arg_key))
            if isinstance(arg_value.get("required"), types.FunctionType) and arg_value["required"](input_data):
                raise KeyError("Missing required parameters [{}]".format(arg_key))
            if arg_value["default"] is not None:
                input_data[arg_key] = deepcopy(arg_value["default"])
            else:
                input_data[arg_key] = None

        else:
            if arg_value.get("type"):
                if not isinstance(arg_value["type"], (tuple, list)):
                    arg_value["type"] = (arg_value["type"])
                if not isinstance(input_data[arg_key], arg_value["type"]) and input_data[arg_key] is not None:
                    raise TypeError("The type '{}' of parameter [{}] is not the set type '{}'".format(
                        type(input_data[arg_key]), arg_key, arg_value["type"]))

            if arg_value.get("choices"):
                if input_data[arg_key] not in arg_value["choices"]:
                    raise ValueError("The parameter [{}] value is allowed '{}'".format(arg_key, arg_value["choices"]))

            if arg_value.get("process"):
                if not isinstance(arg_value["process"], (tuple, list)):
                    arg_value["process"] = (arg_value["process"],)
                for f in arg_value["process"]:
                    try:
                        input_data[arg_key] = f(input_data[arg_key])
                    except Exception as e:
                        e.args = ("The parameter [{}] {}".format(arg_key, e.args[0]),)
                        raise e

        if arg_value.get("children"):
            if isinstance(input_data[arg_key], dict):
                args_process(arg_value["children"], input_data[arg_key])
            elif isinstance(input_data[arg_key], list):
                for input_item in input_data[arg_key]:
                    args_process(arg_value["children"], input_item)
            elif input_data[arg_key] is None:
                pass
            else:
                raise TypeError("The elements in 'children' only attention list or dict")

        if arg_value.get("dest") and arg_key in input_data:
            input_data[arg_value["dest"]] = input_data.pop(arg_key)


def args_parser(args_map: dict):
    """
    现有的参数解析框架感觉不都太好用,写了一套,更倾向于flask_restful.reparser
    """

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            input_data = parser_input(args_map, *args, **kwargs)
            args_process(args_map=args_map, input_data=input_data)
            return func(*args, **input_data)

        return wrapper

    return decorator


class ResUtil(object):
    """作为resource中的扩展类使用"""
    @property
    def user_id(self):
        return getattr(g, "user_id", None)

    @property
    def user(self):
        return getattr(g, "user", {})

    @property
    def visible(self):
        return self.user.get("info", {}).get("visible", 0)

    @property
    def headers(self):
        return getattr(g, "headers", None)

    @headers.setter
    def headers(self, value: dict):
        if self.headers is None:
            g.headers = value
        else:
            g.headers.update(value)

    @staticmethod
    def message(data: Any = None, code: int = 0, message: str = None, pager_info: dict = None, origin: Any = None,
                headers: dict = None):
        """
        标准API返回函数

        :param data: api回调数据
        :param code: 见config.error_code.MESSAGE
        :param message: 明文回调信息,会覆写预设信息
        :param pager_info: 分页信息
        :param origin: 无需封装直接返回,会忽略其他参数
        :param headers: 响应请求头信息
        :return: 规范的json数据
        """
        if origin is not None:
            response = origin
            if "code" in origin and code in MESSAGE:
                http_code = MESSAGE[origin["code"]]["http_code"]
            else:
                http_code = 500
        else:
            if code not in MESSAGE:
                code = 1001
            http_code = MESSAGE[code]["http_code"]
            message = MESSAGE[code]["message"] if message is None else message

            response = {
                "code": code,
                "status": http_code,
                "message": message,
                "data": {}
            }

            # 无数据模式
            if data is None:
                response["data"] = {}

            # 对象格式
            elif isinstance(data, dict):
                response["data"].update(data)

            # 列表格式
            elif isinstance(data, list):
                response["data"]["items"] = data
                if pager_info:
                    response["data"].update(pager_info)

            # 其他情况报错
            else:
                raise TypeError("unsupported response data format")

        resp = jsonify(response)
        resp.status_code = http_code
        if headers:
            resp.headers.update(headers)

        return resp
