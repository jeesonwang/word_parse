#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import datetime
import json
import re
from typing import Any, Type
from functools import partial


def strip(text: str):
    return text.strip()


def not_in(params: Any, words: (tuple, list)):
    for w in words:
        if w in params:
            raise ValueError("'{}' is not allowed".format(w))
    return params


def not_empty(params: Any):
    if not params:
        raise ValueError("not allowed to be empty")
    return params


def length_check(params: Any, operation: str, length: int):
    if operation == "gt":
        check = bool(len(params) > length)
    elif operation == "gte":
        check = bool(len(params) >= length)
    elif operation == "lt":
        check = bool(len(params) < length)
    elif operation == "lte":
        check = bool(len(params) <= length)
    elif operation == "equal":
        check = bool(len(params) == length)
    else:
        raise ValueError("Not support [operation] '{}'".format(operation))

    if check is False:
        raise ValueError("accept length {} {}".format(operation, length))

    return params


def wrapper_list(params: Any):
    if not isinstance(params, list):
        return [params]
    return params


def as_type(params: Any, type: Type):
    return type(params)


def id_check(params: (int, list)):
    if isinstance(params, int):
        if params < -1:
            raise ValueError("not allowed to lt than -1")
    else:
        for p in params:
            if p < -1:
                raise ValueError("not allowed to lt than -1")
    return params


def id_list(params: str, sep: str = ","):
    return [int(_id) for _id in params.split(sep) if _id]


def positive_int(params: (int, list)):
    if isinstance(params, int):
        if params < 1:
            raise ValueError("not allowed to lt than 1")
    else:
        for p in params:
            if p < 1:
                raise ValueError("not allowed to lt than 1")
    return params


def check_datetime(date_str: str):
    if len(date_str) != 19:
        raise ValueError("only supports '%Y-%m-%d %H:%M:%S'")
    return date_str


def check_duplicate(lst: list):
    if len(lst) != len(set(lst)):
        raise ValueError("have duplicate value")
    return lst


def search_start_datetime(start: (datetime.datetime, str)):
    if isinstance(start, str) and len(start) == 19:
        return start
    if isinstance(start, datetime.datetime):
        start = start.date()
    else:
        start = datetime.datetime.strptime(start, "%Y-%m-%d")
    start = start.strftime("%Y-%m-%d %H:%M:%S")
    return start


def search_end_datetime(end: (datetime.datetime, str)):
    if isinstance(end, str) and len(end) == 19:
        return end

    if isinstance(end, datetime.datetime):
        end = end + datetime.timedelta(days=1)
        end = end.date()
    else:
        end = datetime.datetime.strptime(end, "%Y-%m-%d") + datetime.timedelta(days=1)

    end = end.strftime("%Y-%m-%d %H:%M:%S")
    return end


def validate(params: str, fmt: str = "%Y-%m-%d %H:%M:%S"):
    if not params:
        return None
    try:
        datetime.datetime.strptime(params, fmt)
    except ValueError:
        raise ValueError("Incorrect data format, should be YYYY-MM-DD hh:mm:ss")
    return params


def verify_json(params: dict):
    try:
        json.dumps(params, ensure_ascii=False)
    except json.decoder.JSONDecodeError:
        raise ValueError("wrong json format, please check again")
    return params


email_pattern = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b")


def verify_email(params: (str, list)):
    if isinstance(params, str):
        params = [params]
    for p in params:
        if not email_pattern.fullmatch(p):
            raise ValueError(f"wrong email format: {p}")
    return params


lte_255 = partial(length_check, operation="lte", length=255)
lte_100 = partial(length_check, operation="lte", length=100)
lte_50 = partial(length_check, operation="lte", length=50)
lte_15 = partial(length_check, operation="lte", length=15)
lte_10 = partial(length_check, operation="lte", length=10)
ban = partial(not_in, words=(",", "|"))
as_int = partial(as_type, type=int)
as_float = partial(as_type, type=float)
