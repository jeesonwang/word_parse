#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from datetime import date, time

from flask.json import JSONEncoder


class CustomJSONEncoder(JSONEncoder):
    """修复datetime打印格式"""
    def default(self, obj):
        try:
            if isinstance(obj, (date, time)):
                return str(obj)[:19]
            iterable = iter(obj)
        except TypeError:
            pass
        else:
            return list(iterable)
        return JSONEncoder.default(self, obj)
