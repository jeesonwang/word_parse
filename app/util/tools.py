#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from loguru import logger
import time


def time_cost(function):
    """
    装饰器函数timer
    :param function:想要计时的函数
    :return:
    """

    def wrapper(*args, **kwargs):
        time_start = time.time()
        res = function(*args, **kwargs)
        cost_time = time.time() - time_start
        logger.info("{} cost time:{}s".format(function.__name__, cost_time))
        return res

    return wrapper