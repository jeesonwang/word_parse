#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask import Config


class BasicConfig(Config):
    SECRET_KEY = "..."
    JSON_SORT_KEYS = False


class DevelopmentConfig(
    BasicConfig,
):
    ENV = "development"
    DEBUG = True


class TestingConfig(
    BasicConfig
):
    ENV = "testing"
    DEBUG = False


class ProductionConfig(
    BasicConfig
):
    ENV = "production"
    DEBUG = False


Configs = {
    "dev": DevelopmentConfig,
    "test": TestingConfig,
    "prod": ProductionConfig
}
