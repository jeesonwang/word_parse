#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask import Flask, Blueprint
from flask_restful import Api

from .api.generate_word import GenerateWordFile
from .api.paragraph_coloring import ParagraphColoringResource
from .om.healthcheck import HealthCheckResource

api_bp = Blueprint(name="api", import_name="api")
om_bp = Blueprint(name="om", import_name="om")

api = Api(api_bp)
om = Api(om_bp)

# [业务API]
api.add_resource(ParagraphColoringResource, "/paragraph_coloring")

# [运维API]
om.add_resource(HealthCheckResource, "/healthcheck")  # docker健康检查

def register_blueprint(app: Flask):
    """注册蓝图"""
    app.register_blueprint(api_bp)
    app.register_blueprint(om_bp, url_prefix="/om")

