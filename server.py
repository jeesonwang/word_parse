#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask import Flask

from config.conf import ENV
from app.config.flask_config import Configs
from app.resource import register_blueprint
from app.common.errors import register_error_handler
from app.common.log import register_logger
from app.common.patch.json_decoder import CustomJSONEncoder


def create_app():
    app = Flask(__name__)
    app.config.from_object(Configs[ENV])
    app.json_encoder = CustomJSONEncoder

    # [保留flask本身错误处理]
    handle_exceptions = app.handle_exception
    handle_user_exception = app.handle_user_exception

    register_blueprint(app)
    register_logger()
    register_error_handler(app)

    # [覆盖flask-restful的error handler为flask原生error handler]
    app.handle_exception = handle_exceptions
    app.handle_user_exception = handle_user_exception

    return app


if __name__ == "__main__":
    create_app().run(port=5001)
