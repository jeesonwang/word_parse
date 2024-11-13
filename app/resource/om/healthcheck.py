#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask_restful import Resource

from app.common.response import ResUtil


class HealthCheckResource(Resource, ResUtil):
    def get(self):
        return self.message(message="ok")
