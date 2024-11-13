#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import numpy as np

DOWNLOAD_PATH = "download"
COLOR_PATH = "download/color"

ELEMENT_COLOR_MAP = {
    'structured_document_tag': [
        [np.array([0, 255, 0]), np.array([80, 255, 0])],
        [np.array([0, 255, 255]), np.array([80, 255, 255])]
    ],
    'paragraph': [  # 段落
        [np.array([172, 0, 100]), np.array([255, 255, 255])],
        [np.array([172, 0, 0]), np.array([255, 255, 100])],
        [np.array([0, 100, 0]), np.array([172, 255, 255])]
    ], 
    'table': [  # 表格
        [np.array([0, 0, 200]), np.array([0, 10, 250])],
        [np.array([255, 192, 203]), np.array([255, 192, 255])]
    ],
    'other': [  # 页眉
        [np.array([0, 25, 0]), np.array([25, 172, 255])],
        [np.array([199, 12, 20]), np.array([199, 150, 255])],
        [np.array([135, 206, 26]), np.array([172, 255, 250])]
    ]
}

