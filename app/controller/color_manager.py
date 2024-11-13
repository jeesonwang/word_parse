import random
from loguru import logger

from app.engine.aspose import Color
from config.const import ELEMENT_COLOR_MAP

class ColorManager:
    def __init__(self, coloring_background: bool = False):
        self.reset()
        self.coloring_background = coloring_background

    def get_color(self, element_type):
        """根据元素类型和当前索引生成随机不重复的颜色。"""
        if element_type not in ELEMENT_COLOR_MAP:
            element_type = "other"

        color_range_list = ELEMENT_COLOR_MAP[element_type]
        color_index = self.element_index % len(color_range_list)
        color_range = color_range_list[color_index]

        def generate_color():
            r_l = min([color_range[0][0], color_range[1][0]])
            r_r = max(color_range[0][0], color_range[1][0])
            r = random.randint(r_l, r_r)

            g_l = min([color_range[0][1], color_range[1][1]])
            g_r = max(color_range[0][1], color_range[1][1])
            g = random.randint(g_l, g_r)

            b_l = min([color_range[0][2], color_range[1][2]])
            b_r = max(color_range[0][2], color_range[1][2])
            b = random.randint(b_l, b_r)
            return (r, g, b)

        color = generate_color()
        n = 0
        while color in self.used_colors:
            color = generate_color()
            if n >= 10:
                color_range = [[1, 1, 1], [255, 255, 255]]
            n += 1
        self.used_colors.add(color)
        self.element_index += 1

        return Color(*color)

    def reset(self):
        self.used_colors = {(0,0,0),(255,255,255)}
        self.element_index = 0
