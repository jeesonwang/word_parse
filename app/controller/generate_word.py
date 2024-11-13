#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import time
from app.engine.aspose import *
from loguru import logger


class GenerateWordController(object):
    @staticmethod
    def generate_word_file(file_name: str, page_style: dict, content: list):
        # 创建一个新的Word文档
        document = Document()
        # 设置页面布局
        page_setup(document=document, page_style=page_style)

        # 创建DocumentBuilder对象来添加内容
        builder = DocumentBuilder(document)
        for element in content:
            type = (element.get("type") or element.get("element_type", "")).lower()
            if type == "paragraph":
                generate_word.insert_paragraph(builder=builder, element=element)
            elif type == "table":
                generate_word.insert_table(builder=builder, element=element)
            else:
                logger.warning(f"不支持的元素类型{type}, 跳过")
                continue
            builder.writeln("")

        # 保存
        word_dir = os.path.join(DOWNLOAD_PATH, 'word')
        file_info = os.path.splitext(file_name)
        word_path = os.path.join(word_dir, f"{file_info[0]}_{int(time.time())}{file_info[1]}")
        document.save(word_path)
        return word_path

    @staticmethod
    def insert_paragraph(builder, element):
        alignment_dic = {
            "left": ParagraphAlignment.LEFT,
            "center": ParagraphAlignment.CENTER,
            "right": ParagraphAlignment.RIGHT,
            "justify": ParagraphAlignment.JUSTIFY,  # 两端
            "distributed": ParagraphAlignment.DISTRIBUTED  # 分散对齐
        }
        builder.getParagraphFormat().setAlignment(alignment_dic[element.get("align") or element.get("alignment")])
        builder.getParagraphFormat().setLineSpacing(element.get("line_space", 0))
        spacing_rule_value = LineSpacingRule.fromName(element["spacing_rule"].upper())
        builder.getParagraphFormat().setLineSpacingRule(spacing_rule_value)
        builder.getParagraphFormat().setSpaceBefore(element.get("space_before", 0))
        builder.getParagraphFormat().setSpaceAfter(element.get("space_after", 0))
        builder.getParagraphFormat().setFirstLineIndent(element.get("first_line_indent", 0))
        builder.getParagraphFormat().setLeftIndent(element.get("left_indent", 0))
        builder.getParagraphFormat().setRightIndent(element.get("right_indent", 0))

        for c in element["children"]:
            if "type" in c or "element_type" in c:
                type = (c.get("type") or c.get("element_type")).lower()
                # [https://docs.aspose.com/words/java/use-documentbuilder-to-insert-document-elements/#inserting-a-hyperlink]
                if type == "link":
                    builder.getFont().setColor(Color.BLUE)
                    builder.getFont().setUnderline(Underline.SINGLE)
                    builder.insertHyperlink(c["children"][0]["text"], c["href"], False)
                    builder.getFont().clearFormatting()
                if type == "paragraph":
                    generate_word.insert_paragraph(builder=builder, element=c)
                if type == "shape":
                    generate_word.insert_shape(builder=builder, element=c)
                if type == "shape_and_run":
                    generate_word.insert_shape(builder=builder, element=c)

                    builder.getFont().setBold(c.get("bold", False))
                    builder.getFont().setItalic(c.get("italic", False))
                    builder.getFont().setSize(c.get("font_size"))
                    builder.getFont().setName(c.get("font"))
                    if "font_color" in c:
                        builder.getFont().setColor(hex_color_trans_rgb(c["font_color"]))
                    else:
                        builder.getFont().setColor(hex_color_trans_rgb('#000000'))
                    builder.write(c["text"])
            else:
                if c["text"] == "":
                    builder.write(c["text"])
                else:
                    builder.getFont().setBold(c.get("bold", False))
                    builder.getFont().setItalic(c.get("italic", False))
                    builder.getFont().setSize(c.get("font_size"))
                    builder.getFont().setName(c.get("font"))
                    if "font_color" in c:
                        builder.getFont().setColor(hex_color_trans_rgb(c["font_color"]))
                    else:
                        builder.getFont().setColor(hex_color_trans_rgb('#000000'))

                    type = (element.get("type") or element.get("element_type")).lower()
                    if type != "paragraph":
                        builder.writeln(c["text"])
                    else:
                        builder.write(c["text"])
        return builder

    @staticmethod
    def insert_shape(builder, element):
        if element["shape_type"] == "Image" and element.get("image_path"):
            image_path = element.get("image_path")
            shape = builder.insertImage(image_path, element["width"], element["height"])
            shape.setLeft(element["position"][0])
            shape.setTop(element["position"][1])
            h_alignment = "NONE" if element.get("h_alignment") == "NONE | DEFAULT" else element.get("h_alignment")
            v_alignment = "NONE" if element.get("v_alignment") == "NONE | DEFAULT" else element.get("v_alignment")
            shape.setHorizontalAlignment(HorizontalAlignment.fromName(h_alignment))
            shape.setVerticalAlignment(VerticalAlignment.fromName(v_alignment))
            shape.setDistanceTop(element.get("margin_top", 0))
            shape.setDistanceBottom(element.get("margin_bottom", 0))
            shape.setDistanceLeft(element.get("margin_left", 0))
            shape.setDistanceRight(element.get("margin_right", 0))
            shape.setFillColor(hex_color_trans_rgb(element.get("fill_color", "#FFFFFF")))
            shape.setStrokeColor(hex_color_trans_rgb(element.get("stroke_color", "#FFFFFF")))
            shape.setStrokeWeight(element.get("stroke_weight", 0))
        return builder

    @staticmethod
    def insert_table(builder, element):
        max_cell = max([len(t["children"]) for t in element["children"]])

        for r, row in enumerate(element["children"]):
            for c, cell in enumerate(row["children"]):
                content = cell['children'][0]['children'][0].get('text', '') if cell['children'][0]['children'] else ''
                if not cell['children'][0]['children']:
                    # 首行是边界，上一行此列已经表示要被合并，也不修改
                    if r - 1 >= 0:
                        if element["children"][r - 1]["children"][c].get("rowmerge_flag") != "previous":
                            element["children"][r - 1]["children"][c]["rowmerge_flag"] = "first"
                        # 若是上一行下一列不是最后一列，标记列合并结束
                        if c + 1 < max_cell:
                            element["children"][r - 1]["children"][c + 1]["rowmerge_flag"] = "none"

                    cell["rowmerge_flag"] = "previous"
                else:
                    if cell['children'][0]['children'][0].get('text', ''):
                        # 非首列，且前一列是合并列，标记为none
                        if c - 1 >= 0 and element["children"][r]["children"][c - 1].get('rowmerge_flag') == "previous":
                            cell["rowmerge_flag"] = "none"

        cellmerge_none = False  # 水平合并单元格结束的标识
        table = builder.startTable()

        for row in element["children"]:
            cell_num = len(row["children"])
            for i, cell in enumerate(row["children"]):
                builder.insertCell()
                if 'cell_padding' in cell:
                    builder.getCellFormat().setPaddings(cell['cell_padding'][2], cell['cell_padding'][0],
                                                        cell['cell_padding'][3],
                                                        cell['cell_padding'][1])  # left/top/right/bottom
                if cellmerge_none:
                    builder.getCellFormat().setHorizontalMerge(CellMerge.NONE)
                if cell_num < max_cell:
                    # 水平合并单元格开始标识
                    if i == max_cell - cell_num - 1:
                        builder.getCellFormat().setHorizontalMerge(CellMerge.FIRST)

                # 处理垂直合并
                if cell.get("rowmerge_flag"):
                    if cell["rowmerge_flag"] == "first":
                        builder.getCellFormat().setVerticalMerge(CellMerge.FIRST)
                    elif cell["rowmerge_flag"] == "none":
                        builder.getCellFormat().setVerticalMerge(CellMerge.NONE)
                    elif cell["rowmerge_flag"] == "previous":
                        builder.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS)

                for paragraph in cell["children"]:
                    type = (paragraph.get("type") or paragraph.get("element_type")).lower()
                    if type == "paragraph":
                        generate_word.insert_paragraph(builder, paragraph)
                    elif type == "table":
                        generate_word.insert_table(builder, paragraph)
                    else:
                        raise Exception(f"不支持{type}这种元素类型")
            # 若是存在水平合并单元格，将当前元素合并到开始元素中
            if max_cell - cell_num > 0:
                for i in range(max_cell - cell_num):
                    builder.insertCell()
                    builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS)
                    cellmerge_none = True
            else:
                cellmerge_none = False
            builder.endRow()

        table.setLeftIndent(element.get("table_left_indent", 0))
        if element.get("table_text_wrapping") == "AROUND":
            table.setTextWrapping(TextWrapping.AROUND)
            table.setRelativeHorizontalAlignment(
                TableAlignment.fromName(element.get("table_alignment", ["LEFT", "LEFT"])[0]))
            table.setRelativeVerticalAlignment(
                TableAlignment.fromName(element.get("table_alignment", ["LEFT", "LEFT"])[1]))
        else:
            table.setTextWrapping(TextWrapping.NONE)
            table.setAlignment(TableAlignment.fromName(element.get("table_alignment", ["LEFT"])[0]))

        # 先设置单元格之间的空间量,在设置setAllowCellSpacing，否则边框会有重影
        table.setCellSpacing(element.get("table_cell_spacing", 0))
        table.setAllowCellSpacing(element.get("table_allow_cell_spacing", False))
        # 设置边框底纹
        for border in (
                "table_top_border", "table_horizontal_border", "table_bottom_border", "table_left_border",
                "table_vertical_border",
                "table_right_border"):
            boder_type = BorderType.fromName(border.split("_")[1].upper())
            line_style = LineStyle.fromName(element[border]["line_style"])
            color = hex_color_trans_rgb(element[border]["color"])
            table.setBorder(boder_type, line_style, element[border]["width"], color, False)
        table.setAllowAutoFit(element.get("table_allow_auto_fit", True))

        builder.endTable()
        return builder


generate_word = GenerateWordController()
