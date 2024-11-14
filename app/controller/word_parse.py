#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import requests
import tempfile

from loguru import logger
from bs4 import BeautifulSoup

from app.engine.aspose import *
from config.const import COLOR_PATH, DOWNLOAD_PATH
from config.conf import TEMP_PATH
from app.controller.color_manager import ColorManager
from app.engine.parse_results import WordProcessingResults

class WordController(object):
    @staticmethod
    def set_paragraph_colors(file_path, coloring_background: bool = False):
        file_name = os.path.basename(file_path)
        color_manager = ColorManager(coloring_background)
        logger.info(f"============【段落染色开始】 {file_path} =====================")
        doc = get_doc_obj(file_path)
        origin_pdf_path = word.word_convert_pdf(doc_obj=doc, pdf_name=file_name)

        results_ = dict(title=None, paragraph=[])
        for section in doc.getSections():
            ft_lst = word.parse_header_footer(section, color_manager)
            results_["paragraph"].extend(ft_lst)
            body = section.getBody()
            heading_map = {}
            for element in body.getChildNodes():
                word.parse_element(element, heading_map, color_manager, results_)

        colored_file_path = os.path.join(COLOR_PATH, file_name)
        doc.save(colored_file_path)
        colored_pdf_path = word.word_convert_pdf(file_path=colored_file_path)
        logger.info(f"============【段落染色结束】 {file_path}=====================")

        return WordProcessingResults(colored_pdf_path=colored_pdf_path, origin_pdf_path=origin_pdf_path, data=results_)
    
    @staticmethod
    def parse_element(element, heading_map: dict, color_manager: ColorManager, results: dict):

        if element.getNodeType() == NodeType.RUN:
            return
        
        node_type_name = NodeType.getName(element.getNodeType())
        content = str(element.getText()).replace("\r", "").replace("\x07", " ")
        print(f'node type name: {node_type_name}, text: {content}')
        node_type_name = str(node_type_name).lower()
        assist_content = ""
        table_detail = None
        titles = list(heading_map.values())[:3]

        if element.getNodeType() == NodeType.PARAGRAPH:
            para_format = element.getParagraphFormat()
            style_name = para_format.getStyleName()
            color = word.coloring_all_elements(element, color_manager)
            fn_results, fn_content_lst = word.parse_footnote(element, color_manager)
            content = word.filter_content(content, fn_content_lst)
            markdown_content = content

            if style_name[:-2] == "Heading":
                if style_name[-1] != "1":
                    heading_map[style_name] = content
                else:
                    results["title"] = content
                    heading_map = {}
            # 至多3个元素,对应此段落所在的最大的3个标题层级内容
            titles = list(heading_map.values())[:3]
            
            results["paragraph"].extend([{
            "type_": str(style_name).lower(),
            "title": titles if titles else None,
            "content": content,
            "assist": assist_content,
            "markdown_content": markdown_content,
            "color": str(string_color(color)),
            "table_detail": table_detail
            }] + fn_results)
        elif element.getNodeType() == NodeType.STRUCTURED_DOCUMENT_TAG:
            style_name = node_type_name
            color = word.coloring_all_elements(element, color_manager)
            # 过滤格式字符串
            cleaned_text = re.sub(r'\\[A-Za-z]+.*?(\s|$)', '', content)
            cleaned_text = re.sub(r'PAGEREF\s+[A-Za-z0-9_]+', '', cleaned_text)
            cleaned_text = re.sub(r'HYPERLINK _Toc[0-9]+', '', cleaned_text)
            content = re.sub(r'[\u0000-\u001F\u007F]+', '', cleaned_text)
            fn_results, fn_content_lst = word.parse_footnote(element, color_manager)
            content = word.filter_content(content, fn_content_lst)
            markdown_content = content
            results["paragraph"].extend([{
            "type_": str(style_name).lower(),
            "title": titles if titles else None,
            "content": content,
            "assist": assist_content,
            "markdown_content": markdown_content,
            "color": str(string_color(color)),
            "table_detail": table_detail
            }] + fn_results)
            return
        elif element.getNodeType() == NodeType.TABLE:
            word.parse_table(element, heading_map, color_manager, results)
            return
        try:
            child_nodes = element.getChildNodes()
        except AttributeError as e:
            logger.info(f"parse element [{node_type_name}] failed, {e}")
            return
        for child in child_nodes:
            if child.getNodeType() in (NodeType.RUN, NodeType.FOOTNOTE, NodeType.BOOKMARK_START, NodeType.BOOKMARK_END):
                continue
            elif child.getNodeType() == NodeType.SHAPE:
                if child.getShapeType() in [ShapeType.OLE_OBJECT, ShapeType.IMAGE]:
                    image_bytes = child.getImageData().getImageBytes()
                    imgae_base64_str = str(Base64.getMimeEncoder().encodeToString(image_bytes))

                    results["paragraph"].extend([{
                        "type_": "image",
                        "title": titles if titles else None,
                        "content": imgae_base64_str,
                        "assist": assist_content,
                        "markdown_content": "",
                        "color": None,
                        "table_detail": table_detail
                        }])
            elif child.getNodeType() in (NodeType.PARAGRAPH, NodeType.STRUCTURED_DOCUMENT_TAG, NodeType.TABLE):
                word.parse_element(child, heading_map, color_manager, results)

    @staticmethod
    def to_markdown_table(table_element) -> str:
        table_element.convertToHorizontallyMergedCells()
        row_count = table_element.getRows().getCount()
        column_count = table_element.getRows().get(0).getCells().getCount()
        _table = []
        for row in table_element.getRows():
            row_element = []
            for cell in row.getCells():
                row_element.append(str(cell.getText()).replace("\x07", ""))
                if cell.isLastCell():
                    continue
            _table.append(row_element)
        assert len(_table) == row_count
        markdown_table = []
        # 第一行作为表头
        header = _table[0]
        if len(header) < column_count:
            header[1:1] = [''] * (column_count - len(header))
        markdown_table.append("| " + " | ".join(header) + " |")
        markdown_table.append("| " + " | ".join(["---"] * len(header)) + " |")
        for row in _table[1:]:
            if len(row) < column_count:
                row[1:1] = [''] * (column_count - len(row))
            markdown_table.append("| " + " | ".join(row) + " |")
        return "\n".join(markdown_table)

    @staticmethod
    def to_html_table(table_element) -> str:
        new_doc = Document()
        cloned_table = table_element.deepClone(True)
        imported_table = new_doc.importNode(cloned_table, True)
        new_doc.getFirstSection().getBody().appendChild(imported_table)
        temp_file_name = os.path.basename(tempfile.mktemp(suffix='.html', prefix='temp_'))
        file_path = os.path.join(TEMP_PATH, temp_file_name)
        new_doc.save(file_path, SaveFormat.HTML)
        with open(file_path, "r") as f:
            table_html = f.read()
        soup = BeautifulSoup(table_html, 'html.parser')
        for tag in soup.find_all(['span', 'p']):
            tag.unwrap()
        for tag in soup.find_all(True):
            del tag['style']
        tables = soup.find_all('table')
        for table in tables:
            for child_table in table.find_all('table'):
                child_table.decompose()
        table_text = str(tables[0])
        return table_text
    
    @staticmethod
    def coloring_all_elements(node, color_manager: ColorManager):
        """
        对所有元素上色
        """
        def coloring_paragraph(paragraph, p_color):
            if color_manager.coloring_background:
                para_format = paragraph.getParagraphFormat()
                para_format.getShading().setBackgroundPatternColor(p_color)
            runs = paragraph.getRuns()
            for run in runs:
                run.getFont().setColor(p_color)

        if node.getNodeType() == NodeType.PARAGRAPH:
            node_type_name = NodeType.getName(node.getNodeType())
            color = color_manager.get_color(node_type_name)
            coloring_paragraph(node, color)
        elif node.getNodeType() == NodeType.RUN:
            node_type_name = NodeType.getName(node.getNodeType())
            color = color_manager.get_color(node_type_name)
            node.getFont().setColor(color)
        elif node.getNodeType() == NodeType.STRUCTURED_DOCUMENT_TAG:
            node_type_name = NodeType.getName(node.getNodeType())
            color = color_manager.get_color(node_type_name)
            paras = node.getChildNodes(NodeType.PARAGRAPH, True)
            for para in paras:
                coloring_paragraph(para, color)
        else:
            node_type_name = NodeType.getName(node.getNodeType())
            try:
                child_nodes = node.getChildNodes(NodeType.PARAGRAPH, True)
            except AttributeError:
                logger.warning(f"{node.__class__} object has no attribute 'getChildNodes'. The current element type: {node_type_name} coloring failed.")
                return None
            color = color_manager.get_color(node_type_name)
            for child_node in child_nodes:
                coloring_paragraph(child_node, color)
        return color
    
    @staticmethod
    def parse_footnote(element, color_manager: ColorManager):
        style_name = "footnote"
        i_footnote_list, content_lst = [], []
        for fn in element.getChildNodes(NodeType.FOOTNOTE, True):
            text = str(fn.getText()).replace("\r", "").replace("\x07", " ")
            print(f'[Footnote]type_name: {style_name}, text: {text}')
            color = word.coloring_all_elements(fn, color_manager)
            content = text.replace("\x02", "").replace("\u0002", "")
            i_footnote_list.append({
                            "type_": style_name,
                            "title": None,
                            "content": content,
                            "assist": "",
                            "markdown_content": content,
                            "color": str(string_color(color))
                            })
            content_lst.append(text)
        return i_footnote_list, content_lst

    @staticmethod
    def parse_header_footer(section, color_manager: ColorManager):
        """解析页眉页脚"""
        i_header_footer_list = []
        for hf in section.getHeadersFooters():
            node_type_name = str(NodeType.getName(hf.getNodeType()))
            content = str(hf.getText()).replace("\r", "").replace("\x07", " ")
            hf_type = str(HeaderFooterType.toString(hf.getHeaderFooterType()))
            print(f'[HeaderFooter]node type name: {node_type_name}, HeadersFooters type: {hf_type}, text: {content}')
            # 解析页眉页脚中的元素内容
            color = word.coloring_all_elements(hf, color_manager)
            i_header_footer = {
                            "type_": hf_type.lower(),
                            "title": None,
                            "content": content,
                            "assist": "",
                            "markdown_content": content,
                            "color": str(string_color(color))
                            }
            i_header_footer_list.append(i_header_footer)
        return i_header_footer_list

    @staticmethod
    def parse_table(table_node, heading_map: dict, color_manager: ColorManager, results: dict):
        """
        Processing table elements is especially for embedded tables.
        https://docs.aspose.com/words/java/working-with-merged-cells/
        """
        if not heading_map:
            titles = None
        else:
            titles = list(heading_map.values())[:3]

        type_name = "table"
        i_footnote_list, fn_content_lst = word.parse_footnote(table_node, color_manager)

        table_content = str(table_node.getText())
        # color = word.coloring_all_elements(table_node, color_manager)
        assist_content = word.filter_content(word.to_html_table(table_node), fn_content_lst)
        markdown_content = word.filter_content(word.to_markdown_table(table_node), fn_content_lst)
        # if there is an ancestor, the type is a table
        ancestor_table = table_node.getAncestor(Table)
        is_sub_table = True if ancestor_table else False

        table_detail = word.parse_table_detail(table_node, heading_map=heading_map, color_manager=color_manager, is_sub_tabel=is_sub_table, results=results)
        results["paragraph"].extend([{
        "type_": type_name,
        "title": titles,
        "content": word.filter_content(table_content, fn_content_lst),
        "assist": assist_content,
        "markdown_content": markdown_content,
        "color": None,
        "table_detail": table_detail
        }] + i_footnote_list)

    @staticmethod
    def parse_table_detail(table_element, heading_map: dict, color_manager: ColorManager, results: dict, is_sub_tabel: bool):

        def get_current_column(cell, current_column: set, current_index, columns):
            next_cell = cell.getNextSibling()
            if not next_cell:
                return current_column
            if next_cell.getCellFormat().getHorizontalMerge() == CellMerge.PREVIOUS:
                current_index += 1
                current_column.add(columns[current_index])
                new_current_column = get_current_column(next_cell, current_column, current_index, columns)
            else:
                return current_column
            return new_current_column
        
        # 将表格中的单元格转换为具有正确水平合并标志的格式
        table_element.convertToHorizontallyMergedCells()
        row_count = table_element.getRows().getCount()
        # column_count = table_element.getRows().get(0).getCells().getCount()

        cells = []
        columns = []
        table_detail = {
            "is_sub_tabel": is_sub_tabel,
            "row_num": row_count,
            "columns": columns,
            "cells": cells
            }
        for i, row in enumerate(table_element.getRows()):
            previous_content = None
            for j, cell in enumerate(row.getCells()):
                cell_format = cell.getCellFormat()
                horizontal_merge = cell_format.getHorizontalMerge()
                vertical_merge = cell_format.getVerticalMerge()
                column = set()
                cell_paragraph = dict(title=None, paragraph=[])
                word.parse_element(cell, heading_map, color_manager=color_manager, results=cell_paragraph)
                if i == 0:
                    if horizontal_merge == CellMerge.NONE:
                        columns.append(word.filter_content(cell.getText()))
                    elif horizontal_merge == CellMerge.FIRST:
                        previous_content = word.filter_content(cell.getText())
                        columns.append(previous_content)
                    elif horizontal_merge == CellMerge.PREVIOUS:
                        columns.append(previous_content)
                    
                    cells.append({
                    "column": [columns[j]],
                    "HorizontalMerge": horizontal_merge,
                    "VerticalMerge": vertical_merge,
                    "Position": [i, j],
                    "paragraph": cell_paragraph["paragraph"]
                    })
                    table_detail["columns"] = columns
                else:
                    if horizontal_merge == CellMerge.FIRST:
                        column.add(columns[j])
                        column = get_current_column(cell, current_column=column, current_index=j, columns=columns)
                        previous_content = column
                    elif horizontal_merge == CellMerge.PREVIOUS:
                        column = previous_content
                    elif horizontal_merge == CellMerge.NONE:
                        column = [columns[j]]

                    cells.append({
                    "column": list(column),
                    "HorizontalMerge": horizontal_merge,
                    "VerticalMerge": vertical_merge,
                    "Position": [i, j],
                    "paragraph": cell_paragraph["paragraph"]
                    })
        table_detail["cells"] = cells
        return table_detail

    @staticmethod
    def filter_content(content, exclude_contents: list = None) -> str:
        if not isinstance(content, str):
            content = str(content)
        if exclude_contents:
            for exclude_content in exclude_contents:
                content = content.replace(exclude_content, '')
        return content.replace("\r", "").replace("\x07", " ")

    @staticmethod
    def word_convert_pdf(doc_obj: str = None, file_path: str = None, pdf_name: str = None, version = 2016):
        '''
        args:
            doc_obj: aspose document object
            file_path: word file path
            pdf_name: pdf file name
            version: word version
        return: saved pdf file path
        '''
        if not (doc_obj or file_path):
            raise ValueError("Missing input parameters doc_obj or file_path.")
        parent_dir = os.path.dirname(os.path.abspath(__file__))
        grandparent_dir = os.path.dirname(parent_dir)
        grandparent2_dir = os.path.dirname(grandparent_dir)
        font_folder_path = JavaStr(grandparent2_dir + '/fonts')
        font_settins = FontSettings()
        font_settins.setFontsFolder(font_folder_path, True)
        if file_path:
            doc = get_doc_obj(file_path, version=version)
        if doc_obj is not None:
            if not pdf_name:
                raise ValueError("Input parameter doc_obj must input parameter pdf_name.")
            doc = doc_obj
                
        if not pdf_name:
            pdf_name = os.path.basename(file_path)
        pdf_name = pdf_name.split(".")[0] + ".pdf"

        pdf_dir = os.path.join(DOWNLOAD_PATH, 'pdf')
        os.makedirs(pdf_dir, exist_ok=True)
        pdf_path = os.path.join(pdf_dir, pdf_name)
        doc.setFontSettings(font_settins)
        doc.save(pdf_path)
        logger.info(f"pdf file saved in {pdf_path}.")
        return pdf_path


word = WordController()
