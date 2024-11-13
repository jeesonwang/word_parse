import os
import re
import uuid
import base64
import shutil
import jpype
from jpype import JClass, JString
from loguru import logger

from config.const import DOWNLOAD_PATH

# 免费，水印版
# jpype.startJVM(jvm_path, r"-Djava.class.path=jar\aspose-words-21.1.0-jdk17.jar")
# 破解版
jpype.startJVM(r"-Djava.class.path=jar/aspose-words-20.1-jdk17.jar")

LoadOptions = JClass('com.aspose.words.LoadOptions')
MsWordVersion = JClass('com.aspose.words.MsWordVersion')
Document = JClass('com.aspose.words.Document')
NodeCollection = JClass('com.aspose.words.NodeCollection')
NodeType = JClass('com.aspose.words.NodeType')
HeaderFooterType = JClass('com.aspose.words.HeaderFooterType')
Shape = JClass('com.aspose.words.Shape')
HorizontalAlignment = JClass('com.aspose.words.HorizontalAlignment')
VerticalAlignment = JClass('com.aspose.words.HorizontalAlignment')
ShapeType = JClass('com.aspose.words.ShapeType')
Paragraph = JClass('com.aspose.words.Paragraph')
Alignment = JClass('javax.swing.GroupLayout.Alignment')
NumberStyle = JClass('com.aspose.words.NumberStyle')
LayoutCollector = JClass('com.aspose.words.LayoutCollector')
ConvertUtil = JClass('com.aspose.words.ConvertUtil')
DocumentBuilder = JClass('com.aspose.words.DocumentBuilder')
Comment = JClass('com.aspose.words.Comment')
CommentRangeStart = JClass('com.aspose.words.CommentRangeStart')
CommentRangeEnd = JClass('com.aspose.words.CommentRangeEnd')
Run = JClass('com.aspose.words.Run')
ImageType = JClass('com.aspose.words.ImageType')
FieldType = JClass('com.aspose.words.FieldType')
Field = JClass('com.aspose.words.Field')
FieldHyperlink = JClass('com.aspose.words.FieldHyperlink')
ImageSaveOptions = JClass('com.aspose.words.ImageSaveOptions')
SaveFormat = JClass('com.aspose.words.SaveFormat')
FontSettings = JClass('com.aspose.words.FontSettings')
TextWrapping = JClass('com.aspose.words.TextWrapping')
TableAlignment = JClass('com.aspose.words.TableAlignment')
LineStyle = JClass('com.aspose.words.LineStyle')
BorderType = JClass('com.aspose.words.BorderType')
LineSpacingRule = JClass("com.aspose.words.LineSpacingRule")
ParagraphAlignment = JClass('com.aspose.words.ParagraphAlignment')
Underline = JClass('com.aspose.words.Underline')
SectionLayoutMode = JClass('com.aspose.words.SectionLayoutMode')
Orientation = JClass('com.aspose.words.Orientation')
FootnoteType = JClass('com.aspose.words.FootnoteType')
NumberStyle = JClass('com.aspose.words.NumberStyle')
FootnotePosition = JClass('com.aspose.words.FootnotePosition')
FootnoteNumberingRule = JClass('com.aspose.words.FootnoteNumberingRule')
CellMerge = JClass('com.aspose.words.CellMerge')
HtmlSaveOptions = JClass('com.aspose.words.HtmlSaveOptions')
Table = JClass('com.aspose.words.Table')

File = JClass('java.io.File')
Date = JClass('java.util.Date')
FileInputStream = JClass('java.io.FileInputStream')
Color = JClass('java.awt.Color')
Integer = JClass('java.lang.Integer')
FileOutputStream = JClass('java.io.FileOutputStream')
Base64 = JClass('java.util.Base64')
JavaStr = JClass("java.lang.String")
JBoolean = jpype.JClass('java.lang.Boolean')


def string_color(color):
    """
    java颜色对象转字符串
    :param color: java.awt.Color
    :return:
    """
    if not color:
        return color
    rgb = color.getRGB()
    hex_value = hex(rgb & 0xffffff)[2:]
    hex_value = hex_value.zfill(6)
    hex_color = "#" + hex_value
    return hex_color


def hex_color_trans_rgb(hex_color):
    rgb_color = tuple(int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    return Color(*rgb_color)


def get_hide_texts(doc):
    """
    获取文档中隐藏的code集合，用于过滤解析结果中页面不可见内容
    """
    field_codes = []
    for field in doc.getRange().getFields():
        field_code = field.getFieldCode()
        if field_code not in field_codes:
            field_codes.append(field_code)
    return field_codes


def get_doc_obj(file_path, version=2016):
    load_option = LoadOptions()
    load_option.setMswVersion(MsWordVersion.fromName(f"WORD_{version}"))

    file = File(file_path)
    input_stream = FileInputStream(file)
    doc = Document(input_stream)
    return doc


def parse_section_general_info(section, extend_params=None):
    """
    解析section通用信息
    https://reference.aspose.com/words/zh/java/com.aspose.words/section/
    """
    i_section = {}
    # 获取页面布局对象
    pg_setup = section.getPageSetup()
    # 获取左右上下边距
    margin = [pg_setup.getLeftMargin(), pg_setup.getRightMargin(), pg_setup.getTopMargin(),
              pg_setup.getBottomMargin()]

    i_section['margin'] = [round(i, 2) for i in margin]
    # 获取页宽和页高
    i_section['page_size'] = [round(pg_setup.getPageWidth(), 2), round(pg_setup.getPageHeight(), 2)]
    # 获取内容真实的宽和高
    print_width = pg_setup.getPageWidth() - pg_setup.getLeftMargin() - pg_setup.getRightMargin()
    print_height = pg_setup.getPageHeight() - pg_setup.getTopMargin() - pg_setup.getBottomMargin()
    i_section['print_size'] = [round(print_width, 2), round(print_height, 2)]
    i_section['word_count'] = int(section.getText().length())
    # 获取文档网格中每页的行数
    i_section['row_count'] = int(pg_setup.getLinesPerPage())
    # 获取文档网格中每行的字符数
    i_section['col_count'] = int(pg_setup.getCharactersPerLine())

    # 这个字段判断首页是否有页码，为true，首页无页码
    i_section['firstpage_nofooter'] = pg_setup.getDifferentFirstPageHeaderFooter()
    # 这个字段判断奇偶页是否相同，为true，不同
    i_section['oddandevenpages_diff'] = pg_setup.getOddAndEvenPagesHeaderFooter()

    if extend_params and "page_num_style" in extend_params:
        # page_num_style枚举含义见https://apireference.aspose.com/words/java/com.aspose.words/NumberStyle
        # NUMBER_IN_DASH表示“- 2 -”这样的页码格式
        i_section['page_num_style'] = str(NumberStyle.getName(section.getPageSetup().getPageNumberStyle()))
    return i_section


def parse_paragraph_general_info(paragraph, layout=None):
    """
    解析段落通用信息
    https://reference.aspose.com/words/zh/java/com.aspose.words/paragraph/
    https://reference.aspose.com/words/zh/java/com.aspose.words/paragraphformat/
    """
    # 获取段落的行距、左边缩进、右边缩进、首行缩进
    i_paragraph = {
        'line_space': round(float(paragraph.getParagraphFormat().getLineSpacing()), 2),
        'left_indent': round(float(paragraph.getParagraphFormat().getLeftIndent()), 2),
        'right_indent': round(float(paragraph.getParagraphFormat().getRightIndent()), 2),
        'first_line_indent': round(float(paragraph.getParagraphFormat().getFirstLineIndent()), 2),
    }
    i_paragraph['spacing_rule'] = str(LineSpacingRule.getName(paragraph.getParagraphFormat().getLineSpacingRule())).lower()

    alignment_map = {
        0: "left",
        1: "center",
        2: "right",
        3: "justify",
        4: "distributed",
        5: "ARABIC_MEDIUM_KASHIDA".lower(),
        6: "other",  # aspose中6表示其他未知情况
        7: "ARABIC_HIGH_KASHIDA".lower(),
        8: "ARABIC_LOW_KASHIDA".lower(),
        9: "THAI_DISTRIBUTED".lower()
    }
    # 文本对齐方式
    i_paragraph['alignment'] = alignment_map[paragraph.getParagraphFormat().getAlignment()]
    if layout:
        i_paragraph['page'] = layout.getStartPageIndex(paragraph)
    return i_paragraph


def parse_run_general_info(run, hide_texts=None):
    """
    解析run通用信息
    """
    font = run.getFont()
    text = run.getText()
    if hide_texts and text in hide_texts:
        print(f"过滤不显示内容，text={text}")
        return {}
    return {
        "text": str(text).replace("\r", "\n"),
        "font": str(font.getName()),
        "font_color": str(string_color(font.getColor())),
        "font_size": round(float(font.getSize()), 2),
        "bold": font.getBold()
        # "font_line_spacing": round(float(font.getLineSpacing()), 2),
        # "font_spacing": round(float(font.getSpacing()), 2)
    }


def parse_shape_general_info(shape, extend_params=None):
    """
    解析shape通用信息
    https://reference.aspose.com/words/zh/java/com.aspose.words/shape/
    """
    i_shape = {
        "height": round(float(shape.getHeight()), 2),
        "width": round(float(shape.getWidth()), 2),
        "position": [
            round(float(shape.getLeft()), 2),
            round(float(shape.getTop()), 2),
            round(float(shape.getRight()), 2),
            round(float(shape.getBottom()), 2)
        ],
        "h_alignment": str(HorizontalAlignment.getName(shape.getHorizontalAlignment())),
        "v_alignment": str(VerticalAlignment.getName(shape.getVerticalAlignment())),
        "margin_left": round(float(shape.getDistanceLeft()), 2),
        "margin_right": round(float(shape.getDistanceRight()), 2),
        "margin_top": round(float(shape.getDistanceTop()), 2),
        "margin_bottom": round(float(shape.getDistanceBottom()), 2),
        "fill_color": string_color(shape.getFillColor()),
        "stroke_color": string_color(shape.getStrokeColor()),  # 线条颜色
        "stroke_weight": round(float(shape.getStrokeWeight()), 2)
    }
    if extend_params and "image_path" in extend_params:
        # 获取图片类型 ImageType.toString(shape.getImageData().getImageType())  如Jpeg
        if shape.getShapeType() in [ShapeType.OLE_OBJECT, ShapeType.IMAGE]:
            image_name = str(uuid.uuid1()) + '.' + str(ImageType.toString(shape.getImageData().getImageType()))
            image_dir = os.path.join(DOWNLOAD_PATH, 'image')
            os.makedirs(image_dir, exist_ok=True)
            image_path = os.path.join(image_dir, image_name)

            image_bytes = shape.getImageData().getImageBytes()
            imgae_base64_str = str(Base64.getMimeEncoder().encodeToString(image_bytes))
            imgdata = base64.b64decode(imgae_base64_str)
            with open(image_path, 'wb') as file:
                file.write(imgdata)
            # 下面这些方式保存图片，在flask框架使用时一直卡主
            # image_options = ImageSaveOptions(SaveFormat.JPEG)
            # shape.getShapeRenderer().save(image_path, image_options)
            # shape.getShapeRenderer().save(image_path, None)
            # stream = FileOutputStream(image_path)
            # shape.getShapeRenderer().save(stream, None)
            i_shape["image_path"] = image_path

    return i_shape


def get_page_number_alignment(section):
    """
    获取页码对齐方式
    :param section:
    :return:
    """
    for ht in section.getHeadersFooters():
        if ht.getHeaderFooterType() == HeaderFooterType.FOOTER_PRIMARY:
            for p in ht.getParagraphs():
                for n in p.getChildNodes(NodeType.SHAPE, True).toArray():
                    if n.getShapeType() == ShapeType.TEXT_BOX and n.getText().contains("PAGE"):
                        return str(HorizontalAlignment.getName(n.getHorizontalAlignment()))
    return ""


def parse_cell_style(cell):
    cell_format = cell.getCellFormat()
    cell_style = {
        'cell_padding': [cell_format.getTopPadding(), cell_format.getBottomPadding(), cell_format.getLeftPadding(),
                    cell_format.getRightPadding()]  # 获取要添加到单元格内容上\下\左\右的空间量
    }
    return cell_style


def parse_row_style(row):
    row_style = {
        'row_height': row.getRowFormat().getHeight()  # 单位磅
    }
    return row_style


def parse_table_style(element, doc=None):
    # 表格样式
    text_wrapping = str(TextWrapping.getName(element.getTextWrapping()))  # 环绕方式
    alignment = []  # 对齐方式
    if element.getTextWrapping() == TextWrapping.AROUND:
        alignment = [str(TableAlignment.getName(element.getRelativeHorizontalAlignment())),
                     str(TableAlignment.getName(element.getRelativeVerticalAlignment()))]  # 浮动表格的相对水平、垂直对齐方式
    else:
        alignment = [str(TableAlignment.getName(element.getAlignment()))]  # 对齐方式

    borders = element.getFirstRow().getRowFormat().getBorders()
    # 获取表格所有的边框，若是宽度都为0且类型都是None,采用cell方式获取边框样式
    borders_obj = {
        'top': borders.getTop(),  # 获取上边框
        'horizontal': borders.getHorizontal(),  # 获取在单元格或相应段落之间使用的水平边框
        'bottom': borders.getBottom(),  # 底部边框
        'left': borders.getLeft(),
        'vertical': borders.getVertical(),  # 获取单元格之间使用的垂直边框
        'right': borders.getRight()
    }
    # widths = [v.getLineWidth() for v in borders_obj.values()]
    # linestyles = [v.getLineStyle() for v in borders_obj.values()]
    border_style = {f'table_{bo}_border': {
        "color": "#000000",
        "width": 0.0,
        "line_style": "NONE"
    } for bo in borders_obj}
    # if any(widths) or any(linestyles):
    #     for k, v in borders_obj.items():
    #         border_style[f'table_{k}_border'] = {
    #             'color': str(string_color(v.getColor())),
    #             'width': v.getLineWidth(),
    #             'line_style': str(LineStyle.getName(v.getLineStyle()))
    #         }
    # else:
    row_index, row_count = 1, element.getRows().getCount()
    cell_index, cell_count = 1, element.getRows().get(0).getCells().getCount()
    print(f"====================表格边框列式解析，{row_count}行{cell_count}列")
    for row in element.getRows():
        for cell in row.getCells():
            cell_borders = cell.getCellFormat().getBorders()
            borders_obj = {
                'top': cell_borders.getTop(),  # 获取上边框
                'horizontal': cell_borders.getTop(),  # 获取在单元格或相应段落之间使用的水平边框
                'bottom': cell_borders.getBottom(),  # 底部边框
                'left': cell_borders.getLeft(),
                'vertical': cell_borders.getLeft(),  # 获取单元格之间使用的垂直边框
                'right': cell_borders.getRight()
            }
            k_list = []
            if row_count == 1 and cell_count == 1:  # 1行1列表格
                k_list = ['top', 'bottom', 'left', 'right']
            elif row_count == 1 and cell_count > 1:  # 1行多列表格
                if cell_index == 1:
                    k_list = ['top', 'bottom', 'left']
                elif cell_index == cell_count:
                    k_list = ['vertical', 'right']
            elif row_count > 1 and cell_count == 1:  # 多行1列表格
                if row_index == 1:
                    k_list = ['top', 'left', 'right']
                elif row_index == row_count:
                    k_list = ['horizontal', 'bottom']
            elif row_count > 1 and cell_count > 1:  # 多行多列表格
                if row_index == 1:
                    k_list = ['top']
                elif row_index == row_count:
                    k_list = ['horizontal', 'bottom']

                if cell_index == 1:
                    k_list.append('left')
                elif cell_index == cell_count:
                    k_list.extend(['vertical', 'right'])
            print(f'==============================={row_index}-{cell_index}-----{k_list}')
            for k in k_list:
                v = borders_obj[k]
                border_style[f'table_{k}_border'] = {
                    'color': str(string_color(v.getColor())),
                    'width': v.getLineWidth(),
                    'line_style': str(LineStyle.getName(v.getLineStyle()))
                }
            cell_index += 1
        row_index += 1

    # 表格边框对象
    # border = element.getStyle().getBorders()
    table_style = {
        'table_text_wrapping': text_wrapping,
        'table_alignment': alignment,
        'table_left_indent': element.getLeftIndent(),  # 左缩进
        'table_width': element.getPreferredWidth().getValue(),  # 宽，结果是点（point）单位，而不是厘米（cm）单位，可以将点的值除以28.35，得到厘米的值。
        # 'height': '',  # 高 TODO
        # 'vertical_alignment': '', # 垂直对齐方式 TODO
        'table_padding': [element.getTopPadding(), element.getBottomPadding(), element.getLeftPadding(),
                    element.getRightPadding()],  # 单元格边距 上下左右
        'table_allow_cell_spacing': element.getAllowCellSpacing(),  # 允许单元格间距
        'table_cell_spacing': element.getCellSpacing(),  # 单元格间距距离
        'table_allow_auto_fit': element.getAllowAutoFit(),  # 自动重调尺寸以适应内容
        # 'table_page_border_color': str(string_color(border.getColor())),  # 页面边框颜色,
        # 'table_page_border_style': str(LineStyle.getName(border.getLineStyle())),  # 页面边框样式
        # 'table_page_border_width': border.getLineWidth(),  # 页面边框宽度
    }
    table_style.update(border_style)
    return table_style


def page_setup(document, page_style):
    # 设置页面布局
    page_setup = document.getSections().get(0).getPageSetup()
    page_margin = page_style["page_margin"]
    # 设置左右上下页边距
    page_setup.setLeftMargin(page_margin[0])
    page_setup.setRightMargin(page_margin[1])
    page_setup.setTopMargin(page_margin[2])
    page_setup.setBottomMargin(page_margin[3])
    # 设置纸张大小
    page_size = page_style.get("page_size", [595.3, 841.9])  # 默认A4
    page_setup.setPageWidth(page_size[0])
    page_setup.setPageHeight(page_size[1])
    # 设置方向
    page_setup.setOrientation(Orientation.fromName(page_style.get("orientation", "PORTRAIT")))
    # 设置页眉页脚边距
    if "header_istance" in page_style or "header_distance" in page_style:
        page_setup.setHeaderDistance(page_style.get("header_istance") or page_style.get("header_distance"))
    if "footer_distance" in page_style:
        page_setup.setFooterDistance(page_style["footer_distance"])
    # 设置文档网络
    page_setup.setLayoutMode(SectionLayoutMode.fromName(page_style["layout_mode"]))
    try:
        page_setup.setLinesPerPage(page_style["lines_per_page"])
    except Exception as e:
        logger.info(f"设置LinesPerPage异常: {str(e)}")
        max_num_match = re.search(r'and (\d+)', e.args[0])
        if max_num_match:
            max_num = int(max_num_match.group(1))
            page_setup.setLinesPerPage(max_num)
            logger.info(f"设置LinesPerPage为最大值{max_num}")
        else:
            page_setup.setLinesPerPage(page_style["lines_per_page"])
    if page_style["layout_mode"] in ("GRID", "SNAP_TO_CHARS"):
        page_setup.setCharactersPerLine(page_style.get("characters_per_line", 0))
    return document


def parse_footnote(document):
    '''
    解析脚注的格式
    @param document:
    @return:
    '''
    footnotes = document.getFootnoteOptions()
    footnote_type = {
        # https://reference.aspose.com/words/zh/java/com.aspose.words/numberstyle/
        "number_style": str(NumberStyle.getName(footnotes.getNumberStyle())),  # 自动编号脚注的数字格式(ARABIC	阿拉伯数字 (1, 2, 3, …))
        # https://reference.aspose.com/words/zh/java/com.aspose.words/footnoteposition/
        "position": str(FootnotePosition.getName(footnotes.getPosition())),  # 脚注位置 (BOTTOM_OF_PAGE:脚注输出在每页的底部)
        # https://reference.aspose.com/words/zh/java/com.aspose.words/footnotenumberingrule/
        "restart_rule": str(FootnoteNumberingRule.getName(footnotes.getRestartRule())),  # 确定自动编号何时重新开始
        "start_number": footnotes.getStartNumber()
    }
    return footnote_type

def clear_temp_folder(temp_folder):
    if os.path.exists(temp_folder):
        for filename in os.listdir(temp_folder):
            file_path = os.path.join(temp_folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")