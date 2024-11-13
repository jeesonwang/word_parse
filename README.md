# **Word解析**

Python==3.9.2



## 接口输入参数：

filepath：输入的s3中的word文件路径；

pdf_filepath：输出的未染色的pdf文件存储路径；

coloring_pdf_filepath：输出的染色后的pdf文件存储路径。

请求示例：

```shell
curl --location 'http://120.0.0.0:5000/coloring_paragraph' \

--header 'Content-Type: application/json' \

--data '{

​	"filepath": "example-bucket/upload/0/1feca742172a.docx",

​	"pdf_filepath":"example-bucket/upload/0/1feca742172a.pdf",

​	"coloring_pdf_filepath": "example-bucket/upload/0/1feca742172a_coloring.pdf"

​	}'
```



## 输出响应示例：

```json
{

  "code": 0,  // 0为成功

  "message": "请求成功",  // 消息任意,如失败需体现报错情况

  "data": {

​      {

​      "title": None,

​      "paragraph": [

​            {

​            "type_": "正文",

​            "title": None,

​            "content": "发行主体概况",

​            "assist": "",

​            "markdown_content": "发行主体概况",

​            "color": "#c77dc2"

​            "table_detail": None

​            },

​            {

​            "type_": "table",

​            "title": null,

​            "content": "纵向合并 1 3  2 3  横向合并 3  纵向加横向合并 3  7 6  ",

​                        "assist": "<table cellpadding=\"0\" cellspacing=\"0\"><tr><td rowspan=\"2\">纵向合并</td><td>1</td><td>3</td></tr><tr><td>2</td><td>3</td></tr><tr><td colspan=\"2\">横向合并</td><td>3</td></tr><tr><td colspan=\"2\">纵向加横向合并</td><td>3</td></tr><tr><td>7</td><td colspan=\"2\">6</td></tr></table>",

​            "markdown_content": "| 纵向合并 | 1 | 3 |\n| --- | --- | --- |\n|  | 2 | 3 |\n| 横向合并 |  | 3 |\n| 纵向加横向合并 |  | 3 |\n| 7 |  | 6 |",

​            "color": None

​            "table_detail": ...

​    }

​        ]

   }

}
```

 

**"table_detail"中的详细内容如下：**

```json
{
  "is_sub_tabel": False, // 是不是嵌套的子表格

  "row_num": "32" // 这个表格有32行

  "columns": ["column1", "column2", ...]   // 所有列名

  "cells": [

​			{

​				"column": ["column1"]   // 当前单元格所在的列

​				"HorizontalMerge": "FIRST",  // HorizontalMerge和VerticalMerge字段说明below

​				"VerticalMerge": "NONE",

​				"Position": [0, 0]

​				"paragraph": [

​								{
                    ​            "type_": "正文",

                    ​            "title": None,

                    ​            "content": "发行主体概况",

                    ​            "assist": "",

                    ​            "markdown_content": "发行主体概况",

                    ​            "color": "#c77dc2"

                    ​            "table_detail": None

                    ​            },

                    ​            {

                    ​            "type_": "table",

                    ​            "title": null,

                    ​            "content": "...",

                    ​            "assist": "<table></table>",

                    ​            "markdown_content": "|...|",

                    ​            "color": None

                    ​            "table_detail": ...

                    ​    }
​								

​						]

​				},

​				{

​				"column": ["column1"]

​				"HorizontalMerge": "PREVIOUS",

​				"VerticalMerge": "NONE",

​				"Position": [0, 1]

​				"paragraph": [

​								{
                    ​            "type_": "...",

                    ​            "title": None,

                    ​            "content": "...",

                    ​            "assist": "",

                    ​            "markdown_content": "...",

                    ​            "color": "#c113c2"

                    ​            "table_detail": None

                    ​            },

​								{
                    ​            "type_": "...",

                    ​            "title": None,

                    ​            "content": "...",

                    ​            "assist": "",

                    ​            "markdown_content": "...",

                    ​            "color": "#c22dc2"

                    ​            "table_detail": None

                    ​            }

​							]

​				}

​			]
}
```

####  **HorizontalMerge和VerticalMerge值说明：**

HorizontalMerge和VerticalMerge有三种值：

| 名字     | 说明                                           | 对应值 |
| -------- | ---------------------------------------------- | ------ |
| NONE     | 单元格未合并。                                 | 0      |
| FIRST    | 该单元格是合并单元格范围内的第一个单元格。     | 1      |
| PREVIOUS | 该单元格在水平或垂直方向上与前一个单元格合并。 | 2      |

 

