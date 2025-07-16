## Sheet-Parser

腾讯犀牛鸟开源项目-腾讯文档

------

这是一个用 Python 实现的表格解析器 (Sheet Parser)，可以将 **Excel 表格**、**CSV 文件**解析为 **HTML 文件**。

#### 功能支持情况

| 功能                                     | 具体支持情况                                        |
| ---------------------------------------- | --------------------------------------------------- |
| `.xlsx` `.xls` `.csv` `.et` 等 10+ 格式  | 绝大多数 Excel 类文件格式均支持                     |
| 合并单元格（rowspan / colspan）          | 支持                                                |
| 单元格公式悬停提示                       | 支持，设置了公式的单元格会有蓝色框提示              |
| 单元格批注悬停提示                       | 支持，设置了批注的单元格会有黑色框提示              |
| 字体、字号、颜色、背景色                 | 字号、颜色、背景色均能正确解析，字体存在兼容性问题  |
| 字体样式                                 | 支持粗体、斜体、下划线等                            |
| 虚线 / 点线 / 双线边框 + 粗细            | 支持实线、双线、虚线、点线样式，支持调整粗细        |
| 超链接                                   | 支持，能正确识别并转换为\<a>标签                    |
| 自动类型识别（数字、日期、布尔、字符串） | 支持，识别到的单元格类型会被插入到\<td>标签的属性中 |

------

#### 项目结构

```
Sheet-Parser/
│
├── sheet_parser.py          # 核心解析器
├── utils.py                 # 通用工具函数
├── main.py					 # 程序入口（测试用）
├── templates/
│   └── basic_table.html     # HTML 模板及样式
├── examples/                # 示例文件（可自己创建）
│   ├── basic_example.xlsx
│   └── example_csv.csv
├── requirements.txt
└── README.md
```

#### 项目依赖

```
openpyxl>=3.1.5
pandas>=2.2.3
```

------

#### 核心 API

```python
parse_file(
    file_path: str,
    output_file: str | None = None,
    title: str | None = None
) -> str
```

- **file_path**  待解析文件
- **output_file**  如果给出，则写入文件并返回提示；否则返回 HTML 字符串
- **title**  页面标题，默认使用文件名

**示例**

```python
from sheet_parser import SheetParser

parser = SheetParser()

# 直接返回 HTML 字符串
html = parser.parse_file('demo.xlsx')

# 或保存成文件
parser.parse_file('demo.xlsx', output_file='demo.html')
```

------

#### 解析效果

（Markdown 内嵌 HTML 不能完全反映 HTML 效果，具体效果请使用浏览器打开示例文件查看）

`example.html`：

<h1>示例表格</h1>
<div class="sheet-container"><h2 class="sheet-title">example</h2><table class="sheet-table"><tbody><tr><td class="" colspan="1" rowspan="1" style="font-size: 10.5pt; vertical-align: center">普通文本</td><td class="numeric-cell" colspan="1" rowspan="1" style="font-size: 10.5pt; vertical-align: center">12345</td><td class="date-cell" colspan="1" rowspan="1" style="font-size: 10.5pt; vertical-align: center">2025-07-15</td><td class="boolean-cell" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">True</td><td class="" colspan="1" rowspan="1" style="text-decoration: underline; color: #0000FF; font-size: 11.0pt; vertical-align: center"><a href="https://docs.qq.com/aio/DTk1wUUFHUkZCQkZN?p=kBrlITmKIZYap9DtpMNaox" target="_blank">2025腾讯犀牛鸟研学基地</a></td></tr><tr><td class="" colspan="1" rowspan="1" style="font-weight: bold; font-size: 11.0pt; vertical-align: center">加粗</td><td class="" colspan="1" rowspan="1" style="font-style: italic; font-size: 11.0pt; vertical-align: center">斜体</td><td class="" colspan="1" rowspan="1" style="text-decoration: underline; font-size: 11.0pt; vertical-align: center">下划线</td><td class="" colspan="1" rowspan="1" style="color: #FF0000; font-size: 11.0pt; vertical-align: center">红色字体</td><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; background-color: #FFFF00; vertical-align: center">黄色背景</td></tr><tr><td class="merged-cell" colspan="3" rowspan="1" style="font-size: 11.0pt; border-left: solid 1px black; border-right: solid 1px black; border-top: solid 1px black; border-bottom: solid 1px black; text-align: center; vertical-align: center">合并A3:C3</td><td class="merged-cell" colspan="2" rowspan="2" style="font-size: 11.0pt; border-left: solid 1px black; border-right: solid 1px black; border-top: solid 1px black; border-bottom: solid 1px black; text-align: center; vertical-align: center">合并D3:E4</td></tr><tr><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; border-top: solid 3px black; vertical-align: center">粗上框线</td><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; border-right: double 3px black; vertical-align: center">双右框线</td><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; border-bottom: dashed 1px black; vertical-align: center">虚下框线</td></tr><tr><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">公式</td><td class="numeric-cell    formula-cell" data-formula="=B1" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">12345</td><td class="date-cell   formula-cell" data-formula="=TODAY()" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">2025-07-15</td><td class="numeric-cell" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">100</td><td class="formula-cell" data-formula="=IF(D5&gt;50,&quot;大&quot;,&quot;小&quot;)" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">大</td></tr><tr><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">批注</td><td class="commented-cell" data-comment="user:
这是批注" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center"><span class="comment" data-comment="user:
这是批注">此处有批注</span></td><td class="commented-cell formula-cell" data-comment="user:
这是批注，这里也有公式" data-formula="=B6" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center"><span class="comment" data-comment="user:
这是批注，这里也有公式">此处有批注</span></td><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; border-left: solid 1px black; border-right: solid 1px black; border-top: solid 1px black; border-bottom: solid 1px black; text-align: center; vertical-align: center">文字居中</td><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; border-left: solid 1px black; border-right: solid 1px black; border-top: solid 1px black; border-bottom: solid 1px black; text-align: right; vertical-align: center">文字居右</td></tr><tr><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">空值</td><td class="formula-cell" data-formula="=NA()" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center"></td><td class="formula-cell" data-formula="=&quot;&quot;" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center"></td><td class="numeric-cell" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">0</td><td class="" colspan="1" rowspan="1" style="font-size: 10.5pt; vertical-align: center"></td></tr></tbody></table></div><div class="sheet-container"><h2 class="sheet-title">another</h2><table class="sheet-table"><tbody><tr><td class="" colspan="1" rowspan="1" style="font-size: 11.0pt; vertical-align: center">这是另一个sheet</td></tr></tbody></table></div>
`example_csv.html`：

<h1>CSV示例表格</h1>
<div class="sheet-container"><h2 class="sheet-title">example_csv.csv</h2><table class="sheet-table"><thead><tr><th class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Name</th><th class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Age</th><th class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">JoinDate</th><th class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">IsActive</th><th class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Salary</th></tr></thead><tbody><tr><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Alice</td><td class="numeric-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">25</td><td class="date-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">2023-11-01</td><td class="boolean-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">true</td><td class="numeric-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">1234.56</td></tr><tr><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Bob</td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;"></td><td class="date-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">2024-01-15</td><td class="boolean-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">false</td><td class="numeric-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">0</td></tr><tr><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Charlie</td><td class="numeric-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">30</td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;"></td><td class="boolean-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">TRUE</td><td class="numeric-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">12345.67</td></tr><tr><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Diana</td><td class="numeric-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">28</td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">2024/07/16</td><td class="boolean-cell" colspan="1" rowspan="1" style="border: 1px solid #ddd;">FALSE</td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;"></td></tr><tr><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;"></td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;"></td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;"></td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;"></td><td class="" colspan="1" rowspan="1" style="border: 1px solid #ddd;">Empty row</td></tr></tbody></table></div>

------

#### 待改进

| 事项                  | 说明                                |
|---------------------|-----------------------------------|
| **支持多工作表导航**        | 目前所有工作表顺序堆叠，需增加标签页/目录快速跳转         |
| **CSV 自定义分隔符 / 编码** | 目前仅支持 UTF-8 与逗号分隔，增加参数传入          |
| **大文件解析优化**         | 目前未对大文件解析进行优化，耗时长，且会导致前端渲染较慢，考虑分页 |
| **图片 / 图表**         | openpyxl 可读取图片锚点，但目前暂未实现图片图表的渲染   |
| **CLI 入口**          | 提供命令行调用，适应更多场合                    |
| **主题切换**            | 目前仅支持浅色主题，考虑内置明暗两套 CSS 主题，适应深色模式  |
