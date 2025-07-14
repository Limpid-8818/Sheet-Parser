import os
from string import Template

import pandas as pd
from openpyxl import load_workbook
import csv
from utils import check_file_exists, check_file_format, get_default_title, determine_data_type


class SheetParser:
    """表格解析器，用于将多种格式的表格文件转换为HTML表格"""

    def __init__(self):
        """初始化解析器"""
        self.supported_formats = ['.xlsx', '.xls', '.csv']
        self.html_template_path = 'templates/basic_table.html'
        with open(self.html_template_path, 'r', encoding='utf-8') as f:
            self.html_template = Template(f.read())

    def parse_file(self, file_path, output_file=None, title=None):
        """
        解析表格文件并生成HTML

        Args:
            file_path: 表格文件路径
            output_file: 输出HTML文件路径，默认为None，此时返回HTML字符串
            title: HTML页面标题，默认为文件名

        Returns:
            生成的HTML字符串(如果output_file为None)
        """
        check_file_exists(file_path)

        check_file_format(file_path, self.supported_formats)

        if title is None:
            title = get_default_title(file_path)

        # 根据文件格式选择解析方法
        if os.path.splitext(file_path)[1].lower() in ['.xlsx', '.xls']:
            sheets_data = self._parse_excel(file_path)
        elif os.path.splitext(file_path)[1].lower() == '.csv':
            sheets_data = self._parse_csv(file_path)

        # 生成HTML内容
        content = self._generate_html_content(sheets_data)
        html = self.html_template.substitute(title=title, content=content)

        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html)
            return f"HTML已保存到: {output_file}"
        else:
            return html

    def _parse_excel(self, file_path):
        """解析Excel文件(.xlsx, .xls)"""
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        sheets_data = []

        # 使用openpyxl获取合并单元格信息
        wb = load_workbook(file_path, read_only=True, data_only=True)

        for sheet_name in sheet_names:
            df = xls.parse(sheet_name)

            # 处理空数据框
            if df.empty:
                sheets_data.append({
                    'name': sheet_name,
                    'data': [],
                    'merged_cells': []
                })
                continue

            # 获取合并单元格信息
            ws = wb[sheet_name]
            merged_cells = []
            if hasattr(ws, 'merged_cells'):
                for merged_range in ws.merged_cells.ranges:
                    min_col, min_row, max_col, max_row = merged_range.bounds
                    merged_cells.append({
                        'min_row': min_row - 1,  # 转换为0-based索引
                        'min_col': min_col - 1,
                        'max_row': max_row - 1,
                        'max_col': max_col - 1
                    })

            # 处理数据类型和空值
            data = []
            for row_idx, row in df.iterrows():
                processed_row = []
                for col_idx, value in enumerate(row):
                    # 处理NaN值
                    if pd.isna(value):
                        cell_value = ''
                        cell_type = 'string'
                    else:
                        # 确定数据类型
                        cell_type = determine_data_type(value)
                        # 格式化日期
                        if cell_type == 'date':
                            cell_value = value.strftime('%Y-%m-%d')
                        else:
                            cell_value = str(value)

                    # 检查是否为合并单元格
                    is_merged = False
                    for merged_cell in merged_cells:
                        if (row_idx >= merged_cell['min_row'] and row_idx <= merged_cell['max_row'] and
                                col_idx >= merged_cell['min_col'] and col_idx <= merged_cell['max_col']):
                            # 只有左上角的单元格显示内容
                            if row_idx == merged_cell['min_row'] and col_idx == merged_cell['min_col']:
                                is_merged = True
                                processed_row.append({
                                    'value': cell_value,
                                    'type': cell_type,
                                    'rowspan': merged_cell['max_row'] - merged_cell['min_row'] + 1,
                                    'colspan': merged_cell['max_col'] - merged_cell['min_col'] + 1,
                                    'is_merged': True
                                })
                            else:
                                # 其他合并单元格位置设置为空
                                processed_row.append({
                                    'value': '',
                                    'type': 'string',
                                    'rowspan': 1,
                                    'colspan': 1,
                                    'is_merged': False,
                                    'skip': True  # 标记为跳过渲染
                                })
                            break

                    # 如果不是合并单元格或合并单元格的左上角
                    if not is_merged:
                        processed_row.append({
                            'value': cell_value,
                            'type': cell_type,
                            'rowspan': 1,
                            'colspan': 1,
                            'is_merged': False
                        })

                data.append(processed_row)

            sheets_data.append({
                'name': sheet_name,
                'data': data,
                'merged_cells': merged_cells
            })

        return sheets_data

    def _parse_csv(self, file_path):
        """解析CSV文件"""
        sheets_data = []

        # 读取CSV文件
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            data = []

            for row_idx, row in enumerate(reader):
                processed_row = []
                for col_idx, value in enumerate(row):
                    # 确定数据类型
                    cell_type = determine_data_type(value)

                    processed_row.append({
                        'value': value,
                        'type': cell_type,
                        'rowspan': 1,
                        'colspan': 1,
                        'is_merged': False
                    })

                data.append(processed_row)

        # CSV文件只有一个工作表
        sheets_data.append({
            'name': os.path.basename(file_path),
            'data': data,
            'merged_cells': []  # CSV不支持合并单元格
        })

        return sheets_data

    def _generate_html_content(self, sheets_data):
        """生成HTML内容"""
        content = ''

        for sheet in sheets_data:
            sheet_name = sheet['name']
            sheet_data = sheet['data']

            # 生成表格HTML
            table_html = '<div class="sheet-container">'
            table_html += f'<h2 class="sheet-title">{sheet_name}</h2>'
            table_html += '<table class="sheet-table">'

            # 生成表头（假设第一行为表头）
            if sheet_data:
                table_html += '<thead><tr>'
                header_row = sheet_data[0]
                for cell in header_row:
                    if 'skip' in cell and cell['skip']:
                        continue

                    cell_class = ''
                    if cell['is_merged']:
                        cell_class = 'merged-cell'

                    table_html += f'<th colspan="{cell["colspan"]}" rowspan="{cell["rowspan"]}" class="{cell_class}">{cell["value"]}</th>'

                table_html += '</tr></thead><tbody>'

                # 生成表格内容（从第二行开始）
                for row_idx, row in enumerate(sheet_data[1:], 1):
                    table_html += '<tr>'
                    for cell in row:
                        if 'skip' in cell and cell['skip']:
                            continue

                        cell_class = cell['type'] + '-cell'
                        if cell['is_merged']:
                            cell_class += ' merged-cell'

                        table_html += f'<td colspan="{cell["colspan"]}" rowspan="{cell["rowspan"]}" class="{cell_class}">{cell["value"]}</td>'

                    table_html += '</tr>'

            table_html += '</tbody></table></div>'
            content += table_html

        return content
