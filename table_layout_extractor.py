"""
表格视觉信息提取器
用于从PDF/DOCX中提取表格的视觉布局信息
"""

from typing import List, Dict, Any, Optional
import numpy as np


class TableVisualExtractor:
    """表格视觉信息提取器 - 提取和分析表格的视觉布局特征"""

    def __init__(self):
        pass

    @staticmethod
    def extract_cell_properties_pdf(cells: List, table_data: List[List[str]]) -> Dict[str, Any]:
        """
        从PyMuPDF的cells提取单元格属性

        Args:
            cells: PyMuPDF的cells列表
            table_data: 表格文本数据

        Returns:
            包含视觉属性矩阵的字典
        """
        if not cells or not table_data:
            return {"visual_matrix": [], "merged_cells": [], "properties": {}}

        rows = len(table_data)
        cols = max(len(row) for row in table_data) if table_data else 0

        # 初始化视觉矩阵
        visual_matrix = []
        for r in range(rows):
            row_visual = []
            for c in range(cols):
                row_visual.append({
                    'row': r,
                    'col': c,
                    'row_span': 1,
                    'col_span': 1,
                    'is_merged': False,
                    'merge_type': []
                })
            visual_matrix.append(row_visual)

        merged_cells = []

        # 分析单元格合并情况
        for cell in cells:
            if len(cell) >= 4:
                y1, x1, y2, x2 = cell

                row_span = y2 - y1
                col_span = x2 - x1

                if row_span > 1 or col_span > 1:
                    merge_type = []
                    if row_span > 1:
                        merge_type.append('vertical')
                    if col_span > 1:
                        merge_type.append('horizontal')

                    merge_info = {
                        'start_row': y1,
                        'start_col': x1,
                        'end_row': y2 - 1,
                        'end_col': x2 - 1,
                        'row_span': row_span,
                        'col_span': col_span,
                        'merge_type': merge_type
                    }
                    merged_cells.append(merge_info)

                    if y1 < rows and x1 < cols:
                        visual_matrix[y1][x1].update({
                            'row_span': row_span,
                            'col_span': col_span,
                            'is_merged': True,
                            'merge_type': merge_type
                        })

        properties = {
            'total_rows': rows,
            'total_cols': cols,
            'merged_cell_count': len(merged_cells),
            'has_vertical_merge': any('vertical' in mc['merge_type'] for mc in merged_cells),
            'has_horizontal_merge': any('horizontal' in mc['merge_type'] for mc in merged_cells),
            'max_row_span': max((mc['row_span'] for mc in merged_cells), default=1),
            'max_col_span': max((mc['col_span'] for mc in merged_cells), default=1)
        }

        return {
            'visual_matrix': visual_matrix,
            'merged_cells': merged_cells,
            'properties': properties
        }

    @staticmethod
    def extract_cell_properties_docx(table) -> Dict[str, Any]:
        """
        从python-docx的table对象提取单元格属性

        Args:
            table: python-docx的Table对象

        Returns:
            包含视觉属性矩阵的字典
        """
        try:
            from docx.oxml.ns import qn
        except ImportError:
            return {"visual_matrix": [], "merged_cells": [], "properties": {}}

        visual_matrix = []
        merged_cells = []

        rows = len(table.rows)
        cols = 0
        for row in table.rows:
            total_span = sum(
                int(cell._tc.find(qn('w:tcPr')).find(qn('w:gridSpan')).get(qn('w:val'), 1))
                if cell._tc.find(qn('w:tcPr')) is not None
                   and cell._tc.find(qn('w:tcPr')).find(qn('w:gridSpan')) is not None
                else 1
                for cell in row.cells
            )
            cols = max(cols, total_span)

        for row_idx, row in enumerate(table.rows):
            row_visual = []
            for cell_idx, cell in enumerate(row.cells):
                tc = cell._tc
                tcPr = tc.find(qn('w:tcPr'))

                grid_span = 1
                v_merge = None

                if tcPr is not None:
                    grid_span_elem = tcPr.find(qn('w:gridSpan'))
                    if grid_span_elem is not None:
                        grid_span = int(grid_span_elem.get(qn('w:val'), 1))

                    v_merge_elem = tcPr.find(qn('w:vMerge'))
                    if v_merge_elem is not None:
                        v_merge_val = v_merge_elem.get(qn('w:val'), 'continue')
                        v_merge = v_merge_val

                is_merged = grid_span > 1 or v_merge is not None
                merge_type = []
                if grid_span > 1:
                    merge_type.append('horizontal')
                if v_merge is not None:
                    merge_type.append('vertical')

                cell_visual = {
                    'row': row_idx,
                    'col': cell_idx,
                    'row_span': 1,
                    'col_span': grid_span,
                    'is_merged': is_merged,
                    'merge_type': merge_type
                }

                row_visual.append(cell_visual)

                if is_merged:
                    merged_cells.append({
                        'start_row': row_idx,
                        'start_col': cell_idx,
                        'row_span': 1,
                        'col_span': grid_span,
                        'merge_type': merge_type
                    })

            visual_matrix.append(row_visual)

        properties = {
            'total_rows': rows,
            'total_cols': cols,
            'merged_cell_count': len(merged_cells),
            'has_vertical_merge': any('vertical' in mc.get('merge_type', []) for mc in merged_cells),
            'has_horizontal_merge': any('horizontal' in mc.get('merge_type', []) for mc in merged_cells),
            'max_row_span': 1,
            'max_col_span': max((mc.get('col_span', 1) for mc in merged_cells), default=1)
        }

        return {
            'visual_matrix': visual_matrix,
            'merged_cells': merged_cells,
            'properties': properties
        }

    @staticmethod
    def analyze_layout_patterns(visual_info: List[List[Dict]]) -> Dict[str, Any]:
        """分析表格的布局模式"""
        if not visual_info:
            return {'pattern': 'unknown', 'characteristics': []}

        rows = len(visual_info)
        cols = len(visual_info[0]) if visual_info else 0

        characteristics = []

        merged_count = 0
        total_cells = 0
        for row in visual_info:
            for cell in row:
                total_cells += 1
                if isinstance(cell, dict):
                    row_span = cell.get('row_span', 1)
                    col_span = cell.get('col_span', 1)
                    if isinstance(col_span, str):
                        try:
                            col_span = int(col_span)
                        except:
                            col_span = 1
                    if row_span > 1 or col_span > 1:
                        merged_count += 1

        merge_ratio = merged_count / total_cells if total_cells > 0 else 0

        if merge_ratio < 0.1:
            characteristics.append('规则矩阵布局')
            pattern = 'regular_matrix'
        elif merge_ratio < 0.3:
            characteristics.append('部分合并布局')
            pattern = 'partial_merged'
        else:
            characteristics.append('复杂合并布局')
            pattern = 'complex_merged'

        has_wide_merge = False
        for row in visual_info[:3]:
            for cell in row:
                if isinstance(cell, dict):
                    col_span = cell.get('col_span', 1)
                    if isinstance(col_span, str):
                        try:
                            col_span = int(col_span)
                        except:
                            col_span = 1
                    if col_span >= 3:
                        has_wide_merge = True
                        characteristics.append('存在宽表头合并区域')
                        break

        has_tall_merge = False
        for row in visual_info:
            for cell in row:
                if isinstance(cell, dict):
                    row_span = cell.get('row_span', 1)
                    if row_span >= 2:
                        has_tall_merge = True
                        characteristics.append('存在纵向合并区域（可能为备注）')
                        break
            if has_tall_merge:
                break

        return {
            'pattern': pattern,
            'characteristics': characteristics,
            'merge_ratio': merge_ratio,
            'has_wide_merge': has_wide_merge,
            'has_tall_merge': has_tall_merge,
            'rows': rows,
            'cols': cols
        }


def create_visual_description(visual_info: List[List[Dict]], format_type: str = 'pdf') -> str:
    """
    创建视觉布局的文本描述（用于输入LLM）

    Args:
        visual_info: 视觉信息矩阵
        format_type: 来源格式 ('pdf' 或 'docx')

    Returns:
        视觉布局的文本描述
    """
    if not visual_info:
        return "无视觉布局信息"

    extractor = TableVisualExtractor()
    layout_analysis = extractor.analyze_layout_patterns(visual_info)

    rows = layout_analysis['rows']
    cols = layout_analysis['cols']

    description_parts = [
        f"表格视觉布局分析（{format_type.upper()}格式）：",
        f"- 总行数: {rows}, 总列数: {cols}",
        f"- 布局模式: {layout_analysis['pattern']}",
        f"- 特征: {', '.join(layout_analysis['characteristics']) if layout_analysis['characteristics'] else '无明显特征'}"
    ]

    merged_details = []
    for row_idx, row in enumerate(visual_info):
        for col_idx, cell in enumerate(row):
            if isinstance(cell, dict):
                row_span = cell.get('row_span', 1)
                col_span = cell.get('col_span', 1)

                if isinstance(col_span, str):
                    try:
                        col_span = int(col_span)
                    except:
                        col_span = 1

                if row_span > 1 or col_span > 1:
                    detail = f"  单元格[{row_idx},{col_idx}]: "
                    if row_span > 1:
                        detail += f"跨{row_span}行 "
                    if col_span > 1:
                        detail += f"跨{col_span}列"
                    merged_details.append(detail)

    if merged_details:
        description_parts.append(f"- 合并单元格数: {len(merged_details)}")
        if len(merged_details) <= 10:
            description_parts.append("- 合并单元格详情:")
            description_parts.extend(merged_details)
        else:
            description_parts.append(f"- 合并单元格详情（前10个）:")
            description_parts.extend(merged_details[:10])
            description_parts.append(f"  ... 还有 {len(merged_details) - 10} 个")

    return '\n'.join(description_parts)