"""
表格布局信息提取器
从表格数据中检测合并单元格，格式化为提示词文本
"""

from typing import List, Dict


class TableLayoutExtractor:
    """表格布局信息提取器"""

    @staticmethod
    def extract_from_table(table_data: List[List[str]]) -> List[Dict]:
        """
        从表格数据中检测合并单元格。
        横向：连续相同内容视为合并，只记录起始位置，最大跨度优先。
        纵向：连续相同内容视为合并，只记录起始位置，最大跨度优先。
        """
        if not table_data:
            return []

        rows = len(table_data)
        cols = max(len(row) for row in table_data) if table_data else 0
        covered = [[False for _ in range(cols)] for _ in range(rows)]
        merged_cells = []

        # 横向合并
        for r in range(rows):
            c = 0
            while c < cols:
                if c >= len(table_data[r]) or covered[r][c]:
                    c += 1
                    continue
                content = table_data[r][c]
                if not content or content.strip() == "":
                    c += 1
                    continue
                # 向右找连续相同
                end_c = c
                while end_c + 1 < cols and end_c + 1 < len(table_data[r]):
                    if table_data[r][end_c + 1] == content:
                        end_c += 1
                    else:
                        break
                if end_c > c:
                    span = end_c - c + 1
                    merged_cells.append({'row': r, 'col': c, 'row_span': 1, 'col_span': span})
                    for cc in range(c, end_c + 1):
                        covered[r][cc] = True
                    c = end_c + 1
                else:
                    c += 1

        # 纵向合并（跳过已覆盖）
        for c in range(cols):
            r = 0
            while r < rows:
                if c >= len(table_data[r]) or covered[r][c]:
                    r += 1
                    continue
                content = table_data[r][c]
                if not content or content.strip() == "":
                    r += 1
                    continue
                # 向下找连续相同
                end_r = r
                while end_r + 1 < rows and c < len(table_data[end_r + 1]):
                    if not covered[end_r + 1][c] and table_data[end_r + 1][c] == content:
                        end_r += 1
                    else:
                        break
                if end_r > r:
                    span = end_r - r + 1
                    merged_cells.append({'row': r, 'col': c, 'row_span': span, 'col_span': 1})
                    for rr in range(r, end_r + 1):
                        covered[rr][c] = True
                    r = end_r + 1
                else:
                    r += 1

        return merged_cells

    @staticmethod
    def extract_from_docx(table) -> List[Dict]:
        """从DOCX表格提取合并单元格信息"""
        try:
            from docx.oxml.ns import qn
        except ImportError:
            return []

        merged_cells = []
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                tc = cell._tc
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is not None:
                    grid_span_elem = tcPr.find(qn('w:gridSpan'))
                    grid_span = int(grid_span_elem.get(qn('w:val'), 1)) if grid_span_elem is not None else 1
                    if grid_span > 1:
                        merged_cells.append({
                            'row': row_idx, 'col': cell_idx,
                            'row_span': 1, 'col_span': grid_span
                        })
        return merged_cells


def format_merged_cells(merged_cells: List[Dict]) -> str:
    """将合并单元格信息格式化为提示词文本"""
    if not merged_cells:
        return ""

    lines = ["【当前表格的原始合并信息】："]
    for mc in merged_cells:
        lines.append(f"- 单元格[{mc['row']},{mc['col']}]原始跨{mc['row_span']}行{mc['col_span']}列")
    return "\n".join(lines) + "\n"