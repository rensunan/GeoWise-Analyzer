"""
表格去重模块 - 硬编码实现去除重复行和重复列
"""

from typing import List, Dict, Any


class TableDeduplicator:
    """表格去重器 - 使用硬编码方式去除重复行和列"""

    def __init__(self):
        pass

    @staticmethod
    def deduplicate_rows(table_data: List[List[str]]) -> List[List[str]]:
        """
        去除完全相同的重复行

        Args:
            table_data: 原始表格数据

        Returns:
            去重后的表格数据（保持原有顺序，保留第一次出现）
        """
        if not table_data:
            return table_data

        result = []
        seen_rows = []

        for row in table_data:
            # 将行转换为可哈希的元组
            row_tuple = tuple(row)

            # 如果是首次出现，则保留
            if row_tuple not in seen_rows:
                seen_rows.append(row_tuple)
                result.append(row)

        return result

    @staticmethod
    def deduplicate_columns(table_data: List[List[str]]) -> List[List[str]]:
        """
        去除完全相同的重复列

        Args:
            table_data: 原始表格数据

        Returns:
            去重后的表格数据（保持原有顺序，保留第一次出现）
        """
        if not table_data or len(table_data) == 0:
            return table_data

        # 确保所有行长度一致
        max_cols = max(len(row) for row in table_data) if table_data else 0
        padded_data = []
        for row in table_data:
            padded_row = row + [''] * (max_cols - len(row))
            padded_data.append(padded_row)

        # 标记需要保留的列
        cols_to_keep = []
        seen_columns = []

        for j in range(max_cols):
            # 提取第j列的所有值
            column = [padded_data[i][j] for i in range(len(padded_data))]
            column_tuple = tuple(column)

            # 如果是首次出现，则保留
            if column_tuple not in seen_columns:
                seen_columns.append(column_tuple)
                cols_to_keep.append(j)

        # 构建新表格，只保留需要的列
        result = []
        for i in range(len(padded_data)):
            new_row = [padded_data[i][j] for j in cols_to_keep]
            # 去除行尾的空字符串（可选）
            while new_row and new_row[-1] == '':
                new_row.pop()
            result.append(new_row)

        return result

    @staticmethod
    def deduplicate_all(table_data: List[List[str]],
                        remove_duplicate_rows: bool = True,
                        remove_duplicate_cols: bool = True) -> Dict[str, Any]:
        """
        同时去除重复行和重复列

        Args:
            table_data: 原始表格数据
            remove_duplicate_rows: 是否去除重复行
            remove_duplicate_cols: 是否去除重复列

        Returns:
            包含去重结果和统计信息的字典
        """
        if not table_data:
            return {
                "optimized_table": table_data,
                "stats": {
                    "rows_removed": 0,
                    "cols_removed": 0,
                    "original_rows": 0,
                    "original_cols": 0,
                    "final_rows": 0,
                    "final_cols": 0
                }
            }

        original_rows = len(table_data)
        original_cols = max(len(row) for row in table_data) if table_data else 0

        result = table_data
        rows_removed = 0
        cols_removed = 0

        # 先去重行
        if remove_duplicate_rows:
            before_rows = len(result)
            result = TableDeduplicator.deduplicate_rows(result)
            rows_removed = before_rows - len(result)

        # 再去重列
        if remove_duplicate_cols:
            before_cols = max(len(row) for row in result) if result else 0
            result = TableDeduplicator.deduplicate_columns(result)
            after_cols = max(len(row) for row in result) if result else 0
            cols_removed = before_cols - after_cols

        final_rows = len(result)
        final_cols = max(len(row) for row in result) if result else 0

        return {
            "optimized_table": result,
            "has_optimization": rows_removed > 0 or cols_removed > 0,
            "stats": {
                "rows_removed": rows_removed,
                "cols_removed": cols_removed,
                "original_rows": original_rows,
                "original_cols": original_cols,
                "final_rows": final_rows,
                "final_cols": final_cols
            }
        }

    @staticmethod
    def deduplicate_with_header_protection(table_data: List[List[str]]) -> Dict[str, Any]:
        """
        去重时保护表头行（第一行）

        Args:
            table_data: 原始表格数据

        Returns:
            去重后的表格数据和统计信息
        """
        if not table_data or len(table_data) < 2:
            return {
                "optimized_table": table_data,
                "has_optimization": False,
                "stats": {"rows_removed": 0, "cols_removed": 0}
            }

        # 分离表头和数据行
        header = [table_data[0]]
        data_rows = table_data[1:]

        # 对数据行去重
        dedup_result = TableDeduplicator.deduplicate_all(data_rows)
        optimized_data = dedup_result["optimized_table"]

        # 重新组合
        optimized_table = header + optimized_data

        return {
            "optimized_table": optimized_table,
            "has_optimization": dedup_result["has_optimization"],
            "stats": dedup_result["stats"]
        }