import os

path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app3.py")
with open(path, "r", encoding="utf-8") as f:
    content = f.read()

old_method = '''    def _trim_by_consistency(self, table_data, adapted):
        """用数据一致性修剪主表边界：从底部向上收缩不符合列类型的行。"""
        if not adapted.get("has_main_table") or not adapted.get("main_table"):
            return adapted
        mt = adapted["main_table"]
        if mt["end_row"] <= mt["start_row"]:
            return adapted

        # 1. 从缓存区域分析列类型（用当前表格的实际数据）
        col_types = self._get_column_types(
            table_data, mt["start_row"], mt["end_row"],
            mt["start_col"], mt["end_col"]
        )
        # 如果没有任何 numeric/text 列，跳过
        if all(ct in ('mixed', 'empty') for ct in col_types):
            return adapted

        # 2. 从底部向上收缩
        old_end = mt["end_row"]
        trimmed_rows = []
        while mt["end_row"] > mt["start_row"]:
            row = table_data[mt["end_row"]]
            if self._row_matches_types(row, col_types, mt["start_col"]):
                break
            trimmed_rows.append(mt["end_row"])
            mt["end_row"] -= 1

        if trimmed_rows:
            trimmed_rows.reverse()
            print("  [consistency trim] removed {} row(s) from main bottom: {}"
                  .format(len(trimmed_rows), trimmed_rows))

        return adapted'''

new_method = '''    def _trim_by_consistency(self, table_data, adapted):
        """用数据一致性修正主表边界：从上往下扫描数据行，遇到第一行列类型不匹配就暂停，
           暂停点之后的行不再属于主表。"""
        if not adapted.get("has_main_table") or not adapted.get("main_table"):
            return adapted
        mt = adapted["main_table"]
        if mt["end_row"] <= mt["start_row"]:
            return adapted

        # 1. 分析当前主表区域的列类型
        col_types = self._get_column_types(
            table_data, mt["start_row"], mt["end_row"],
            mt["start_col"], mt["end_col"]
        )
        if all(ct in ('mixed', 'empty') for ct in col_types):
            return adapted

        # 2. 从表头下一行开始，自上而下找第一个不匹配的行
        old_end = mt["end_row"]
        for i in range(mt["start_row"] + 1, mt["end_row"] + 1):
            row = table_data[i]
            if not self._row_matches_types(row, col_types, mt["start_col"]):
                # 第一行不匹配 → 从这里断开，前面是主表，后面不是
                mt["end_row"] = i - 1
                print("  [consistency] main table stopped at row {}, column types broke at row {}"
                      .format(mt["end_row"], i))
                break

        return adapted'''

content = content.replace(old_method, new_method, 1)
print("Replaced" if old_method in content else "old_method NOT FOUND, trying partial match")

# Fallback if exact match fails
if "def _trim_by_consistency(self, table_data, adapted):" in content:
    # Already replaced above, check
    pass

with open(path, "w", encoding="utf-8") as f:
    f.write(content)
print("Done")
