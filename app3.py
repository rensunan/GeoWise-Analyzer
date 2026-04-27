"""
- 岩土报告智能解析系统
基于大语言模型的表格解析 - 主表/备注/附表识别
"""

import os
import re
import json
import uuid
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any
from collections import defaultdict
import threading
import warnings
import hashlib
from table_header_alignment import get_header_aligner, align_table_headers, align_and_merge_units

warnings.filterwarnings('ignore')

from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from openai import OpenAI

app = Flask(__name__)
CORS(app)

app.config['UPLOAD_FOLDER'] = './uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['SECRET_KEY'] = 'geotechnical-parser-secret-key'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('./results', exist_ok=True)

# DeepSeek API配置
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-db068eaf4d794e33a0d452203e4d8e9a")
DEEPSEEK_BASE_URL = "https://api.deepseek.com"

# 初始化OpenAI客户端
client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url=DEEPSEEK_BASE_URL
)

# ============================================================================
# 文档解析器
# ============================================================================

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    print("警告: python-docx未安装")

try:
    import fitz
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    print("警告: PyMuPDF未安装")


class DocumentParser:
    def __init__(self):
        self.content_keywords = [
            '黏聚力', '内摩擦角', '压缩模量', '孔隙比', '含水量', '液限', '塑限',
            '颗粒级配', '粒径', '筛分', '剪切', '压缩', '载荷', '承载力'
        ]
        self.skip_keywords = ['目录', 'Contents', '图', 'Fig', '表目录', '图目录', '参考文献']

    def parse_document(self, file_path: str) -> Tuple[List[Dict], List[List[List[str]]], int]:
        file_ext = Path(file_path).suffix.lower()
        if file_ext == '.docx':
            return self._parse_docx(file_path)
        elif file_ext == '.pdf':
            return self._parse_pdf(file_path)
        else:
            return self._parse_txt(file_path)

    def _parse_docx(self, file_path: str) -> Tuple[List[Dict], List[List[List[str]]], int]:
        if not DOCX_AVAILABLE:
            raise ImportError("请安装 python-docx")

        doc = Document(file_path)
        text_paragraphs = []
        tables = []

        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            if table_data:
                tables.append(table_data)

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if len(text) < 15:
                continue
            if any(kw in text for kw in self.skip_keywords) and len(text) < 50:
                continue
            has_keyword = any(kw in text for kw in self.content_keywords)
            is_table_title = "表" in text and re.search(r'\d+\.\d+', text) is not None
            if has_keyword or is_table_title:
                text_paragraphs.append({
                    "id": f"P_{i+1:04d}",
                    "content": text,
                    "page": 1,
                    "is_table_title": is_table_title
                })

        return text_paragraphs, tables, 1

    def _parse_pdf(self, file_path: str) -> Tuple[List[Dict], List[List[List[str]]], int]:
        if not PDF_AVAILABLE:
            raise ImportError("请安装 PyMuPDF")

        doc = fitz.open(file_path)
        text_paragraphs = []
        tables = []
        para_counter = 0

        for page_num, page in enumerate(doc):
            text = page.get_text()
            for para in text.split('\n\n'):
                para = para.strip()
                if len(para) < 20:
                    continue
                if any(kw in para for kw in self.content_keywords):
                    para_counter += 1
                    text_paragraphs.append({
                        "id": f"P_{page_num+1:04d}_{para_counter:03d}",
                        "content": para,
                        "page": page_num + 1,
                        "is_table_title": False
                    })

            page_tables = page.find_tables()
            for tab in page_tables:
                table_data = []
                for row in tab.extract():
                    row_data = [str(cell).strip() if cell else "" for cell in row]
                    table_data.append(row_data)
                if table_data:
                    tables.append(table_data)

        doc.close()
        return text_paragraphs, tables, len(doc)

    def _parse_txt(self, file_path: str) -> Tuple[List[Dict], List[List[List[str]]], int]:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        text_paragraphs = []
        tables = []

        for i, para in enumerate(content.split('\n\n')):
            para = para.strip()
            if len(para) > 20 and any(kw in para for kw in self.content_keywords):
                text_paragraphs.append({
                    "id": f"P_{i+1:04d}",
                    "content": para,
                    "page": 1,
                    "is_table_title": False
                })

        return text_paragraphs, tables, 1


# ============================================================================
# 表格处理器
# ============================================================================

class TableProcessor:
    """表格处理器"""

    @staticmethod
    def merge_adjacent_rows(table_data: List[List[str]]) -> List[List[str]]:
        """合并完全相同的相邻行，删除多余行"""
        if not table_data:
            return table_data

        result = []
        prev_row = None

        for row in table_data:
            if prev_row is None or row != prev_row:
                result.append(row)
                prev_row = row

        return result

    @staticmethod
    def merge_adjacent_cols(table_data: List[List[str]]) -> List[List[str]]:
        """合并完全相同的相邻列，删除多余列"""
        if not table_data:
            return table_data

        rows = len(table_data)
        cols = max(len(row) for row in table_data) if table_data else 0

        padded = []
        for row in table_data:
            while len(row) < cols:
                row.append("")
            padded.append(row)

        cols_to_delete = [False] * cols

        for j in range(cols - 1):
            is_same = True
            for i in range(rows):
                if padded[i][j] != padded[i][j + 1]:
                    is_same = False
                    break
            if is_same:
                cols_to_delete[j + 1] = True

        new_table = []
        for i in range(rows):
            new_row = []
            for j in range(cols):
                if not cols_to_delete[j]:
                    new_row.append(padded[i][j])
            new_table.append(new_row)

        return new_table

    @staticmethod
    def split_by_header(table_data: List[List[str]]) -> List[List[List[str]]]:
        """按相同表头切分表格 - 相似度超过2/3即可"""
        if not table_data or len(table_data) < 2:
            return [table_data]

        header_row = table_data[0]
        header_non_empty = [(j, header_row[j]) for j in range(len(header_row)) if header_row[j]]

        if not header_non_empty:
            return [table_data]

        split_indices = [0]
        for i in range(1, len(table_data)):
            current_row = table_data[i]

            match_count = 0
            for j, header_text in header_non_empty:
                if j < len(current_row) and current_row[j] == header_text:
                    match_count += 1

            similarity = match_count / len(header_non_empty) if header_non_empty else 0

            if similarity >= 2/3:
                split_indices.append(i)

        result = []
        for idx, start in enumerate(split_indices):
            end = split_indices[idx + 1] if idx + 1 < len(split_indices) else len(table_data)
            result.append(table_data[start:end])

        return result


# ============================================================================
# 大语言模型表格解析器
# ============================================================================

class LLMTableParser:
    """基于大语言模型的表格解析器"""

    def __init__(self):
        self.cache = {}  # 缓存相同结构表格的解析结果

    def _get_table_signature(self, table_data: List[List[str]]) -> str:
        """获取表格结构签名（用于判断是否为相同类型表格）"""
        if not table_data:
            return ""

        rows = len(table_data)
        cols = max(len(row) for row in table_data) if table_data else 0

        # 提取表头行
        header_row = table_data[0] if rows > 0 else []

        # 签名 = 行数范围 + 列数 + 表头内容哈希
        row_range = "small" if rows < 20 else "medium" if rows < 50 else "large"
        header_hash = hashlib.md5("|".join(header_row).encode()).hexdigest()[:8]

        return f"{row_range}_{cols}_{header_hash}"

    def _call_deepseek(self, table_data: List[List[str]]) -> Dict:
        """调用DeepSeek API解析表格 - 使用OpenAI SDK"""

        # 将表格转换为文本表示
        table_text = []
        for i, row in enumerate(table_data):
            row_text = " | ".join([str(cell) if cell else "" for cell in row])
            table_text.append(f"行{i}: {row_text}")
        table_str = "\n".join(table_text)

        prompt = f"""你是岩土工程表格解析专家。分析表格，将每个单元格互斥地划分为"主表"、"附表"或"备注"。

        【优先级】备注 > 主表 > 附表

        【判断标准】

        1. **备注**：说明性文字
           - 包括键值对格式（如"参数: 值"）
           - 或者以"建议"、"说明"、"注"、"备注"开头的文本
           - 大段文本（>20字符）或包含换行符\\n
           - 相邻列或行内容相同或类似（被切分的备注单元格或同一类型）

        2. **主表**：核心结构化数据
           - 包含岩土关键参数：孔号、孔深、土层编号、含水量、孔隙比、液限、塑限、压缩模量、黏聚力、内摩擦角等
           - 有明确的表头和数据行（数字/短文本）
           - 通常位于表格左侧/顶部，数据行数最多

        3. **附表**：仅在主表无法覆盖时才使用
           - 主表之外另起的结构化数据区域
           - 有独立的表头和数据行
           - **非必要不使用附表，优先将区域归入主表**

        【输出格式】
        {{
            "has_main_table": true/false,
            "main_table": {{
                "start_row": 0,
                "end_row": 行数-1,
                "start_col": 0,
                "end_col": 列数-1,
                "headers": [{{"col_index": 0, "start_col": 0, "end_col": 0, "header_name": "表头名"}}],
                "description": "描述"
            }},
            "has_sub_tables": true/false,
            "sub_tables": [
                {{
                    "start_row": 起始行,
                    "end_row": 结束行,
                    "start_col": 起始列,
                    "end_col": 结束列,
                    "headers": [{{"col_index": 0, "start_col": 0, "end_col": 0, "header_name": "表头名"}}],
                    "description": "描述"
                }}
            ],
            "has_remarks": true/false,
            "remarks": [
                {{
                    "start_row": 起始行,
                    "end_row": 结束行,
                    "start_col": 起始列,
                    "end_col": 结束列,
                    "content": "备注摘要"
                }}
            ]
        }}

        【规则】
        - 行号和列号从0开始
        - 每个表格必须有且只有一个主表
        - 附表是非必要的，仅在主表确实无法覆盖时才使用
        - 备注列若相邻内容相同则合并为同一备注区域

        表格数据：
        {table_str}

        只返回JSON。"""

        # 输出大模型输入
        print("\n[第一次优化 - 大模型输入]")
        print("-" * 80)
        print(f"表格大小: {len(table_data)}行 x {len(table_data[0]) if table_data else 0}列")
        print(f"\n原始表格数据（完整显示）:")
        for i, row in enumerate(table_data):
            # 显示每行的前10列，避免输出过长
            display_row = row[:10] if len(row) > 10 else row
            if len(row) > 10:
                print(f"  行{i}: {display_row}... (共{len(row)}列)")
            else:
                print(f"  行{i}: {display_row}")
        print(f"\n完整Prompt:")
        print(prompt)
        print(f"\nPrompt长度: {len(prompt)}字符")
        print("-" * 80)

        try:
            print(f"\n[第一次优化 - API调用] 发送请求到 {DEEPSEEK_BASE_URL}")
            print(f"[第一次优化 - API调用] 模型: deepseek-chat")
            print(f"[第一次优化 - API调用] temperature: 0.1, max_tokens: 2000")

            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个岩土工程表格解析专家，擅长识别表格结构。只返回JSON格式结果。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=2000,
                stream=False
            )

            content = response.choices[0].message.content

            # 输出大模型输出
            print("\n[第一次优化 - 大模型输出]")
            print("-" * 80)
            print(f"响应状态: 成功")
            print(f"响应长度: {len(content)}字符")
            print(f"使用tokens: {response.usage.total_tokens if hasattr(response, 'usage') else 'N/A'}")
            print(f"\n原始响应内容:")
            print(content)
            print("-" * 80)

            # 提取JSON
            json_match = re.search(r'\{[\s\S]*\}', content)
            if json_match:
                parsed = json.loads(json_match.group())

                # 输出解析结果
                print("\n[第一次优化 - 解析结果]")
                print("-" * 80)
                print(f"✓ has_main_table: {parsed.get('has_main_table')}")
                if parsed.get('has_main_table'):
                    main_table = parsed.get('main_table', {})
                    print(
                        f"  └─ main_table: 行{main_table.get('start_row')}-{main_table.get('end_row')}, 列{main_table.get('start_col')}-{main_table.get('end_col')}")
                    print(f"     description: {main_table.get('description')}")
                    # 输出headers信息
                    headers = main_table.get('headers', [])
                    if headers:
                        print(f"     表头信息 ({len(headers)}列):")
                        for header in headers:
                            print(
                                f"       列{header.get('col_index')}: [{header.get('start_col')}-{header.get('end_col')}] {header.get('header_name')}")
                    else:
                        print(f"     警告: 未返回headers信息")

                print(f"✓ has_remarks: {parsed.get('has_remarks')}")
                if parsed.get('has_remarks'):
                    remarks = parsed.get('remarks', [])
                    print(f"  └─ 共{len(remarks)}个备注区域:")
                    for i, remark in enumerate(remarks):
                        print(
                            f"     备注{i + 1}: 行{remark.get('start_row')}-{remark.get('end_row')}, 列{remark.get('start_col')}-{remark.get('end_col')}")
                        content_preview = remark.get('content', '')[:100]
                        print(
                            f"        content: {content_preview}{'...' if len(remark.get('content', '')) > 100 else ''}")

                print(f"✓ has_sub_tables: {parsed.get('has_sub_tables')}")
                if parsed.get('has_sub_tables'):
                    sub_tables = parsed.get('sub_tables', [])
                    print(f"  └─ 共{len(sub_tables)}个附表区域:")
                    for i, sub_tab in enumerate(sub_tables):
                        print(
                            f"     附表{i + 1}: 行{sub_tab.get('start_row')}-{sub_tab.get('end_row')}, 列{sub_tab.get('start_col')}-{sub_tab.get('end_col')}")
                        print(f"        description: {sub_tab.get('description')}")
                        # 输出headers信息
                        headers = sub_tab.get('headers', [])
                        if headers:
                            print(f"        表头信息 ({len(headers)}列):")
                            for header in headers:
                                print(
                                    f"          列{header.get('col_index')}: [{header.get('start_col')}-{header.get('end_col')}] {header.get('header_name')}")
                        else:
                            print(f"        警告: 未返回headers信息")

                print("=" * 80 + "\n")
                return parsed
            else:
                print("\n[第一次优化 - 错误]")
                print("-" * 80)
                print("✗ 无法从响应中提取JSON")
                print(f"响应内容前500字符: {content[:500]}")
                print("=" * 80 + "\n")
                return self._get_default_parse_result(table_data)

        except Exception as e:
            print("\n[第一次优化 - 异常]")
            print("-" * 80)
            print(f"✗ 异常类型: {type(e).__name__}")
            print(f"✗ 异常信息: {e}")
            import traceback
            print(f"\n完整堆栈跟踪:")
            traceback.print_exc()
            print("=" * 80 + "\n")
            return self._get_default_parse_result(table_data)

    def _get_default_parse_result(self, table_data: List[List[str]]) -> Dict:
        """默认解析结果（API失败时的降级方案）"""
        rows = len(table_data)
        cols = max(len(row) for row in table_data) if table_data else 0

        # 构建默认headers
        default_headers = []
        if rows > 0 and table_data[0]:
            for j in range(cols):
                header_name = table_data[0][j] if j < len(table_data[0]) else f"列{j}"
                default_headers.append({
                    "col_index": j,
                    "start_col": j,
                    "end_col": j,
                    "header_name": header_name
                })

        return {
            "has_main_table": True,
            "main_table": {
                "start_row": 0,
                "end_row": rows - 1,
                "start_col": 0,
                "end_col": cols - 1,
                "headers": default_headers,
                "description": "整个表格"
            },
            "has_remarks": False,
            "remarks": [],
            "has_sub_tables": False,
            "sub_tables": []
        }

    def parse(self, table_data: List[List[str]]) -> Dict:
        """解析表格，返回主表/备注/附表的位置信息"""
        if not table_data:
            return {
                "has_main_table": False,
                "main_table": None,
                "has_remarks": False,
                "remarks": [],
                "has_sub_tables": False,
                "sub_tables": []
            }

        # 获取表格签名
        signature = self._get_table_signature(table_data)

        # 检查缓存
        if signature in self.cache:
            print(f"  [缓存命中] 使用相同结构表格的解析结果")
            return self.cache[signature]

        # 调用大模型
        print(f"  [调用DeepSeek] 解析表格 {len(table_data)}行 x {len(table_data[0]) if table_data else 0}列")
        result = self._call_deepseek(table_data)

        # 缓存结果
        self.cache[signature] = result

        return result


# ============================================================================
# 表格提取器 - 根据LLM解析结果提取子表格
# ============================================================================

class TableExtractor:
    """根据LLM解析结果提取子表格"""

    @staticmethod
    def extract_sub_table(table_data: List[List[str]], region: Dict) -> List[List[str]]:
        """提取指定区域的子表格"""
        if not region:
            return []

        start_row = region.get("start_row", 0)
        end_row = region.get("end_row", len(table_data) - 1)
        start_col = region.get("start_col", 0)
        end_col = region.get("end_col", 0)

        start_row = max(0, min(start_row, len(table_data) - 1))
        end_row = max(start_row, min(end_row, len(table_data) - 1))
        start_col = max(0, start_col)
        end_col = max(start_col, end_col)

        result = []
        for i in range(start_row, end_row + 1):
            if i < len(table_data):
                row = table_data[i]
                if start_col < len(row):
                    result.append(row[start_col:min(end_col + 1, len(row))])
                else:
                    result.append([])

        return result

    @staticmethod
    def fix_remark_boundary(table_data: List[List[str]], parse_result: Dict) -> Dict:
        """修正备注边界：将包含备注内容的列从附表中移除"""

        # 获取备注区域
        remarks = parse_result.get("remarks", [])
        if not remarks:
            return parse_result

        # 获取附表区域
        sub_tables = parse_result.get("sub_tables", [])
        if not sub_tables:
            return parse_result

        # 对于每个附表，检查其起始列是否包含备注内容
        for sub_tab in sub_tables:
            sub_start_col = sub_tab.get("start_col", 0)
            sub_start_row = sub_tab.get("start_row", 0)

            # 检查附表起始列在表头行（sub_start_row）的内容
            if sub_start_row < len(table_data) and sub_start_col < len(table_data[sub_start_row]):
                cell_content = table_data[sub_start_row][sub_start_col]

                # 判断是否为备注内容（包含换行符、长度>50、包含"建议"等关键词）
                is_remark = False
                if cell_content:
                    if '\n' in cell_content:
                        is_remark = True
                    elif len(cell_content) > 50:
                        is_remark = True
                    elif any(keyword in cell_content for keyword in ['建议', '说明', '注', '备注']):
                        is_remark = True

                # 如果是备注内容，调整附表的起始列
                if is_remark:
                    # 找到第一个非备注内容的列
                    new_start_col = sub_start_col + 1
                    while new_start_col < len(table_data[sub_start_row]):
                        next_cell = table_data[sub_start_row][new_start_col]
                        is_next_remark = False
                        if next_cell:
                            if '\n' in next_cell:
                                is_next_remark = True
                            elif len(next_cell) > 50:
                                is_next_remark = True
                            elif any(keyword in next_cell for keyword in ['建议', '说明', '注', '备注']):
                                is_next_remark = True

                        if not is_next_remark:
                            break
                        new_start_col += 1

                    # 更新附表的起始列
                    sub_tab["start_col"] = new_start_col

                    # 更新附表的headers
                    old_headers = sub_tab.get("headers", [])
                    new_headers = []
                    for header in old_headers:
                        if header.get("col_index", 0) >= new_start_col:
                            # 调整col_index
                            header["col_index"] = header["col_index"] - (sub_start_col - new_start_col)
                            new_headers.append(header)
                    sub_tab["headers"] = new_headers

                    # 将备注列合并到备注区域
                    # 查找或创建包含这些列的备注区域
                    remark_found = False
                    for remark in remarks:
                        if remark.get("start_row", 0) <= sub_start_row <= remark.get("end_row", 0):
                            # 扩展备注区域的列范围
                            remark["start_col"] = min(remark.get("start_col", sub_start_col), sub_start_col)
                            remark["end_col"] = max(remark.get("end_col", sub_start_col), sub_start_col)
                            remark_found = True
                            break

                    if not remark_found:
                        # 创建新的备注区域
                        remarks.append({
                            "start_row": sub_start_row,
                            "end_row": sub_start_row,
                            "start_col": sub_start_col,
                            "end_col": sub_start_col,
                            "content": "备注内容"
                        })

        parse_result["remarks"] = remarks
        parse_result["sub_tables"] = sub_tables

        return parse_result
    @staticmethod
    def extract_sub_table(table_data: List[List[str]], region: Dict) -> List[List[str]]:
        """提取指定区域的子表格"""
        if not region:
            return []

        start_row = region.get("start_row", 0)
        end_row = region.get("end_row", len(table_data) - 1)
        start_col = region.get("start_col", 0)
        end_col = region.get("end_col", 0)

        # 确保范围有效
        start_row = max(0, min(start_row, len(table_data) - 1))
        end_row = max(start_row, min(end_row, len(table_data) - 1))
        start_col = max(0, start_col)
        end_col = max(start_col, end_col)

        result = []
        for i in range(start_row, end_row + 1):
            if i < len(table_data):
                row = table_data[i]
                if start_col < len(row):
                    result.append(row[start_col:min(end_col + 1, len(row))])
                else:
                    result.append([])

        return result


# ============================================================================
# 二次大模型解析器 - 处理重复行/列和合并表头
# ============================================================================

class SecondaryLLMParser:
    """二次大模型解析器 - 处理重复行/列和合并表头"""

    def __init__(self):
        pass  # 不使用缓存，每次独立调用

    def _call_deepseek_for_optimization(self, table_data: List[List[str]], table_name: str) -> Dict:
        """调用DeepSeek API优化表格 - 处理重复行/列和合并表头"""

        # 将表格转换为文本表示
        table_text = []
        for i, row in enumerate(table_data[:50]):  # 限制行数避免token过多
            row_text = " | ".join([str(cell) if cell else "" for cell in row])
            table_text.append(f"行{i}: {row_text}")
        table_str = "\n".join(table_text)

        prompt = f"""你是一个岩土工程表格优化专家。请分析以下表格（{table_name}），并进行优化处理。

优化规则：
1. **重复行处理**：识别并删除完全相同的重复行（除表头外），保留第一次出现
2. **重复列处理**：识别并删除完全相同的重复列（内容完全一致），保留第一次出现
3. **合并表头处理**：识别表头中是否包含复合信息，需要智能拆分成独立列：
   - 检测表头中是否包含多个概念（如"杆塔号塔型"、"孔号孔深"、"地层编号时代成因"等）
   - 将复合表头拆分为独立的列名,其中塔型，成因等信息在原数据中若只有一个，则很大可能对应多行数据，也就是对应多个杆塔号或者编号。
4.对于拆分表头后的所有表头需要进行对齐，如果没有单位，请忽略。如果有单位，以“数据（单位）”的格式进行重命名，如果不是非常确定，则保持原内容。
5.只优化上面四项内容，其他无关信息请忽略。保持原数据不变。

**请返回JSON格式结果，必须返回优化后的完整表格数据**。如果无需优化，返回原始表格。

优化后的表格数据示例：
{{
    "has_optimization": true,
    "optimization_type": []
}}

请返回JSON格式结果，**必须返回优化后的完整表格数据**：
{{
    "has_optimization": true/false,
    "optimization_type": ["duplicate_rows", "duplicate_cols", "merged_headers"],
    "optimized_table": [
        ["优化后的表头行1", "列2", "列3"],
        ["优化后的数据行1", "值2", "值3"]
    ],
    "explanation": "优化说明"
}}

原始表格数据：
{table_str}

注意：只返回JSON，optimized_table必须是优化后的完整表格数据（包含表头）。如果无需优化，返回原始表格。"""

        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个岩土工程表格优化专家。只返回JSON格式结果，必须包含optimized_table字段。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=4000,
                stream=False
            )

            content = response.choices[0].message.content
            print(f"  [二次解析响应] {content[:300]}...")

            json_match = re.search(r'\{[\s\S]*\}', content)
            if json_match:
                parsed = json.loads(json_match.group())
                return parsed
            else:
                return self._get_default_result(table_data)

        except Exception as e:
            print(f"  [二次解析异常] {e}")
            return self._get_default_result(table_data)

    def _get_default_result(self, table_data: List[List[str]]) -> Dict:
        return {
            "has_optimization": False,
            "optimization_type": [],
            "optimized_table": table_data,
            "explanation": "无需优化"
        }

    def optimize_table(self, table_data: List[List[str]], table_name: str) -> Dict:
        """优化表格 - 每次独立调用大模型"""
        if not table_data or len(table_data) < 1:
            return {
                "has_optimization": False,
                "optimized_table": table_data,
                "explanation": "表格为空"
            }

        print(f"  [二次解析] 调用DeepSeek优化{table_name}表格 {len(table_data)}行 x {len(table_data[0]) if table_data else 0}列")
        result = self._call_deepseek_for_optimization(table_data, table_name)

        return result


# ============================================================================
# 主解析系统 - 原有逻辑完全不变，只存储解析后的数据
# ============================================================================

class GeotechnicalParserSystem:
    def __init__(self):
        print("=" * 60)
        print("岩土报告智能解析系统")
        print("基于大语言模型的表格解析 - 主表/备注/附表识别")
        print("=" * 60)

        print("\n[1] 初始化文档解析器...")
        self.doc_parser = DocumentParser()

        print("\n[2] 初始化表格处理器...")
        self.table_processor = TableProcessor()

        print("\n[3] 初始化大语言模型解析器...")
        self.llm_parser = LLMTableParser()

        print("\n[4] 初始化二次大语言模型解析器...")
        self.secondary_parser = SecondaryLLMParser()

        print("\n[5] 初始化表格提取器...")
        self.table_extractor = TableExtractor()

        print("\n[6] 初始化表头对齐器...")
        self.header_aligner = get_header_aligner()

        print("\n✓ 系统初始化完成！")

    def parse_document(self, file_path: str) -> Dict:
        print(f"\n[解析文档] {file_path}")

        print("  步骤1: 提取文本段落和表格...")
        text_paragraphs, raw_tables, total_pages = self.doc_parser.parse_document(file_path)
        print(f"    文本段落: {len(text_paragraphs)}")
        print(f"    原始表格: {len(raw_tables)}")

        print("  步骤2: 按表头切分表格...")
        print("  步骤3: 对每个子表格合并相邻相同行列...")
        print("  步骤4: 大语言模型解析表格结构...")

        parsed_tables = []
        for idx, table_data in enumerate(raw_tables):
            print(f"    处理表格 T_{idx + 1:04d}: 原始 {len(table_data)}行")

            # 第一步：按表头切分
            split_tables = self.table_processor.split_by_header(table_data)
            print(f"      切分为 {len(split_tables)} 个独立表格")

            for sub_idx, sub_table in enumerate(split_tables):
                # 第二步：合并相邻相同行和列
                merged_rows = self.table_processor.merge_adjacent_rows(sub_table)
                merged_table = self.table_processor.merge_adjacent_cols(merged_rows)

                # 第三步：大语言模型解析表格结构
                parse_result = self.llm_parser.parse(merged_table)

                # 修正备注边界
                parse_result = self.table_extractor.fix_remark_boundary(merged_table, parse_result)

                # 第四步：根据解析结果提取子表格（不做表头对齐）
                main_table_data = []
                remark_data = []
                sub_table_data = []

                if parse_result.get("has_main_table"):
                    main_table_data = self.table_extractor.extract_sub_table(
                        merged_table, parse_result["main_table"]
                    )

                if parse_result.get("has_remarks"):
                    for remark in parse_result.get("remarks", []):
                        remark_content = self.table_extractor.extract_sub_table(merged_table, remark)
                        deduplicated_content = self._deduplicate_remark_content(remark_content)
                        remark_data.append({
                            "region": remark,
                            "data": deduplicated_content
                        })

                remark_data = self._deduplicate_remarks_list(remark_data)

                if parse_result.get("has_sub_tables"):
                    for sub_tab_idx, sub_tab in enumerate(parse_result.get("sub_tables", [])):
                        sub_content = self.table_extractor.extract_sub_table(merged_table, sub_tab)
                        sub_table_data.append({
                            "region": sub_tab,
                            "data": sub_content
                        })

                parsed_tables.append({
                    "table_id": f"T_{idx + 1:04d}_{sub_idx + 1}",
                    "original_data": merged_table,
                    "parse_result": parse_result,
                    "main_table": main_table_data,
                    "remarks": remark_data,
                    "sub_tables": sub_table_data,
                    "secondary_optimized": None,
                    "metadata": {
                        "rows": len(merged_table),
                        "cols": max(len(r) for r in merged_table) if merged_table else 0
                    }
                })
                print(
                    f"        子表格 {sub_idx + 1}: {len(merged_table)}行, 主表: {parse_result.get('has_main_table')}, 备注: {parse_result.get('has_remarks')}, 附表: {parse_result.get('has_sub_tables')}")

        return {
            "file_name": Path(file_path).name,
            "total_pages": total_pages,
            "text_paragraphs": text_paragraphs[:50],
            "text_count": len(text_paragraphs),
            "tables": parsed_tables
        }

    def _deduplicate_remark_content(self, remark_content: List[List[str]]) -> List[List[str]]:
        """对单个备注块的内容进行去重（去除内部重复的单元格和行）"""
        if not remark_content:
            return remark_content

        # 第一步：对每一行内的重复单元格进行去重
        deduplicated_rows = []
        for row in remark_content:
            # 去除行内重复的单元格（保留顺序，但去除相邻重复）
            unique_cells = []
            prev_cell = None
            for cell in row:
                cell_clean = cell.strip() if cell else ""
                if cell_clean and cell_clean != prev_cell:
                    unique_cells.append(cell)
                    prev_cell = cell_clean
                elif not cell_clean:
                    unique_cells.append(cell)
            deduplicated_rows.append(unique_cells)

        # 第二步：去除完全相同的行
        unique_rows = []
        seen_row_strs = set()
        for row in deduplicated_rows:
            # 将行转换为字符串用于比较（忽略空单元格）
            row_str = " | ".join([str(cell) for cell in row if cell and cell.strip()])
            if row_str and row_str not in seen_row_strs:
                seen_row_strs.add(row_str)
                unique_rows.append(row)
            elif not row_str and not any(cell for cell in row):
                # 空行跳过
                continue
            else:
                # 重复行，跳过
                continue

        # 第三步：合并多行备注内容（如果内容连续且相似）
        if len(unique_rows) > 1:
            merged_rows = []
            current_merged = []
            for row in unique_rows:
                row_text = " ".join([str(cell) for cell in row if cell and cell.strip()])
                if current_merged:
                    last_text = " ".join([str(cell) for cell in current_merged[-1] if cell and cell.strip()])
                    # 如果当前行是上一行的延续（不以数字开头或不是新条目）
                    if row_text and not re.match(r'^\d+[、.]', row_text):
                        # 合并到上一行
                        current_merged[-1] = [current_merged[-1][0] + " " + row_text] if current_merged[-1] else [
                            row_text]
                    else:
                        current_merged.append(row)
                else:
                    current_merged.append(row)
            unique_rows = current_merged

        return unique_rows

    def _deduplicate_remarks_list(self, remarks: List[Dict]) -> List[Dict]:
        """对备注列表进行全局去重（基于内容相似度）"""
        if not remarks:
            return remarks

        # 提取每个备注的内容文本
        remark_texts = []
        for remark in remarks:
            data = remark.get("data", [])
            # 将备注数据合并成文本，去重
            text_parts = []
            seen_phrases = set()
            for row in data:
                for cell in row:
                    if cell and cell.strip():
                        # 如果单元格内容过长，提取关键部分
                        cell_clean = cell.strip()
                        # 去除重复的段落（如相同内容重复出现）
                        if cell_clean not in seen_phrases:
                            seen_phrases.add(cell_clean)
                            text_parts.append(cell_clean)
            full_text = " ".join(text_parts)
            remark_texts.append(full_text)

        # 去重（内容相同或高度相似的只保留一个）
        unique_remarks = []
        unique_texts = []

        for i, remark in enumerate(remarks):
            text = remark_texts[i]
            # 如果内容为空，跳过
            if not text:
                continue

            # 检查是否与已有内容重复
            is_duplicate = False
            for j, existing_text in enumerate(unique_texts):
                # 如果当前文本是已有文本的子串，或者已有文本是当前文本的子串
                if text in existing_text or existing_text in text:
                    is_duplicate = True
                    # 保留内容更长的那个
                    if len(text) > len(existing_text):
                        unique_texts[j] = text
                        unique_remarks[j] = remark
                    break

                # 计算相似度
                min_len = min(len(text), len(existing_text))
                if min_len > 50:
                    common = self._longest_common_substring(text, existing_text)
                    if common / min_len > 0.6:  # 60%相似度认为是重复
                        is_duplicate = True
                        break

            if not is_duplicate:
                unique_texts.append(text)
                unique_remarks.append(remark)

        return unique_remarks

    def _longest_common_substring(self, s1: str, s2: str) -> int:
        """计算两个字符串的最长公共子串长度"""
        if not s1 or not s2:
            return 0

        len1, len2 = len(s1), len(s2)
        # 使用较短的字符串作为内层循环
        if len1 > len2:
            s1, s2 = s2, s1
            len1, len2 = len2, len1

        max_len = 0
        for i in range(len1):
            for j in range(len2):
                k = 0
                while i + k < len1 and j + k < len2 and s1[i + k] == s2[j + k]:
                    k += 1
                if k > max_len:
                    max_len = k
        return max_len


# ============================================================================
# Flask API
# ============================================================================

parser_system = None
tasks = {}


class ParseTask:
    def __init__(self, task_id, file_path, file_name):
        self.task_id = task_id
        self.file_path = file_path
        self.file_name = file_name
        self.status = "pending"
        self.progress = 0
        self.result = None
        self.error = None


def process_document(task):
    try:
        task.status = "processing"
        task.progress = 30

        if parser_system:
            result = parser_system.parse_document(task.file_path)
            task.result = result
            task.status = "completed"
            task.progress = 100

            # 解析完成后删除原始文件
            try:
                os.remove(task.file_path)
                print(f"已删除临时文件: {task.file_path}")
            except Exception as e:
                print(f"删除文件失败: {e}")
    except Exception as e:
        task.status = "failed"
        task.error = str(e)


def clean_for_json(obj):
    if obj is None:
        return None
    if isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, (list, tuple)):
        return [clean_for_json(item) for item in obj]
    if isinstance(obj, dict):
        return {str(k): clean_for_json(v) for k, v in obj.items()}
    return str(obj)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "没有上传文件"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "文件名为空"}), 400

    task_id = str(uuid.uuid4())[:8]
    saved_filename = f"{task_id}_{file.filename}"
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], saved_filename)
    file.save(file_path)

    task = ParseTask(task_id, file_path, file.filename)
    tasks[task_id] = task

    thread = threading.Thread(target=process_document, args=(task,))
    thread.daemon = True
    thread.start()

    return jsonify({"task_id": task_id, "status": "pending"})


@app.route('/api/task/<task_id>', methods=['GET'])
def get_task_status(task_id):
    if task_id not in tasks:
        return jsonify({"error": "任务不存在"}), 404
    task = tasks[task_id]
    return jsonify({
        "task_id": task.task_id,
        "status": task.status,
        "progress": task.progress,
        "error": task.error
    })


@app.route('/api/result/<task_id>', methods=['GET'])
def get_result(task_id):
    if task_id not in tasks:
        return jsonify({"error": "任务不存在"}), 404
    task = tasks[task_id]
    if task.status == "completed":
        return jsonify(clean_for_json(task.result))
    elif task.status == "failed":
        return jsonify({"error": task.error}), 500
    else:
        return jsonify({"status": task.status, "progress": task.progress}), 202


@app.route('/api/optimize/<task_id>/<int:table_index>', methods=['POST'])
def optimize_single_table(task_id, table_index):
    """对单个表格进行二次优化，优化后进行表头对齐"""
    if task_id not in tasks:
        return jsonify({"error": "任务不存在"}), 404

    task = tasks[task_id]
    if task.status != "completed" or not task.result:
        return jsonify({"error": "任务未完成"}), 400

    try:
        data = request.get_json()
        table_name = data.get('table_name', f'表格{table_index}')

        tables = task.result.get('tables', [])
        if table_index >= len(tables):
            return jsonify({"error": "表格索引无效"}), 400

        table = tables[table_index]
        aligner = get_header_aligner()

        main_table = table.get('main_table', [])
        sub_tables = table.get('sub_tables', [])

        optimization_result = {
            "main_table_optimized": None,
            "sub_tables_optimized": [],
            "header_alignment": {
                "main_table": None,
                "sub_tables": []
            }
        }

        # 优化主表
        if main_table and len(main_table) > 0:
            print(f"\n[二次优化] 开始优化主表: {table_name}")
            opt_result = parser_system.secondary_parser.optimize_table(main_table, f"{table_name}主表")
            optimized_main = opt_result.get("optimized_table", main_table)

            # 优化后进行表头对齐
            print(f"[表头对齐] 对优化后的主表进行对齐...")
            aligned_main, align_stats = aligner.align_table_headers(optimized_main)
            aligned_main_with_units = aligner.align_and_merge_units(optimized_main)

            optimization_result["main_table_optimized"] = {
                "optimized": aligned_main_with_units,
                "has_optimization": opt_result.get("has_optimization", False),
                "optimization_type": opt_result.get("optimization_type", []),
                "explanation": opt_result.get("explanation", ""),
                "header_alignment": {
                    "stats": align_stats,
                    "changes": [d for d in align_stats.get('details', []) if d['original'] != d['aligned']]
                }
            }

        # 优化附表
        for sub_idx, sub_tab in enumerate(sub_tables):
            sub_data = sub_tab.get('data', [])
            if sub_data and len(sub_data) > 0:
                print(f"\n[二次优化] 开始优化附表{sub_idx + 1}")
                opt_result = parser_system.secondary_parser.optimize_table(sub_data, f"{table_name}附表{sub_idx + 1}")
                optimized_sub = opt_result.get("optimized_table", sub_data)

                # 优化后进行表头对齐
                print(f"[表头对齐] 对优化后的附表{sub_idx + 1}进行对齐...")
                aligned_sub, sub_align_stats = aligner.align_table_headers(optimized_sub)
                aligned_sub_with_units = aligner.align_and_merge_units(optimized_sub)

                optimization_result["sub_tables_optimized"].append({
                    "sub_table_index": sub_idx,
                    "optimized": aligned_sub_with_units,
                    "has_optimization": opt_result.get("has_optimization", False),
                    "optimization_type": opt_result.get("optimization_type", []),
                    "explanation": opt_result.get("explanation", ""),
                    "header_alignment": {
                        "stats": sub_align_stats,
                        "changes": [d for d in sub_align_stats.get('details', []) if d['original'] != d['aligned']]
                    }
                })

        tables[table_index]["secondary_optimized"] = optimization_result

        # 输出对齐结果摘要
        print(f"\n[对齐结果摘要]")
        if optimization_result["main_table_optimized"]:
            stats = optimization_result["main_table_optimized"].get("header_alignment", {}).get("stats", {})
            print(
                f"  主表: 精确匹配={stats.get('matched', 0)}, 语义匹配={stats.get('semantic', 0)}, 保留={stats.get('original', 0)}")
        for sub_opt in optimization_result["sub_tables_optimized"]:
            stats = sub_opt.get("header_alignment", {}).get("stats", {})
            print(
                f"  附表{sub_opt['sub_table_index'] + 1}: 精确匹配={stats.get('matched', 0)}, 语义匹配={stats.get('semantic', 0)}, 保留={stats.get('original', 0)}")

        return jsonify({
            "success": True,
            "table_index": table_index,
            "optimization_result": optimization_result
        })

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/optimize_all/<task_id>', methods=['POST'])
def optimize_all_tables(task_id):
    """对所有表格进行二次优化，优化后进行表头对齐"""
    if task_id not in tasks:
        return jsonify({"error": "任务不存在"}), 404

    task = tasks[task_id]
    if task.status != "completed" or not task.result:
        return jsonify({"error": "任务未完成"}), 400

    try:
        tables = task.result.get('tables', [])
        aligner = get_header_aligner()
        optimized_count = 0

        for table_idx, table in enumerate(tables):
            table_name = table.get('table_id', f'表格{table_idx}')

            main_table = table.get('main_table', [])
            sub_tables = table.get('sub_tables', [])

            optimization_result = {
                "main_table_optimized": None,
                "sub_tables_optimized": [],
                "header_alignment": {
                    "main_table": None,
                    "sub_tables": []
                }
            }

            has_any_optimization = False

            # 优化主表
            if main_table and len(main_table) > 0:
                opt_result = parser_system.secondary_parser.optimize_table(main_table, f"{table_name}主表")
                optimized_main = opt_result.get("optimized_table", main_table)

                # 表头对齐
                aligned_main, align_stats = aligner.align_table_headers(optimized_main)
                aligned_main_with_units = aligner.align_and_merge_units(optimized_main)

                optimization_result["main_table_optimized"] = {
                    "optimized": aligned_main_with_units,
                    "has_optimization": opt_result.get("has_optimization", False),
                    "optimization_type": opt_result.get("optimization_type", []),
                    "explanation": opt_result.get("explanation", ""),
                    "header_alignment": {
                        "stats": align_stats,
                        "changes": [d for d in align_stats.get('details', []) if d['original'] != d['aligned']]
                    }
                }
                if opt_result.get("has_optimization"):
                    has_any_optimization = True

            # 优化附表
            for sub_idx, sub_tab in enumerate(sub_tables):
                sub_data = sub_tab.get('data', [])
                if sub_data and len(sub_data) > 0:
                    opt_result = parser_system.secondary_parser.optimize_table(sub_data,
                                                                               f"{table_name}附表{sub_idx + 1}")
                    optimized_sub = opt_result.get("optimized_table", sub_data)

                    # 表头对齐
                    aligned_sub, sub_align_stats = aligner.align_table_headers(optimized_sub)
                    aligned_sub_with_units = aligner.align_and_merge_units(optimized_sub)

                    optimization_result["sub_tables_optimized"].append({
                        "sub_table_index": sub_idx,
                        "optimized": aligned_sub_with_units,
                        "has_optimization": opt_result.get("has_optimization", False),
                        "optimization_type": opt_result.get("optimization_type", []),
                        "explanation": opt_result.get("explanation", ""),
                        "header_alignment": {
                            "stats": sub_align_stats,
                            "changes": [d for d in sub_align_stats.get('details', []) if d['original'] != d['aligned']]
                        }
                    })
                    if opt_result.get("has_optimization"):
                        has_any_optimization = True

            if has_any_optimization:
                optimized_count += 1

            tables[table_idx]["secondary_optimized"] = optimization_result

        return jsonify({
            "success": True,
            "optimized_count": optimized_count
        })

    except Exception as e:
        print(f"[错误] {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

def init_parser():
    global parser_system
    try:
        parser_system = GeotechnicalParserSystem()
        return True
    except Exception as e:
        print(f"初始化失败: {e}")
        return False


if __name__ == '__main__':
    print("=" * 60)
    print("岩土报告智能解析系统 - 后端服务")
    print("访问地址: http://localhost:5000")
    print("=" * 60)
    with app.app_context():
        init_parser()
    app.run(debug=True, host='0.0.0.0', port=5000)