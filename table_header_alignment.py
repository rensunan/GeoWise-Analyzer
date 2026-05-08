"""
岩土工程领域表头对齐模块
采用领域专用表头库进行语义对齐，支持同义词匹配和语义相似度匹配
"""

import re
import json
import os
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any
from collections import defaultdict
import warnings

warnings.filterwarnings('ignore')

# 尝试导入语义相似度模型
try:
    from sentence_transformers import SentenceTransformer
    import numpy as np
    SEMANTIC_AVAILABLE = True
except ImportError:
    SEMANTIC_AVAILABLE = False
    print("警告: sentence-transformers未安装，语义相似度匹配将不可用")


class GeotechnicalHeaderLibrary:
    """岩土工程领域表头库"""

    def __init__(self):
        """初始化表头库，包含同义词、常用符号和单位映射"""

        # 同义词映射库 {标准表头: [同义词列表]}
        self.synonyms = {
            "孔号": ["钻孔编号", "钻孔号", "孔编号", "勘探点编号", "钻孔ID", "孔ID", "编号", "点号"],
            "孔深": ["钻孔深度", "钻孔深", "深度", "孔深度", "钻探深度"],
            "土层编号": ["地层编号", "层号", "土层层号", "地层号", "层序号", "分层号"],
            "土层名称": ["地层名称", "土类名称", "岩土名称", "土层定名", "地层定名", "土名称"],
            "层底深度": ["层底标高", "层底埋深", "底面深度", "层底深度(m)", "层底标高(m)"],
            "层底标高": ["底面标高", "层底高程", "底标高", "高程"],
            "层厚": ["厚度", "土层厚度", "地层厚度", "层厚度(m)"],
            "颜色": ["土层颜色", "土的颜色", "岩土颜色", "色"],
            "状态": ["土的状态", "稠度状态", "状态描述", "可塑性"],
            "湿度": ["含水量状态", "湿度描述", "潮湿程度"],
            "密实度": ["密实程度", "相对密实度", "密实状态"],
            "压缩模量": ["Es", "压缩模量Es", "变形模量", "E_s", "压缩系数"],
            "黏聚力": ["c", "凝聚力", "粘聚力", "内聚力", "C值", "c值"],
            "内摩擦角": ["φ", "内摩擦角φ", "摩擦角", "Φ", "φ值"],
            "含水量": ["w", "天然含水量", "含水率", "ω", "W", "天然含水率"],
            "天然密度": ["ρ", "密度", "天然重度", "γ", "容重", "重度"],
            "孔隙比": ["e", "天然孔隙比", "e0", "孔隙率"],
            "液限": ["wL", "液性限度", "WL", "液限含水量"],
            "塑限": ["wP", "塑性限度", "WP", "塑限含水量"],
            "塑性指数": ["Ip", "塑性指数Ip", "I_p"],
            "液性指数": ["IL", "液性指数IL", "I_L"],
            "承载力特征值": ["fak", "承载力", "地基承载力", "承载力值", "f_ak", "容许承载力"],
            "标贯击数": ["N", "标准贯入击数", "SPT-N", "标贯值", "N值", "标准贯入"],
            "动探击数": ["Nd", "动力触探击数", "DPT", "动探值"],
            "取样位置": ["取样深度", "取土深度", "试样位置", "取样点"],
            "取样编号": ["样号", "土样编号", "样品编号", "试样编号"],
            "试验日期": ["日期", "测试日期", "检测日期"],
            "检测方法": ["试验方法", "测试方法", "检测标准"],
            "备注": ["说明", "注", "附注", "注释"],
            "极限侧阻力": ["qsk","侧阻力","极限端侧阻力"],
            "极限端阻力": ["qpk","端阻力"],
        }

        # 常用符号映射 {符号: 标准表头}
        self.symbols = {
            "Es": "压缩模量",
            "E_s": "压缩模量",
            "c": "黏聚力",
            "C": "黏聚力",
            "φ": "内摩擦角",
            "Φ": "内摩擦角",
            "w": "含水量",
            "ω": "含水量",
            "W": "含水量",
            "ρ": "天然密度",
            "γ": "天然密度",
            "e": "孔隙比",
            "e0": "孔隙比",
            "wL": "液限",
            "WL": "液限",
            "wP": "塑限",
            "WP": "塑限",
            "Ip": "塑性指数",
            "I_p": "塑性指数",
            "IL": "液性指数",
            "I_L": "液性指数",
            "fak": "承载力特征值",
            "f_ak": "承载力特征值",
            "N": "标贯击数",
            "Nd": "动探击数",
        }

        # 单位映射 {标准表头: [可能的单位列表]}
        self.units = {
            "孔深": ["m", "米", "M"],
            "层底深度": ["m", "米", "M"],
            "层底标高": ["m", "米", "M"],
            "层厚": ["m", "米", "M", "cm", "厘米"],
            "压缩模量": ["MPa", "Mpa", "kPa", "KPa", "mpa"],
            "黏聚力": ["kPa", "KPa", "MPa", "Mpa", "Pa"],
            "内摩擦角": ["°", "度", "deg"],
            "含水量": ["%", "百分数"],
            "天然密度": ["g/cm³", "kg/m³", "t/m³", "g/cm3", "kN/m³", "KN/m3"],
            "孔隙比": ["", "无单位"],
            "液限": ["%"],
            "塑限": ["%"],
            "塑性指数": ["%"],
            "液性指数": ["%"],
            "承载力特征值": ["kPa", "KPa", "MPa", "Mpa"],
        }

        # 地质参数正常范围 {参数名: (最小值, 最大值, 单位)}
        self.geo_ranges = {
            "孔深": (0.5, 200, "m"),
            "层底深度": (0, 200, "m"),
            "层底标高": (-100, 100, "m"),
            "层厚": (0.1, 50, "m"),
            "压缩模量": (0.5, 100, "MPa"),
            "黏聚力": (0, 500, "kPa"),
            "内摩擦角": (0, 60, "°"),
            "含水量": (0, 100, "%"),
            "天然密度": (1.0, 3.0, "g/cm³"),
            "孔隙比": (0.2, 2.0, ""),
            "液限": (10, 100, "%"),
            "塑限": (5, 80, "%"),
            "塑性指数": (0, 60, "%"),
            "液性指数": (-1, 2, ""),
            "承载力特征值": (30, 1000, "kPa"),
            "标贯击数": (0, 100, "击"),
            "动探击数": (0, 50, "击"),
        }

        # 单位转换因子 {标准表头: {原单位: 转换因子}}
        self.unit_converters = {
            "压缩模量": {
                "MPa": 1.0,
                "kPa": 0.001,
                "Pa": 0.000001
            },
            "黏聚力": {
                "kPa": 1.0,
                "MPa": 1000,
                "Pa": 0.001
            },
            "天然密度": {
                "g/cm³": 1.0,
                "kg/m³": 0.001,
                "t/m³": 1.0,
                "kN/m³": 0.102,  # kN/m³ 转 g/cm³ 约除以9.8
                "KN/m3": 0.102
            },
            "孔深": {
                "m": 1.0,
                "cm": 0.01,
                "mm": 0.001
            },
            "层厚": {
                "m": 1.0,
                "cm": 0.01
            }
        }

    def get_standard_header(self, raw_header: str) -> Tuple[str, float, str]:
        """
        获取标准表头
        返回: (标准表头, 匹配置信度, 匹配方式)
        匹配方式: exact, synonym, symbol, case_insensitive, partial, none
        """
        if not raw_header:
            return raw_header, 0.0, "none"

        raw_clean = raw_header.strip()

        # 清理换行符
        raw_clean = raw_clean.replace('\n', '').replace('\r', '')

        # 1. 精确匹配
        for standard, syn_list in self.synonyms.items():
            if raw_clean == standard:
                return standard, 1.0, "exact"
            if raw_clean in syn_list:
                return standard, 0.95, "synonym"

        # 2. 符号匹配
        if raw_clean in self.symbols:
            return self.symbols[raw_clean], 0.9, "symbol"

        # 3. 忽略括号内容的匹配（如"孔号(m)" -> "孔号"）
        without_unit = re.sub(r'[\(（].*?[\)）]', '', raw_clean).strip()
        if without_unit != raw_clean:
            sub_result = self.get_standard_header(without_unit)
            if sub_result[1] > 0.7:
                return sub_result

        # 4. 大小写归一化后的匹配
        lower_clean = raw_clean.lower()
        for standard, syn_list in self.synonyms.items():
            if lower_clean == standard.lower():
                return standard, 0.85, "case_insensitive"
            if lower_clean in [s.lower() for s in syn_list]:
                return standard, 0.85, "case_insensitive"

        # 5. 部分匹配（包含关键词）
        for standard, syn_list in self.synonyms.items():
            if standard in raw_clean:
                return standard, 0.7, "partial"
            for syn in syn_list:
                if syn in raw_clean:
                    return standard, 0.7, "partial"

        return raw_header, 0.0, "none"

    def extract_unit(self, header: str) -> Optional[str]:
        """
        从表头中提取单位
        例如: "孔深(m)" -> "m", "含水量(%)" -> "%"
        """
        if not header:
            return None

        # 匹配括号内的单位
        match = re.search(r'[\(（]([^\)）]+)[\)）]', header)
        if match:
            unit = match.group(1).strip()
            # 清理单位格式
            unit = unit.replace('(', '').replace(')', '').replace('（', '').replace('）', '')
            # 处理换行符
            unit = unit.replace('\n', '').replace('\r', '')
            return unit if unit else None

        # 匹配常见单位模式（在末尾）
        common_units = ['m', 'cm', 'mm', 'MPa', 'Mpa', 'kPa', 'KPa', 'Pa', '%', '°',
                        'g/cm³', 'g/cm3', 'kg/m³', 't/m³', 'kN/m³', 'KN/m3']
        for unit in common_units:
            if header.endswith(unit):
                return unit

        return None

    def validate_unit_with_data(self, header: str, unit: str, data_values: List[Any]) -> Tuple[bool, float]:
        """
        结合列数据验证单位是否合理
        返回: (是否合理, 合理数值比例)
        """
        if header not in self.geo_ranges:
            return True, 1.0

        min_val, max_val, expected_unit = self.geo_ranges[header]

        # 单位兼容性检查
        unit_compatible = False
        if expected_unit == "":
            unit_compatible = True
        elif unit in self.units.get(header, []):
            unit_compatible = True
        elif self._is_unit_convertible(header, unit, expected_unit):
            unit_compatible = True

        if not unit_compatible:
            return False, 0.0

        # 验证数据是否在合理范围内
        valid_count = 0
        total_count = 0

        for val in data_values:
            if val is None or val == "":
                continue

            total_count += 1
            try:
                # 清理数值字符串
                val_str = str(val).strip()
                # 移除可能的单位后缀
                for u in ['m', 'kPa', 'MPa', '%', '°', 'g/cm³', 'kN/m³']:
                    if val_str.endswith(u):
                        val_str = val_str[:-len(u)].strip()
                num_val = float(val_str)

                # 单位转换到标准单位
                converted_val = self._convert_unit(header, num_val, unit, expected_unit)
                if min_val <= converted_val <= max_val:
                    valid_count += 1
            except (ValueError, TypeError):
                continue

        if total_count == 0:
            return False, 0.0

        ratio = valid_count / total_count
        return ratio >= 0.5, ratio

    def _is_unit_convertible(self, header: str, from_unit: str, to_unit: str) -> bool:
        """检查两个单位之间是否可以转换"""
        if header not in self.unit_converters:
            return False
        converters = self.unit_converters[header]
        # 标准化单位名称
        from_std = self._normalize_unit(from_unit)
        to_std = self._normalize_unit(to_unit)
        return from_std in converters and to_std in converters

    def _normalize_unit(self, unit: str) -> str:
        """标准化单位名称"""
        if not unit:
            return ""
        unit_lower = unit.lower()
        # kN/m³ 转 KN/m3
        if unit_lower in ['kn/m³', 'kn/m3', 'kn/m^3']:
            return "kN/m³"
        if unit_lower in ['g/cm³', 'g/cm3', 'g/cm^3']:
            return "g/cm³"
        return unit

    def _convert_unit(self, header: str, value: float, from_unit: str, to_unit: str) -> float:
        """单位转换"""
        if header not in self.unit_converters:
            return value

        converters = self.unit_converters[header]
        from_std = self._normalize_unit(from_unit)
        to_std = self._normalize_unit(to_unit)

        if from_std in converters and to_std in converters:
            # 先转换到标准单位（因子1），再转换到目标单位
            value_in_standard = value * converters[from_std]
            return value_in_standard / converters[to_std]

        return value


class SemanticMatcher:
    """语义相似度匹配器"""

    def __init__(self, model_path: str = r"D:\sentence-embedding"):
        """
        初始化语义匹配器
        model_path: sentence-embedding模型路径
        """
        self.model = None
        self.model_path = model_path
        self.is_available = False

        if not SEMANTIC_AVAILABLE:
            print("警告: sentence-transformers未安装，语义相似度匹配不可用")
            return

        try:
            if os.path.exists(model_path):
                print(f"加载语义模型: {model_path}")
                self.model = SentenceTransformer(model_path)
                self.is_available = True
                print("语义模型加载成功")
            else:
                # 尝试下载默认模型
                print(f"模型路径不存在: {model_path}，尝试下载默认模型")
                self.model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
                self.is_available = True
                print("使用默认多语言模型")
        except Exception as e:
            print(f"语义模型加载失败: {e}")
            self.is_available = False

    def compute_similarity(self, text1: str, text2: str) -> float:
        """计算两个文本的语义相似度"""
        if not self.is_available or self.model is None:
            return 0.0

        try:
            # 清理文本
            t1 = text1.replace('\n', ' ').strip()
            t2 = text2.replace('\n', ' ').strip()
            embeddings = self.model.encode([t1, t2])
            from numpy.linalg import norm
            similarity = np.dot(embeddings[0], embeddings[1]) / (norm(embeddings[0]) * norm(embeddings[1]))
            return float(similarity)
        except Exception as e:
            print(f"相似度计算失败: {e}")
            return 0.0

    def find_best_match(self, query: str, candidates: List[str], threshold: float = 0.5) -> Tuple[Optional[str], float]:
        """
        在候选列表中查找最佳语义匹配
        返回: (最佳匹配, 相似度分数)
        """
        if not self.is_available or not candidates:
            return None, 0.0

        best_match = None
        best_score = 0.0

        for candidate in candidates:
            score = self.compute_similarity(query, candidate)
            if score > best_score and score >= threshold:
                best_score = score
                best_match = candidate

        return best_match, best_score


class TableHeaderAligner:
    """表头对齐器 - 整合关键词匹配和语义匹配"""

    def __init__(self):
        self.library = GeotechnicalHeaderLibrary()
        # self.semantic_matcher = SemanticMatcher()

        # 缓存对齐结果
        self.cache = {}

        print("表头对齐器初始化完成")

    def align_header(self, raw_header: str, column_data: List[Any] = None) -> Dict:
        """
        对齐单个表头
        返回: {
            "original": 原始表头,
            "aligned": 对齐后的标准表头,
            "unit": 单位(如果有),
            "confidence": 置信度,
            "match_type": 匹配方式,
            "validated": 是否通过验证
        }
        """
        # 检查缓存
        cache_key = f"{raw_header}_{len(column_data) if column_data else 0}"
        if cache_key in self.cache:
            return self.cache[cache_key]

        result = {
            "original": raw_header,
            "aligned": raw_header,
            "unit": None,
            "confidence": 0.0,
            "match_type": "none",
            "validated": False
        }

        if not raw_header:
            return result

        # 提取单位
        unit = self.library.extract_unit(raw_header)
        result["unit"] = unit

        # 1. 关键词匹配
        standard, confidence, match_type = self.library.get_standard_header(raw_header)

        if confidence > 0.7:
            result["aligned"] = standard
            result["confidence"] = confidence
            result["match_type"] = match_type
            result["validated"] = True

            # 单位验证（如果有列数据）
            if column_data and unit:
                is_valid, ratio = self.library.validate_unit_with_data(standard, unit, column_data)
                result["validated"] = is_valid
                if not is_valid:
                    result["warning"] = f"单位验证失败，仅{ratio:.0%}的数据在合理范围内"

            self.cache[cache_key] = result
            return result

        # 2. 语义相似度匹配（如果关键词匹配不成功）
        # if self.semantic_matcher.is_available:
        #     candidates = list(self.library.synonyms.keys())
        #     best_match, sim_score = self.semantic_matcher.find_best_match(raw_header, candidates, threshold=0.5)
        #
        #     if best_match and sim_score > 0.5:
        #         result["aligned"] = best_match
        #         result["confidence"] = sim_score
        #         result["match_type"] = "semantic"
        #         result["validated"] = True
        #
        #         # 单位验证
        #         if column_data and unit:
        #             is_valid, ratio = self.library.validate_unit_with_data(best_match, unit, column_data)
        #             result["validated"] = is_valid
        #             if not is_valid:
        #                 result["warning"] = f"单位验证失败，仅{ratio:.0%}的数据在合理范围内"
        #
        #         self.cache[cache_key] = result
        #         return result

        # 3. 不成功，使用原表头
        result["confidence"] = confidence
        result["match_type"] = "original"
        result["warning"] = "未匹配到标准表头，保留原值"

        self.cache[cache_key] = result
        return result

    def align_table_headers(self, table_data: List[List[str]]) -> Tuple[List[List[str]], Dict]:
        """
        对齐整个表格的表头
        返回: (对齐后的表格数据, 对齐结果统计)

        重要规则：如果两个不同的原始表头对齐后变成相同的标准表头，
        则两个都保持原样，不进行转换，避免列名重复
        """
        if not table_data or len(table_data) < 1:
            return table_data, {"total": 0, "matched": 0, "semantic": 0, "original": 0}

        # 获取表头行
        header_row = table_data[0]

        print(f"\n  [表头对齐] 开始处理 {len(header_row)} 个表头")

        # 提取每列的数据（用于单位验证）
        column_data = []
        for col_idx in range(len(header_row)):
            col_values = []
            for row in table_data[1:]:
                if col_idx < len(row) and row[col_idx]:
                    col_values.append(row[col_idx])
            column_data.append(col_values)

        # 第一遍：获取每个原始表头的对齐结果
        align_results = []
        for col_idx, raw_header in enumerate(header_row):
            col_data = column_data[col_idx] if col_idx < len(column_data) else []
            align_result = self.align_header(raw_header, col_data)
            align_results.append(align_result)

        # 第二遍：检测是否有重复的标准表头（排除空表头）
        standard_header_counts = {}
        for result in align_results:
            standard = result["aligned"]
            if standard and standard.strip():
                standard_header_counts[standard] = standard_header_counts.get(standard, 0) + 1

        # 标记哪些标准表头是重复的（出现次数 > 1）
        duplicate_standards = {std for std, count in standard_header_counts.items() if count > 1}

        # 第三遍：构建新表头，处理重复
        aligned_headers = []
        stats = {
            "total": len(header_row),
            "matched": 0,
            "semantic": 0,
            "original": 0,
            "duplicate_protected": 0,
            "unit_validated": 0,
            "unit_failed": 0,
            "details": []
        }

        for col_idx, result in enumerate(align_results):
            raw_header = result["original"]
            standard = result["aligned"]
            unit = result.get("unit")
            confidence = result.get("confidence", 0)
            match_type = result.get("match_type", "none")

            # 检查是否需要保护（重复且原始表头不同）
            is_duplicate = False
            if standard in duplicate_standards:
                for other_idx, other_result in enumerate(align_results):
                    if other_idx != col_idx and other_result["aligned"] == standard:
                        if other_result["original"] != raw_header:
                            is_duplicate = True
                            break

            if is_duplicate:
                # 重复的情况：保持原样
                final_header = raw_header
                stats["duplicate_protected"] += 1
                stats["original"] += 1
                print(f"    [去重保护] 列{col_idx}: 标准表头'{standard}'重复，保持原样: '{raw_header}'")
            else:
                # 不重复的情况：正常使用标准表头
                final_header = standard

                # 统计
                if match_type in ["exact", "synonym", "symbol", "case_insensitive", "partial"]:
                    stats["matched"] += 1
                    if raw_header != standard:
                        print(f"    [匹配] 列{col_idx}: '{raw_header}' → '{standard}' (类型: {match_type})")
                elif match_type == "semantic":
                    stats["semantic"] += 1
                    print(f"    [语义] 列{col_idx}: '{raw_header}' → '{standard}' (相似度: {confidence:.2f})")
                else:
                    stats["original"] += 1
                    if raw_header:
                        print(f"    [保留] 列{col_idx}: '{raw_header}' (未匹配到标准表头)")

            if result.get("validated", False):
                stats["unit_validated"] += 1
            elif result.get("warning") and "单位验证失败" in result.get("warning", ""):
                stats["unit_failed"] += 1

            stats["details"].append({
                **result,
                "final_header": final_header,
                "was_duplicate_protected": is_duplicate
            })

            aligned_headers.append(final_header)

        print(f"  [表头对齐] 完成: 匹配{stats['matched']}, 语义{stats['semantic']}, 保留{stats['original']}, 去重保护{stats['duplicate_protected']}")

        # 构建新的表格（替换表头行）
        aligned_table = [aligned_headers] + table_data[1:] if len(table_data) > 1 else [aligned_headers]

        return aligned_table, stats

    def align_and_merge_units(self, table_data: List[List[str]]) -> List[List[str]]:
        """
        对齐表头并处理单位：将单位合并到表头名称中
        格式：标准表头(单位)

        重要规则：如果两个不同的原始表头对齐后变成相同的标准表头，
        则两个都保持原样，不进行转换，避免列名重复
        """
        if not table_data or len(table_data) < 1:
            return table_data

        header_row = table_data[0]

        # 提取每列的数据
        column_data = []
        for col_idx in range(len(header_row)):
            col_values = []
            for row in table_data[1:]:
                if col_idx < len(row) and row[col_idx]:
                    col_values.append(row[col_idx])
            column_data.append(col_values)

        # 第一遍：获取每个原始表头的对齐结果
        align_results = []
        for col_idx, raw_header in enumerate(header_row):
            col_data = column_data[col_idx] if col_idx < len(column_data) else []
            align_result = self.align_header(raw_header, col_data)
            align_results.append(align_result)

        # 第二遍：检测是否有重复的标准表头（排除空表头）
        standard_header_counts = {}
        for result in align_results:
            standard = result["aligned"]
            if standard and standard.strip():
                standard_header_counts[standard] = standard_header_counts.get(standard, 0) + 1

        # 标记哪些标准表头是重复的（出现次数 > 1）
        duplicate_standards = {std for std, count in standard_header_counts.items() if count > 1}

        # 第三遍：构建新表头，重复的标准表头保持原样
        new_headers = []
        for col_idx, result in enumerate(align_results):
            raw_header = result["original"]
            standard = result["aligned"]
            unit = result.get("unit")

            # 检查这个标准表头是否重复（且原始表头不同）
            is_duplicate = False
            if standard in duplicate_standards:
                for other_idx, other_result in enumerate(align_results):
                    if other_idx != col_idx and other_result["aligned"] == standard:
                        if other_result["original"] != raw_header:
                            is_duplicate = True
                            break

            if is_duplicate:
                # 重复的情况：保持原样
                new_header = raw_header
                print(f"    [去重] 列{col_idx}: 标准表头'{standard}'重复，保持原样: '{raw_header}'")
            else:
                # 不重复的情况：正常处理
                if unit and unit.strip():
                    # 检查单位是否已经是标准格式
                    if f"({unit})" not in standard and f"（{unit}）" not in standard:
                        new_header = f"{standard}({unit})"
                    else:
                        new_header = standard
                else:
                    new_header = standard

            new_headers.append(new_header)

        # 返回新表格
        return [new_headers] + table_data[1:] if len(table_data) > 1 else [new_headers]


# 全局实例
_header_aligner = None


def get_header_aligner() -> TableHeaderAligner:
    """获取表头对齐器单例"""
    global _header_aligner
    if _header_aligner is None:
        _header_aligner = TableHeaderAligner()
    return _header_aligner


def align_table_headers(table_data: List[List[str]]) -> Tuple[List[List[str]], Dict]:
    """便捷函数：对齐表格表头"""
    aligner = get_header_aligner()
    return aligner.align_table_headers(table_data)


def align_and_merge_units(table_data: List[List[str]]) -> List[List[str]]:
    """便捷函数：对齐表头并合并单位"""
    aligner = get_header_aligner()
    return aligner.align_and_merge_units(table_data)