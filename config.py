"""
配置文件 - 岩土报告智能解析系统
"""

import os

# Flask配置
FLASK_CONFIG = {
    'UPLOAD_FOLDER': './uploads',
    'MAX_CONTENT_LENGTH': 50 * 1024 * 1024,  # 50MB
    'SECRET_KEY': 'geotechnical-parser-secret-key',
    'DEBUG': True,
    'HOST': '0.0.0.0',
    'PORT': 5000
}

# DeepSeek API配置
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-db068eaf4d794e33a0d452203e4d8e9a")
DEEPSEEK_BASE_URL = "https://api.deepseek.com"

# 文档解析配置
DOC_PARSER_CONFIG = {
    'content_keywords': [
        '黏聚力', '内摩擦角', '压缩模量', '孔隙比', '含水量', '液限', '塑限',
        '颗粒级配', '粒径', '筛分', '剪切', '压缩', '载荷', '承载力'
    ],
    'skip_keywords': ['目录', 'Contents', '图', 'Fig', '表目录', '图目录', '参考文献'],
    'min_paragraph_length': 15,  # 最小段落长度
    'min_content_length': 20,     # 最小内容长度
    'skip_paragraph_length': 50   # 跳过段落长度阈值
}

# 表格处理配置
TABLE_PROCESSOR_CONFIG = {
    'header_similarity_threshold': 2/3,  # 表头相似度阈值
    'remark_keywords': ['建议', '说明', '注', '备注'],
    'remark_max_length': 50,              # 备注最大长度阈值
    'remark_min_length_for_merge': 50,    # 合并备注的最小长度
    'remark_similarity_threshold': 0.6    # 备注相似度阈值
}

# 大语言模型配置
LLM_CONFIG = {
    'model': "deepseek-chat",
    'temperature': 0.1,
    'max_tokens': 2000,
    'secondary_max_tokens': 4000,  # 二次解析的最大tokens
    'max_table_rows_for_prompt': 50  # 提示词中最大表格行数
}

# 缓存配置
CACHE_CONFIG = {
    'enabled': True,
    'max_size': 100  # 最大缓存数量
}

# 路径配置
PATH_CONFIG = {
    'uploads': './uploads',
    'results': './results'
}

# 表头对齐配置
HEADER_ALIGNMENT_CONFIG = {
    'min_similarity': 0.7,  # 最小相似度
    'use_semantic': True     # 使用语义匹配
}