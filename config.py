"""
配置文件 - 岩土报告智能解析系统
"""

import os

# ---------- 从 .env 文件加载环境变量 ----------
def _load_dotenv():
    '''简陋的 .env 加载器，避免引入额外依赖'''
    env_path = os.path.join(os.path.dirname(__file__), '.env')
    if not os.path.isfile(env_path):
        return
    with open(env_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '=' not in line:
                continue
            k, v = line.split('=', 1)
            k = k.strip()
            v = v.strip().strip('"').strip("'")
            if k and k not in os.environ:
                os.environ[k] = v

_load_dotenv()

# Flask配置
FLASK_CONFIG = {
    'UPLOAD_FOLDER': './uploads',
    'MAX_CONTENT_LENGTH': 200 * 1024 * 1024,  # 200MB
    'SECRET_KEY': os.environ.get('SECRET_KEY', 'geotechnical-parser-secret-key'),
    'DEBUG': os.environ.get('DEBUG', 'True').lower() == 'true',
    'HOST': os.environ.get('HOST', '0.0.0.0'),
    'PORT': int(os.environ.get('PORT', 5000))
}

# DeepSeek API配置
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")
if not DEEPSEEK_API_KEY:
    raise RuntimeError("未设置 DEEPSEEK_API_KEY 环境变量，请在 .env 文件中配置")

DEEPSEEK_BASE_URL = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com")

# 文档解析配置
DOC_PARSER_CONFIG = {
    'content_keywords': [
        '凝聚力', '内摩擦角', '压缩模量', '孔隙比', '含水量', '液限', '塑限',
        '颗粒级配', '粒径', '筛分', '剪切', '压缩', '荷载', '承载力'
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
    'groundwater_keywords': [
        '地下水', '地下水位', '地下水埋深', '地下水变幅',
        '初见水位', '稳定水位', '水位埋深', '水位变幅'
    ],
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
