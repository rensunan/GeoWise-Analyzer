import os

from openai import OpenAI
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-db068eaf4d794e33a0d452203e4d8e9a")
client = OpenAI(
 api_key=DEEPSEEK_API_KEY,
 base_url="https://api.deepseek.com/v1" # 聚合接口，一个 Key 调用多家模型
)

response = client.chat.completions.create(
 model="deepseek-v4-pro",
 messages=[
 {"role": "system", "content": "你是一个资深 Python 开发者"},
 {"role": "user", "content": "用 Python 实现一个带重试机制的 HTTP 客户端"}
 ],
 temperature=0.7,
 max_tokens=4096
)

print(response.choices[0].message.content)