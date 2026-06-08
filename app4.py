import os
from google import genai

# 1. 先设置 API 密钥！！！顺序绝对不能反
os.environ["GEMINI_API_KEY"] = "AIzaSyCu2ogsGcB_W8VYlrMveXX-lTESNx9vDF4"

# 2. 再创建客户端
client = genai.Client()

# 3. 使用正确的模型名
response = client.models.generate_content(
    model="gemini-2.0-flash",  # 正确可用模型
    contents="Explain how AI works in a few words"
)

print(response.text)