import os

from openai import OpenAI
from pathlib import Path
from docx import Document

import threading

from pathlib import Path
from docx import Document
from openai import OpenAI

# 初始化 API 客户端（需要确保 API Key 正确）
client = OpenAI(
    api_key="sk-gyqGihWcBDXjapcnDnOuDeLDDTgQ2oNmM5DUBmd5urZhwpl9",
    base_url="https://api.moonshot.cn/v1",
)


input_doc_path = "output_documents_new_combine/merged_document_1.docx"
output_dir = 'output_documents'
totalfile_num = 0

def zhipuparse(filepath):
    try:
        # 读取输入的Word文档
        input_doc = Document(filepath)

        # 将文档中的所有段落拼接成一个完整的字符串
        full_text = "\n".join([para.text for para in input_doc.paragraphs])

        # 执行API请求并获取处理结果
        messages = [
            {"role": "system", "content": "你是一位自媒体领域的大V"},
            {"role": "system", "content": full_text},
            {"role": "user",
             "content": "  请提取分析其中的金句10到15句，反共识观点10到15句，以及钩子语句。"}
        ]

        # 然后调用 chat-completion, 获取 Kimi 的回答
        completion = client.chat.completions.create(
            model="moonshot-v1-8k",
            messages=messages,
            temperature=0.3,
        )

        result = str(completion.choices[0].message)
        print(result)

    except Exception as e:
        print(f"处理文件 {filepath} 时发生错误: {e}")


if __name__ == '__main__':
    zhipuparse(input_doc_path)