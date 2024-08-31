import os

from openai import OpenAI
from pathlib import Path
from docx import Document

import threading

client = OpenAI(
    api_key="aecefa8c98656304dd5f26a3c6656bab.ne82Z5rap89I8H3d",
    base_url="https://open.bigmodel.cn/api/paas/v4/"
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

        completion = client.chat.completions.create(
            model="glm-4",
            messages=[
                {"role": "system", "content": "你是一位自媒体领域的大V"},
                {"role": "user",
                 "content": full_text + "  请提取分析其中的金句10到15句，反共识观点10到15句，以及钩子语句。"
                                        "如果系统检测到输入或生成内容可能包含不安全或敏感内容，请输出出来，帮助我主动删除"}

            ],
            top_p=0.7,
            temperature=0.9
        )
        print(completion.choices[0].message)

    except Exception as e:
        print(f"处理文件 {filepath} 时发生错误: {e}")


if __name__ == '__main__':
    zhipuparse(input_doc_path)