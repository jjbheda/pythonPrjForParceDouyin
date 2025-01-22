import re
from docx import Document
import re

def process_text(content):
    # 维护顺序并去重
    jingles = []
    contra_views = []
    hooks = []

    seen_jingles = set()
    seen_contra_views = set()
    seen_hooks = set()

    category_mapping = {
        "金句：": (jingles, seen_jingles),
        "反共识观点：": (contra_views, seen_contra_views),
        "钩子语句：": (hooks, seen_hooks)
    }

    current_category = None
    lines = content.split("\n")

    for line in lines:
        line = line.strip()
        if line in category_mapping:
            current_category = category_mapping[line]
        elif current_category:
            # 使用正则表达式以支持数字后多种格式的点
            match = re.match(r"^(\d+)\.\s*(.*)$", line)
            if match:
                sentence = match.group(2).strip('"')
                if sentence not in current_category[1]:
                    current_category[0].append(sentence)
                    current_category[1].add(sentence)

    return jingles, contra_views, hooks
def read_word_file(file_path):
    try:
        document = Document(file_path)
        content = []
        for para in document.paragraphs:
            content.append(para.text)
        return "\n".join(content)
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return ""

def write_word_file(jingles, contra_views, hooks, output_path):
    try:
        doc = Document()
        doc.add_paragraph('合并后的内容', style='Title')

        if jingles:
            doc.add_paragraph('金句:', style='Heading 2')
            for jingle in jingles:
                doc.add_paragraph(f'- "{jingle}"', style='Body Text')

        if contra_views:
            doc.add_paragraph('反共识观点:', style='Heading 2')
            for view in contra_views:
                doc.add_paragraph(f'- "{view}"', style='Body Text')

        if hooks:
            doc.add_paragraph('钩子语句:', style='Heading 2')
            for hook in hooks:
                doc.add_paragraph(f'- "{hook}"', style='Body Text')

        doc.save(output_path)
        print(f"Document saved successfully to {output_path}.")
    except Exception as e:
        print(f"An error occurred while writing the file {output_path}: {e}")


if __name__ == '__main__':

    print("Failed to read content. Please check the file path and permissions.")
