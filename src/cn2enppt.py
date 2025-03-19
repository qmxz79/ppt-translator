from pptx import Presentation
from pptx.util import Pt
from openai import OpenAI
import re
from concurrent.futures import ThreadPoolExecutor

# 配置 DeepSeek API（已填入你的密钥 sk-22b8fd0b8d1c480cb30328623f4f847a）
deepseek_client = OpenAI(
    api_key="sk-22b8fd0b8d1c480cb30328623f4f847a",
    base_url="https://api.deepseek.com/v1"
)

translation_engine = "DeepSeek"

def set_translation_engine(engine):
    global translation_engine
    translation_engine = engine

def is_chinese(text):
    """检测文本是否包含中文字符"""
    return re.search(r'[\u4e00-\u9fff]', text)

def translate_text(text):
    """调用相应的翻译引擎翻译中文到英文"""
    try:
        if translation_engine == "DeepSeek":
            response = deepseek_client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个医学专业翻译引擎，**仅翻译中文到英文**，不要做任何其它解释，保留所有数字、符号、格式和英文内容。"},
                    {"role": "user", "content": text}
                ],
                temperature=0.1
            )
        elif translation_engine == "DomesticAPI":
            # 调用国内大模型API的翻译函数
            response = ...  # 替换为实际调用代码
        elif translation_engine == "LocalOllama":
            # 调用本地ollama部署的翻译函数
            response = ...  # 替换为实际调用代码
        else:
            raise ValueError("未知的翻译引擎")

        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"翻译失败: {e}")
        return text  # 失败时返回原文

def process_shape(shape):
    """处理单个形状"""
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if is_chinese(run.text):
                    translated = translate_text(run.text)
                    run.text = translated
                    run.font.size = Pt(20)
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if is_chinese(cell.text):
                    translated = translate_text(cell.text)
                    cell.text = translated
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(20)

def process_slide(slide):
    """处理单个幻灯片"""
    for shape in slide.shapes:
        process_shape(shape)

def translate_ppt(input_path, output_path, update_progress):
    """主处理函数"""
    prs = Presentation(input_path)
    total_slides = len(prs.slides)
    processed_slides = 0

    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(process_slide, slide) for slide in prs.slides]
        for future in futures:
            future.result()
            processed_slides += 1
            update_progress((processed_slides / total_slides) * 100)

    prs.save(output_path)
    print(f"翻译完成，结果已保存至: {output_path}")