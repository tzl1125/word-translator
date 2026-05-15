"""
Word文档智能翻译工具 - 主入口

使用方法:
1. 安装依赖: pip install -r requirements.txt
2. 修改下方的配置参数。
3. 运行: python main.py
"""

from src.extractor import Extractor
from src.translator import Translator
from src.injector import Injector
import asyncio
import os


def main():
    # --- 配置区 (请在此处修改您的参数) ---
    INPUT_FILE = "source_files/your_document.docx"  # 待翻译的Word文档路径
    OUTPUT_DIR = "translated_files"                       # 输出目录
    API_KEY = "your-api-key-here"               # 您的大模型API密钥
    MODEL = "qwen-long"                         # 使用的模型
    SOURCE_LANG = "中文"                        # 源语言
    TARGET_LANG = "英文"                        # 目标语言
    MAX_CHUNK_SIZE = 2000                       # 单次翻译的最大字符数
    ADDITIONAL_INSTRUCTIONS = """
    请严格遵循以下翻译规则：
    1. 保持专业性和准确性，使用正式的商务语言。
    2. 保留原文中的数字、单位和专有名词不变。
    """

    # --- 执行流程 ---
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    name = os.path.splitext(os.path.basename(INPUT_FILE))[0]
    checkpoint_file = os.path.join(OUTPUT_DIR, f"{name}_checkpoint.json")
    output_file = os.path.join(OUTPUT_DIR, f"{name}_translated.docx")

    print("步骤 1/3: 正在提取文档内容...")
    Extractor(INPUT_FILE, checkpoint_file).extract()

    print("步骤 2/3: 正在翻译文本...")
    translator = Translator(
        checkpoint_file=checkpoint_file,
        api_key=API_KEY,
        model=MODEL,
        source_lang=SOURCE_LANG,
        target_lang=TARGET_LANG,
        max_chunk_size=MAX_CHUNK_SIZE,
        additional_instructions=ADDITIONAL_INSTRUCTIONS
    )
    asyncio.run(translator.translate())

    print("步骤 3/3: 正在写入翻译结果...")
    Injector(INPUT_FILE, checkpoint_file, output_file).inject()

    # 清理临时文件
    if os.path.exists(checkpoint_file):
        os.remove(checkpoint_file)

    print(f"✅ 翻译完成！结果已保存至: {output_file}")


if __name__ == "__main__":
    main()
