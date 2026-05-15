"""
大语言模型（LLM）接口模块。

本模块封装了与通义千问（Qwen）API的调用。
您可以修改 `call_llm_api_async` 函数来接入您选择的任何大模型。
"""

import asyncio
import random
from typing import Optional
import aiohttp


async def call_llm_api_async(
    text: str,
    api_key: str,
    model: str = "qwen-max",
    additional_instructions: str = "",
    source_lang: str = "中文",
    target_lang: str = "英文",
    session: Optional[aiohttp.ClientSession] = None,
    extra_prompt: str = ""
) -> str:
    """
    异步调用大模型API进行翻译。

    注意：此函数当前配置为调用阿里云通义千问（Qwen）API。
    要使用其他模型，请修改此函数的实现。

    Args:
        text: 需要翻译的带标记文本。
        api_key: 大模型API的密钥。
        model: 使用的模型名称。
        additional_instructions: 用户自定义的翻译指令。
        source_lang: 源语言。
        target_lang: 目标语言。
        session: aiohttp会话（用于连接复用）。
        extra_prompt: 额外的提示词。

    Returns:
        str: 大模型返回的翻译结果。
    """
    # --- Qwen API 配置 ---
    url = "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    system_content = (
        "你的任务是保留所有<R0>、<R1>...<Rn>、<SEG0>、<SEG1>...<SEGn>、<HF0>、<HF1>...<HFn>等XML标签，"
        f"同时准确翻译所有{source_lang}文本为{target_lang}，非{source_lang}的文本不翻译，不添加任何与译文无关的内容，也不要添加除<Rn>、<SEGn>、<HFn>外的其他标签。"
        "【绝对禁止】不要使用任何Markdown格式！不要使用星号(*)、下划线(_)、反引号(`)、井号(#)等任何Markdown符号来格式化文本。"
        "输出必须是纯粹的、干净的文本，除了指定的XML标签外，不得包含任何特殊格式字符。"
        "XML标签后的文本之间是上下文关系，每个<Rn>标签后面对应一个待翻译的文本片段，<SEGn>和<HFn>标签对应一个段落，XML标签只有开头标签没有闭合标签。\n"
        "翻译时，请保持所有XML标签不变，仅翻译XML标签后的文本。\n"
        "以下为中译英例子：\n"
        "原文：<SEG0><R0>隧道在交通基础设施中展现出显著优势<R1>，尤其在<R2>山岭区域<R3>，其能够有效克服线性条件与高程限制等自然障碍，<R4>大幅缩短运输距离。\n"
        "翻译后：<SEG0><R0>Tunnels demonstrate significant advantages in transportation infrastructure<R1>, especially in<R2> mountainous regions<R3>, where they effectively overcome natural obstacles such as alignment constraints and elevation limitations,<R4> greatly shortening travel distances.\n"
        "注意：不同段落的xml标签数量不同，请保证<Rn>、<SEGn>、<HFn>等XML标签数量、编号和位置不变,不要减少或添加XML标签。以下为示例：\n"
        "原文：<SEG14><R0>专 <R1>业班级 <R2>:<R3>园林 <R4>2001\n"
        "错误添加标签的译文：<SEG14><R0>Major <R1>and <R2>Class <R3>:<R4>Landscape Gardening <R5>2001\n"
        "错误减少标签的译文：<SEG14><R0>Major and Class <R1>:<R2>Landscape Gardening <R3>2001\n"
        "正确的译文：<SEG14><R0>Major <R1>and Class <R2>:<R3>Landscape Gardening <R4>2001\n"
    )
    prompt = f"翻译要求：\n{additional_instructions.strip()}\n\n" if additional_instructions.strip() else ""
    user_content = f"{prompt}翻译{source_lang}文本为{target_lang}，以下是原文：\n{text}"
    if extra_prompt:
        user_content = f"{extra_prompt}\n\n{user_content}"
    payload = {
        "model": model,
        "input": {"messages": [
            {"role": "system", "content": system_content},
            {"role": "user", "content": user_content}
        ]},
        "parameters": {"result_format": "message"}
    }
    session = session or aiohttp.ClientSession()
    async with session.post(url, headers=headers, json=payload) as resp:
        if resp.status != 200:
            raise RuntimeError(f"API error: {resp.status} - {await resp.text()}")
        result = await resp.json()
        return result["output"]["choices"][0]["message"]["content"]
