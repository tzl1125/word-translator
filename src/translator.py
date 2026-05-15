"""
翻译器模块：协调整个翻译过程，包括分块、调用LLM、重试机制等。
"""

import asyncio
import json
import random
import re
from typing import List, Dict, Optional
import aiohttp
from lxml import etree
from docx.oxml.shared import qn

from .core import create_marked_text, extract_translated_runs, NAMESPACES, should_translate_text
from .llm_interface import call_llm_api_async


class Translator:
    """
    翻译控制器。

    负责加载检查点文件，将文本分块，通过异步并发的方式调用大模型API，
    并处理目录、数学公式、表格、页眉页脚等特殊元素的翻译。
    """

    def __init__(self, checkpoint_file: str, api_key: str, model: str, source_lang: str, target_lang: str,
                 max_chunk_size: int, max_concurrent: int = 10, additional_instructions: str = ""):
        """
        初始化翻译器。

        Args:
            checkpoint_file: 提取器生成的检查点文件路径。
            api_key: LLM API密钥。
            model: 使用的LLM模型。
            source_lang: 源语言。
            target_lang: 目标语言。
            max_chunk_size: 单次API请求的最大文本长度。
            max_concurrent: 最大并发请求数。
            additional_instructions: 用户自定义翻译指令。
        """
        self.checkpoint_file = checkpoint_file
        self.api_key = api_key
        self.model = model
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.max_chunk_size = max_chunk_size
        self.semaphore = asyncio.Semaphore(max_concurrent)
        self.additional_instructions = additional_instructions
        with open(checkpoint_file, encoding="utf-8") as f:
            self.data = json.load(f)

    async def _translate_text(self, text: str, session: aiohttp.ClientSession, retries: int = 3, delay: int = 2,
                              extra_prompt: str = "") -> str:
        """带重试和退避机制的文本翻译。"""
        async with self.semaphore:
            for i in range(retries):
                try:
                    return await call_llm_api_async(text, self.api_key, self.model, self.additional_instructions,
                                                     self.source_lang, self.target_lang, session,
                                                     extra_prompt=extra_prompt)
                except RuntimeError as e:
                    if "429" in str(e) and i < retries - 1:
                        base_delay = delay * (2 ** i)
                        jitter = random.uniform(0, base_delay * 0.5)
                        total_delay = base_delay + jitter
                        await asyncio.sleep(total_delay)
                        continue
                    print(f"翻译错误: {e}")
                    return text
            return text

    def _chunk_segments(self, segments: List[Dict]) -> List[List[Dict]]:
        """将段落列表按最大长度分块。"""
        chunks = []
        cur = []
        size = 0
        for seg in segments:
            sz = len(seg['full_text'])
            if size + sz > self.max_chunk_size and cur:
                chunks.append(cur)
                cur, size = [], 0
            cur.append(seg)
            size += sz
        if cur: chunks.append(cur)
        return chunks

    async def _translate_single_segment_with_retry(self, seg: Dict, seg_type: str, expected_r_count: int,
                                                   session: aiohttp.ClientSession):
        """对单个段落进行翻译，并确保<R>标签数量正确。"""
        marked, indices = create_marked_text(seg['runs_list'])
        marked_text = f"<{seg_type}0>{marked}"
        extra_prompt = (
            "【特别强调】请务必确保输出中的<R>标签数量、顺序和编号与输入完全一致。"
            f"输入中有{expected_r_count}个<R>标签（从<R0>到<R{expected_r_count - 1}>），"
            "请勿增加、删除或更改任何<R>标签。每个<R>标签后的译文文本必须与原文对应，可以合并或拆分字词以保证<R>标签数量不变，"
            "输出格式必须为<SEG0><R0>...<R1>...的形式，不要添加额外的标记或格式。"
        )
        translate_content = marked
        for attempt in range(1, 4):
            translated = await self._translate_text(marked_text, session, extra_prompt=extra_prompt)
            match = re.search(f'<{seg_type}0>(.*)', translated, re.DOTALL)
            if not match:
                continue
            translate_content = match.group(1).strip()
            r_matches = re.findall(r'<R\d+>', translate_content)
            if len(r_matches) == expected_r_count:
                break
            await asyncio.sleep(0.3)  # 重试间隔
        extract_translated_runs(translate_content, seg['runs_list'], indices)

    async def _translate_segment_chunk(self, chunk: List[Dict], session: aiohttp.ClientSession,
                                       seg_type: str = "SEG") -> List[Dict]:
        """翻译一个文本块。"""
        marked_list = []
        translatable_map = {}
        expected_r_counts = []
        for i, seg in enumerate(chunk):
            marked, indices = create_marked_text(seg['runs_list'])
            key = self._seg_key(seg, seg_type)
            translatable_map[key] = indices
            expected_r_counts.append(len(indices))
            marked_list.append(f"<{seg_type}{i}>{marked}")
        combined = "\n\n".join(marked_list)
        translated = await self._translate_text(combined, session)
        parts = re.split(f'<{seg_type}\\d+>', translated)[1:]

        for i, seg in enumerate(chunk):
            key = self._seg_key(seg, seg_type)
            expected = expected_r_counts[i]
            if i < len(parts):
                part_content = parts[i].strip()
                r_matches = re.findall(r'<R\d+>', part_content)
                actual = len(r_matches)
                if actual == expected:
                    extract_translated_runs(part_content, seg['runs_list'], translatable_map[key])
                else:
                    await self._translate_single_segment_with_retry(seg, seg_type, expected, session)
            else:
                for r in seg['runs_list']: r['translated_text'] = r['text']
            seg['full_text'] = "".join(r.get('translated_text', r['text']) for r in seg['runs_list'])

        return chunk

    def _seg_key(self, seg: Dict, seg_type: str) -> str:
        """为段落生成唯一键。"""
        if seg_type == "SEG": return str(seg['seg_idx'])
        return f"{seg['section_idx']}_{'H' if seg['is_header'] else 'F'}_{seg['para_idx']}"

    async def _translate_table_cells(self, session: aiohttp.ClientSession):
        """翻译表格单元格。"""
        for seg in self.data["table_cell_segments"]:
            full = "".join(r['text'] for r in seg['runs_list'])
            if full.strip() and not re.search(r'[^\d.\s]', full):
                for r in seg['runs_list']: r['translated_text'] = r['text']
                continue
            marked, idxs = create_marked_text(seg['runs_list'])
            translated = await self._translate_text(marked, session)
            extract_translated_runs(translated, seg['runs_list'], idxs)

    async def _translate_headers_footers(self, session: aiohttp.ClientSession):
        """翻译页眉页脚。"""
        segs = self.data["header_footer_segments"]
        if not segs: return
        tasks = [self._translate_segment_chunk(chunk, session, "HF") for chunk in self._chunk_segments(segs)]
        await asyncio.gather(*tasks)

    async def _translate_special_runs(self, session: aiohttp.ClientSession, field_key: str, translate_func):
        """通用处理特殊run（如TOC, Math）。"""
        runs = []
        for seg_list in [self.data["text_segments"], self.data["table_cell_segments"],
                         self.data["header_footer_segments"]]:
            for seg in seg_list:
                for r in seg.get('runs_list', []):
                    if r.get(field_key):
                        runs.append(r)
        if not runs: return
        tasks = [translate_func(r, session) for r in runs]
        await asyncio.gather(*tasks)

    async def _translate_toc_run(self, run: Dict, session: aiohttp.ClientSession):
        """翻译目录项。"""
        if not run.get('original_xml'): return
        try:
            root = etree.fromstring(run['original_xml'])
            nodes_to_translate = []
            for t in root.findall('.//w:t', NAMESPACES):
                if t.text and not (
                        t.getparent() is not None and t.getparent().tag == qn('w:instrText')) and should_translate_text(
                    t.text):
                    nodes_to_translate.append(t)
            if not nodes_to_translate:
                return
            marked_parts = []
            for i, node in enumerate(nodes_to_translate):
                marked_parts.append(f"<R{i}>{node.text}")
            marked_text = "".join(marked_parts)
            translated = await self._translate_text(marked_text, session)

            matches = re.findall(r'<R\d+>(.*?)(?=<R\d+>|$)', translated, re.DOTALL)
            if len(matches) == len(nodes_to_translate):
                for i, node in enumerate(nodes_to_translate):
                    node.text = matches[i]
            else:
                print(f"警告：目录翻译结果片段数量不匹配，回退到单独翻译")
                for node in nodes_to_translate:
                    single_translated = await self._translate_text(f"<R0>{node.text}", session)
                    match = re.search(r'<R0>(.*?)(?=<R\d+>|$)', single_translated, re.DOTALL)
                    if match:
                        node.text = match.group(1)

            run['original_xml'] = etree.tostring(root, encoding='unicode')
            run['translated_text'] = ''.join(node.text for node in nodes_to_translate if node.text)
        except Exception as e:
            print(f"翻译目录错误: {e}")

    async def _translate_math_run(self, run: Dict, session: aiohttp.ClientSession):
        """翻译数学公式中的文本。"""
        if not run.get('original_xml'): return
        try:
            root = etree.fromstring(run['original_xml'])
            text_nodes = [n for n in root.findall('.//m:t', NAMESPACES) if n.text and should_translate_text(n.text)]
            if not text_nodes: return

            async def translate_node(node):
                translated = await self._translate_text(f"<R0>{node.text}", session)
                match = re.search(r'<R0>(.*?)(?=<R\d+>|$)', translated, re.DOTALL)
                if match: node.text = match.group(1)

            await asyncio.gather(*[translate_node(n) for n in text_nodes])
            run['original_xml'] = etree.tostring(root, encoding='unicode')
            run['translated_text'] = ''.join(n.text for n in text_nodes if n.text)
        except Exception as e:
            print(f"翻译数学公式错误: {e}")

    async def _translate_charts_and_smartart(self, session: aiohttp.ClientSession):
        """翻译图表和SmartArt。"""
        async def translate_one(seg):
            if seg['text'].strip():
                seg['translated_text'] = await self._translate_text(seg['text'], session)

        tasks = [translate_one(seg) for seg in self.data["chart_segments"] + self.data["smartart_segments"]]
        await asyncio.gather(*tasks)

    async def translate(self):
        """执行完整的翻译流程。"""
        async with aiohttp.ClientSession() as session:
            # 翻译目录段落和数学公式
            await self._translate_special_runs(session, 'is_toc_field', self._translate_toc_run)
            await self._translate_special_runs(session, 'is_math', self._translate_math_run)

            # 获取所有正文段落，过滤掉已经处理过的目录段落
            text_segments = self.data["text_segments"]
            non_toc_segments = [seg for seg in text_segments if not seg.get("is_toc_paragraph", False)]

            # 对非目录段落进行分块翻译
            chunks = self._chunk_segments(non_toc_segments)
            tasks = [self._translate_segment_chunk(chunk, session, "SEG") for chunk in chunks]
            await asyncio.gather(*tasks)

            # 翻译表格、页眉页脚、图表/SmartArt
            await self._translate_table_cells(session)
            await self._translate_headers_footers(session)
            await self._translate_charts_and_smartart(session)

            # 保存翻译结果到检查点文件
            with open(self.checkpoint_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
