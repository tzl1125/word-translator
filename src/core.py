"""
核心模块：定义数据结构和公共辅助函数。
"""

import re
from dataclasses import asdict, dataclass, field
from typing import Optional, Union, Tuple, List, Dict


# ---------- 全局常量 ----------
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
}


# ---------- 数据类定义 ----------
@dataclass
class RunInfo:
    """存储Word文档中单个run（文本片段）的所有信息。"""
    text: str
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None
    strike: Optional[bool] = None
    highlight: Optional[Union[int, str]] = None
    shading: Optional[Dict[str, str]] = None
    translated_text: str = ""
    color: Optional[Tuple[int, int, int]] = None
    is_field_code: bool = False
    original_xml: str = ""
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    underline_type: Optional[int] = None
    font_name_east_asia: Optional[str] = None
    font_name_cs: Optional[str] = None
    is_math: bool = False
    is_image: bool = False
    is_toc_field: bool = False
    is_page_break: bool = False

    def __eq__(self, other):
        if not isinstance(other, RunInfo):
            return False
        return (
                self.bold == other.bold and self.italic == other.italic and
                self.underline == other.underline and self.superscript == other.superscript and
                self.subscript == other.subscript and self.strike == other.strike and
                self.highlight == other.highlight and self.shading == other.shading and
                self.color == other.color and self.is_field_code == other.is_field_code and
                self.font_size == other.font_size and self.underline_type == other.underline_type
        )


@dataclass
class TextSegment:
    """存储文档正文中的一个段落。"""
    seg_idx: int
    full_text: str
    has_smartart_or_chart: bool = False
    runs_list: List[RunInfo] = field(default_factory=list)
    is_toc_paragraph: bool = False


@dataclass
class TableCellSegment:
    """存储表格单元格中的一个段落。"""
    table_idx: int
    row_idx: int
    cell_idx: int
    para_idx: int
    runs_list: List[RunInfo] = field(default_factory=list)


@dataclass
class HeaderFooterSegment:
    """存储页眉或页脚中的一个段落。"""
    section_idx: int
    is_header: bool
    para_idx: int
    full_text: str
    runs_list: List[RunInfo] = field(default_factory=list)


@dataclass
class ChartSegment:
    """存储图表中的可翻译文本。"""
    chart_idx: int
    element_type: str
    element_idx: int
    text: str
    file_path: str
    translated_text: str = ""


@dataclass
class SmartArtSegment:
    """存储SmartArt中的可翻译文本。"""
    smartart_idx: int
    element_idx: int
    text: str
    file_path: str
    translated_text: str = ""


# ---------- 公共辅助函数 ----------
def should_translate_text(text: str) -> bool:
    """
    判断文本是否需要翻译（非空且包含有效文字字符）。

    Args:
        text: 待判断的文本。

    Returns:
        bool: 如果需要翻译返回True，否则返回False。
    """
    if not text or not text.strip():
        return False
    stripped = re.sub(r'[\d\s.,;:!?\'"\\-_+=+*/\\|$$${}()<>~`@#$%^&]', '', text)
    return bool(stripped)


def create_marked_text(runs_list: List[Dict]) -> Tuple[str, List[int]]:
    """
    从 runs 列表生成带 <R0>, <R1>... 标记的文本，用于大模型翻译。

    Args:
        runs_list: RunInfo对象的字典列表。

    Returns:
        Tuple[str, List[int]]: 标记后的文本和对应的可翻译run索引列表。
    """
    parts = []
    indices = []
    rid = 0
    for idx, run in enumerate(runs_list):
        if run.get('is_field_code') or not run.get('text'):
            continue
        parts.append(f"<R{rid}>{run['text']}")
        indices.append(idx)
        rid += 1
    return "".join(parts), indices


def extract_translated_runs(translated_text: str, runs_list: List[Dict], translatable_indices: List[int]) -> None:
    """
    将大模型返回的带标记的翻译文本，解析并填充回原始的runs_list中。

    Args:
        translated_text: 大模型返回的已翻译文本。
        runs_list: 原始的RunInfo对象字典列表。
        translatable_indices: 可翻译run在runs_list中的索引。
    """
    for run in runs_list:
        run['translated_text'] = ''
    matches = re.findall(r'<R\d+>(.*?)(?=<R\d+>|$)', translated_text, re.DOTALL)
    expected = len(translatable_indices)
    if len(matches) == expected:
        for i, idx in enumerate(translatable_indices):
            runs_list[idx]['translated_text'] = matches[i]
    else:
        print(f"警告：翻译片段数量({len(matches)})与预期({expected})不一致。")
        if len(matches) < expected:
            for i, idx in enumerate(translatable_indices[:len(matches)]):
                runs_list[idx]['translated_text'] = matches[i]
        elif expected > 0:
            for i, idx in enumerate(translatable_indices[:-1]):
                runs_list[idx]['translated_text'] = matches[i]
            runs_list[translatable_indices[-1]]['translated_text'] = ''.join(matches[expected - 1:])
