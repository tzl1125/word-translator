"""
提取器模块：负责从Word文档中提取所有文本、格式和复杂元素。
"""

import json
import os
import zipfile
from typing import List, Dict, Tuple, Optional
from lxml import etree
from docx import Document
from docx.oxml.shared import qn
from docx.text.paragraph import Paragraph

from .core import (
    NAMESPACES, RunInfo, TextSegment, TableCellSegment,
    HeaderFooterSegment, ChartSegment, SmartArtSegment,
    should_translate_text
)


class Extractor:
    """
    Word文档内容提取器。

    负责遍历整个.docx文件，提取正文、表格、页眉页脚、图表、SmartArt、
    数学公式、EndNote引用等所有元素，并将其转换为内部数据结构，
    同时保存到检查点文件中。
    """

    def __init__(self, input_file: str, checkpoint_file: str):
        """
        初始化提取器。

        Args:
            input_file: 输入的.docx文件路径。
            checkpoint_file: 用于保存提取结果的检查点文件路径。
        """
        self.input_file = input_file
        self.checkpoint_file = checkpoint_file
        self.doc = Document(input_file)
        self.text_segments: List[TextSegment] = []
        self.table_cell_segments: List[TableCellSegment] = []
        self.header_footer_segments: List[HeaderFooterSegment] = []
        self.chart_segments: List[ChartSegment] = []
        self.smartart_segments: List[SmartArtSegment] = []

    def _get_font_info(self, rPr_elem) -> dict:
        """从XML元素中提取字体信息。"""
        info = {'name': None, 'name_east_asia': None, 'name_cs': None, 'size': None, 'underline_type': None,
                'strike': None}
        if rPr_elem is None:
            return info
        rFonts = rPr_elem.find(qn('w:rFonts'))
        if rFonts is not None:
            info.update({'name': rFonts.get(qn('w:ascii')), 'name_east_asia': rFonts.get(qn('w:eastAsia')),
                         'name_cs': rFonts.get(qn('w:cs'))})
        sz = rPr_elem.find(qn('w:sz'))
        if sz is not None:
            try:
                info['size'] = int(sz.get(qn('w:val'))) / 2
            except:
                pass
        u = rPr_elem.find(qn('w:u'))
        if u is not None and (val := u.get(qn('w:val'))) and val.isdigit():
            info['underline_type'] = int(val)
        strike = rPr_elem.find(qn('w:strike'))
        d_strike = rPr_elem.find(qn('w:dstrike'))
        if strike is not None:
            val = strike.get(qn('w:val'))
            info['strike'] = val is None or val in ('true', '1')
        elif d_strike is not None:
            val = d_strike.get(qn('w:val'))
            info['strike'] = val is None or val in ('true', '1')
        return info

    def _create_run_info(self, text: str, run, font_info: dict, **kwargs) -> RunInfo:
        """根据Python-docx的run对象创建RunInfo数据类。"""
        is_math = kwargs.get('is_math', False)
        is_header_footer = kwargs.get('is_header_footer', False)
        is_field_code = kwargs.get('is_field_code', False)
        original_xml = kwargs.get('original_xml', "")
        is_space_only = kwargs.get('is_space_only', False)
        is_image = kwargs.get('is_image', False)

        highlight = color = shading = None
        if not is_math and not is_image:
            hl = run._element.get_or_add_rPr().find(qn('w:highlight'))
            if hl is not None and hl.get(qn('w:val')) != 'none':
                highlight = run.font.highlight_color
            if run.font.color and run.font.color.rgb:
                color = tuple(run.font.color.rgb)
            shd = run._element.get_or_add_rPr().find(qn('w:shd'))
            if shd is not None:
                shading = {'val': shd.get(qn('w:val')), 'color': shd.get(qn('w:color')), 'fill': shd.get(qn('w:fill'))}

        if is_header_footer and text.strip() == "" and text != "":
            is_field_code = True

        return RunInfo(
            text=text, bold=run.bold if not is_math and not is_image else False,
            italic=run.italic if not is_math and not is_image else False,
            underline=run.underline if not is_math and not is_image else False,
            superscript=run.font.superscript if not is_math and not is_image else False,
            subscript=run.font.subscript if not is_math and not is_image else False,
            strike=font_info.get('strike'), highlight=highlight, color=color, shading=shading,
            is_field_code=is_field_code or is_space_only or is_math or is_image,
            original_xml=original_xml if is_field_code or is_space_only or is_math or is_image else "",
            font_name=font_info.get('name'), font_size=font_info.get('size'),
            underline_type=font_info.get('underline_type'), font_name_east_asia=font_info.get('name_east_asia'),
            font_name_cs=font_info.get('name_cs'), is_math=is_math, is_image=is_image,
            is_toc_field=kwargs.get('is_toc_field', False)
        )

    def _extract_runs(self, para: Paragraph, is_header_footer: bool = False, is_toc_paragraph: bool = False) -> List[
        RunInfo]:
        """提取段落中的所有run信息。"""
        if is_toc_paragraph:
            original_xml = etree.tostring(para._element, encoding='unicode')
            toc_texts = []
            for hyperlink in para._element.findall('.//w:hyperlink', NAMESPACES):
                for t_elem in hyperlink.findall('.//w:t', NAMESPACES):
                    if t_elem.text:
                        toc_texts.append(t_elem.text)
            return [RunInfo(text=''.join(toc_texts), is_field_code=True, original_xml=original_xml, is_toc_field=True)]

        runs_list = []
        for child in para._element:
            tag = child.tag
            if tag.endswith('pPr'): continue
            if tag.endswith('br') and child.get(qn('w:type')) == 'page':
                runs_list.append(RunInfo(text="", is_field_code=True, is_page_break=True,
                                         original_xml=etree.tostring(child, encoding='unicode')))
                continue
            if tag.endswith('pb'):
                runs_list.append(RunInfo(text="", is_field_code=True, is_page_break=True,
                                         original_xml=etree.tostring(child, encoding='unicode')))
                continue

            if tag.endswith('bookmarkStart') or tag.endswith('bookmarkEnd'):
                runs_list.append(
                    RunInfo(text="", is_field_code=True, original_xml=etree.tostring(child, encoding='unicode')))
            elif tag.endswith('r'):
                br = child.find('.//w:br', NAMESPACES)
                if br is not None and br.get(qn('w:type')) == 'page':
                    runs_list.append(RunInfo(text="", is_field_code=True, is_page_break=True,
                                             original_xml=etree.tostring(child, encoding='unicode')))
                    continue

                run = next((r for r in para.runs if r._element == child), None)
                if not run:
                    continue
                is_image = child.find('.//w:drawing', NAMESPACES) is not None or child.find('.//w:pict',
                                                                                            NAMESPACES) is not None
                if is_image:
                    runs_list.append(self._create_run_info("", run, {}, is_field_code=True,
                                                           original_xml=etree.tostring(child, encoding='unicode'),
                                                           is_header_footer=is_header_footer, is_image=True))
                    continue
                is_field_code = self._check_is_field_code(child)
                original_xml = etree.tostring(child, encoding='unicode') if is_field_code else ""
                is_space_only = run.text.strip() == "" and run.text != "" and len(run.text) > 2
                if is_space_only and not original_xml:
                    original_xml = etree.tostring(child, encoding='unicode')
                font_info = self._get_font_info(child.get_or_add_rPr())
                run_info = self._create_run_info(run.text, run, font_info, is_field_code=is_field_code,
                                                 original_xml=original_xml, is_header_footer=is_header_footer,
                                                 is_space_only=is_space_only, is_image=False)
                if is_field_code or is_space_only:
                    runs_list.append(run_info)
                elif runs_list and run_info == runs_list[-1]:
                    runs_list[-1].text += run.text
                else:
                    runs_list.append(run_info)
            elif tag.endswith('oMath') or tag.endswith('oMathPara'):
                for math_elem in (child.findall('.//m:oMath', NAMESPACES) if tag.endswith('oMathPara') else [child]):
                    formula_xml = etree.tostring(math_elem, encoding='unicode')
                    math_text = "".join(
                        [t.text for t in math_elem.findall('.//m:t', NAMESPACES) if t.text]) or "[MATH_FORMULA]"
                    runs_list.append(
                        RunInfo(text=math_text, is_field_code=True, original_xml=formula_xml, is_math=True))
        return runs_list

    def _check_is_field_code(self, run_elem) -> bool:
        """检查一个run是否是域代码（如页码、引用等）。"""
        if run_elem.find('.//w:fldChar', NAMESPACES) is not None: return True
        instr = run_elem.find('.//w:instrText', NAMESPACES)
        if instr is not None and instr.text:
            if any(kw in instr.text for kw in
                   ['ADDIN', 'TOC', 'PAGE', 'NUMPAGES', 'DATE', 'TIME', 'REF', 'HYPERLINK']): return True
        if run_elem.find('.//w:bookmarkStart', NAMESPACES) is not None or run_elem.find('.//w:bookmarkEnd',
                                                                                        NAMESPACES) is not None: return True
        return False

    def extract(self):
        """执行完整的文档提取流程。"""
        # 收集所有段落（包括 sdt 内的目录）
        all_paragraphs = []  # (Paragraph, is_toc)
        for child in self.doc.element.body:
            if child.tag == qn('w:p'):
                all_paragraphs.append((Paragraph(child, self.doc), False))
            elif child.tag == qn('w:sdt'):
                is_toc = any(
                    instr.text and 'TOC' in instr.text for instr in child.findall('.//w:instrText', NAMESPACES))
                if not is_toc:
                    doc_part = child.find('.//w:docPartGallery', NAMESPACES)
                    if doc_part is not None and doc_part.get(qn('w:val'), '').lower() == 'table of contents':
                        is_toc = True
                sdt_content = child.find(qn('w:sdtContent'))
                if sdt_content is not None:
                    for p_elem in sdt_content.findall(qn('w:p')):
                        all_paragraphs.append((Paragraph(p_elem, self.doc), is_toc))

        for seg_idx, (para, is_toc) in enumerate(all_paragraphs):
            full_text = para.text.strip()
            if para.style and para.style.name:
                if any(kw in para.style.name.lower() for kw in ['endnote', 'bibliography', 'reference', 'references']):
                    continue
            if full_text or is_toc or self._has_smartart_or_chart(para):
                seg = TextSegment(seg_idx, full_text, self._has_smartart_or_chart(para), is_toc_paragraph=is_toc)
                seg.runs_list = self._extract_runs(para, is_toc_paragraph=is_toc)
                self.text_segments.append(seg)

        # 表格
        for ti, tbl in enumerate(self.doc.tables):
            for ri, row in enumerate(tbl.rows):
                for ci, cell in enumerate(row.cells):
                    for pi, para in enumerate(cell.paragraphs):
                        if para.text.strip() or self._has_smartart_or_chart(para):
                            seg = TableCellSegment(ti, ri, ci, pi)
                            seg.runs_list = self._extract_runs(para)
                            self.table_cell_segments.append(seg)

        # 页眉页脚
        for si, sec in enumerate(self.doc.sections):
            for is_header, elem in [(True, sec.header), (False, sec.footer)]:
                if not elem: continue
                for pi, p_elem in enumerate(elem._element.iter(qn('w:p'))):
                    para = Paragraph(p_elem, elem)
                    runs = self._extract_runs(para, is_header_footer=True)
                    if runs:
                        seg = HeaderFooterSegment(si, is_header, pi,
                                                  "".join(r.text for r in runs if not r.is_field_code))
                        seg.runs_list = runs
                        self.header_footer_segments.append(seg)

        # 图表和 SmartArt
        try:
            with zipfile.ZipFile(self.input_file) as z:
                for f in z.namelist():
                    if 'chart' in f.lower() and f.endswith('.xml'):
                        root = etree.fromstring(z.read(f))
                        for ti, title in enumerate(root.findall('.//c:title', NAMESPACES)):
                            for t_elem in title.findall('.//a:t', NAMESPACES):
                                if t_elem.text and t_elem.text.strip():
                                    self.chart_segments.append(
                                        ChartSegment(len(self.chart_segments), "title", ti, t_elem.text.strip(), f))
                    elif 'diagram' in f.lower() and f.endswith('.xml'):
                        root = etree.fromstring(z.read(f))
                        for ei, t_elem in enumerate(root.findall('.//a:t', NAMESPACES)):
                            if t_elem.text and t_elem.text.strip():
                                self.smartart_segments.append(
                                    SmartArtSegment(len(self.smartart_segments), ei, t_elem.text.strip(), f))
        except Exception as e:
            print(f"访问ZIP内容错误: {e}")

        self._save_checkpoint()

    def _has_smartart_or_chart(self, para: Paragraph) -> bool:
        """检查段落是否包含SmartArt或图表。"""
        e = para._element
        return e.find('.//w:drawing', NAMESPACES) is not None or e.find('.//m:oMath', NAMESPACES) is not None

    def _save_checkpoint(self):
        """将提取结果序列化并保存到JSON检查点文件。"""
        import dataclasses
        def serialize(obj):
            if hasattr(obj, '__dataclass_fields__'): return dataclasses.asdict(obj)
            if isinstance(obj, tuple) and len(obj) == 3 and all(isinstance(x, int) for x in obj): return list(obj)
            return obj

        with open(self.checkpoint_file, "w", encoding="utf-8") as f:
            json.dump({k: [serialize(s) for s in getattr(self, k)] for k in
                       ['text_segments', 'table_cell_segments', 'header_footer_segments', 'chart_segments',
                        'smartart_segments']}, f, ensure_ascii=False, indent=2)
