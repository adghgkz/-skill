#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
会议总结 Word 文档生成脚本
用于生成结构化的会议纪要 Word 文档
"""

import argparse
import os
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def create_meeting_summary(
    date,
    time,
    location,
    attendees,
    agenda,
    content,
    conclusions,
    action_items,
    output_path
):
    """生成会议总结 Word 文档"""

    # 创建文档
    doc = Document()

    # 设置中文字体
    style = doc.styles['Normal']
    style.font.name = '微软雅黑'
    style.font.size = Pt(12)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

    # ===== 标题 =====
    title = doc.add_heading('会议总结', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ===== 会议基本信息 =====
    doc.add_heading('会议基本信息', level=1)

    # 创建表格
    table_info = doc.add_table(rows=4, cols=2)
    table_info.style = 'Light Grid Accent 1'

    # 表格数据
    info_data = [
        ('会议日期', date),
        ('会议时间', time),
        ('会议地点', location),
        ('参会人员', attendees)
    ]

    for i, (label, value) in enumerate(info_data):
        row = table_info.rows[i]
        row.cells[0].text = label
        row.cells[1].text = value
        # 设置单元格背景色
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        row.cells[0].paragraphs[0].runs[0].font.name = '微软雅黑'
        row.cells[1].paragraphs[0].runs[0].font.name = '微软雅黑'

    doc.add_paragraph()

    # ===== 会议议题 =====
    doc.add_heading('会议议题', level=1)
    p_agenda = doc.add_paragraph(agenda)
    p_agenda.paragraph_format.space_before = Pt(0)
    p_agenda.paragraph_format.space_after = Pt(12)

    # ===== 会议背景 =====
    doc.add_heading('会议背景', level=1)
    p_background = doc.add_paragraph('本次会议旨在回顾近期工作进展，讨论存在的问题，并确定下一步工作计划。')
    p_background.paragraph_format.space_before = Pt(0)
    p_background.paragraph_format.space_after = Pt(12)

    # ===== 讨论内容 =====
    doc.add_heading('讨论内容', level=1)
    if content:
        content_paragraphs = content.split('\n')
        for para in content_paragraphs:
            if para.strip():
                p = doc.add_paragraph(para.strip(), style='List Bullet')
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
    else:
        doc.add_paragraph('（本次会议无详细讨论记录）')

    # ===== 会议结论 =====
    doc.add_heading('会议结论', level=1)
    if conclusions:
        conclusions_paragraphs = conclusions.split('\n')
        for para in conclusions_paragraphs:
            if para.strip():
                p = doc.add_paragraph(para.strip())
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
    else:
        doc.add_paragraph('（本次会议无明确结论）')

    # ===== 行动项 =====
    doc.add_heading('行动项', level=1)

    if action_items and action_items.strip():
        # 创建行动项表格
        table_actions = doc.add_table(rows=1, cols=4)
        table_actions.style = 'Light Grid Accent 1'

        # 表头
        header_cells = table_actions.rows[0].cells
        header_cells[0].text = '序号'
        header_cells[1].text = '任务内容'
        header_cells[2].text = '负责人'
        header_cells[3].text = '截止日期'

        # 设置表头格式
        for cell in header_cells:
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 解析行动项
        items = action_items.split(',')
        for idx, item in enumerate(items):
            parts = item.split('|')
            if len(parts) >= 3:
                row_cells = table_actions.add_row().cells
                row_cells[0].text = str(idx + 1)
                row_cells[1].text = parts[0].strip()
                row_cells[2].text = parts[1].strip()
                row_cells[3].text = parts[2].strip()

                # 居中对齐
                for cell in row_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph('（本次会议无行动项）')

    # ===== 页脚 =====
    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = footer.add_run(f'生成时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(128, 128, 128)

    # 保存文档
    doc.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description='生成会议总结 Word 文档')
    parser.add_argument('--date', required=True, help='会议日期')
    parser.add_argument('--time', default='', help='会议时间')
    parser.add_argument('--location', default='', help='会议地点')
    parser.add_argument('--attendees', required=True, help='参会人员（逗号分隔）')
    parser.add_argument('--agenda', default='', help='会议议题')
    parser.add_argument('--content', default='', help='讨论内容')
    parser.add_argument('--conclusions', default='', help='会议结论')
    parser.add_argument('--actions', default='', help='行动项（格式：任务|负责人|日期,任务|负责人|日期）')
    parser.add_argument('--output', default='会议总结.docx', help='输出文件名')

    args = parser.parse_args()

    # 生成文档
    output_path = create_meeting_summary(
        date=args.date,
        time=args.time,
        location=args.location,
        attendees=args.attendees,
        agenda=args.agenda,
        content=args.content,
        conclusions=args.conclusions,
        action_items=args.actions,
        output_path=args.output
    )

    print(f'会议总结已生成: {output_path}')


if __name__ == '__main__':
    main()
