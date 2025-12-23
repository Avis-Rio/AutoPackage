#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
箱贴生成模块 - 生成箱明細シール PDF
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from config import BoxLabelConfig
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class BoxLabelGenerator:
    def __init__(self, boxes: list, output_path: str):
        """
        初始化箱贴生成器
        
        Args:
            boxes: 箱数据列表 (由BoxSettingReader读取)
            output_path: 输出PDF路径
        """
        self.boxes = boxes
        self.output_path = output_path
        self.font_registered = False
        self._register_font()
        
    def _register_font(self):
        """注册字体，如果失败则使用默认字体"""
        try:
            # 1. 尝试加载配置的 TTF 字体
            font_path = os.path.join(os.path.dirname(__file__), BoxLabelConfig.FONT_PATH)
            if os.path.exists(font_path) and font_path.lower().endswith(('.ttf', '.ttc')):
                pdfmetrics.registerFont(TTFont(BoxLabelConfig.FONT_NAME, font_path))
                self.font_name = BoxLabelConfig.FONT_NAME
                self.font_registered = True
                logger.info(f"Loaded font: {font_path}")
                return

            # 2. 尝试使用 ReportLab 内置的日文 CID 字体 (无需外部文件)
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            cid_font_name = "HeiseiKakuGo-W5" # 黑体
            pdfmetrics.registerFont(UnicodeCIDFont(cid_font_name))
            self.font_name = cid_font_name
            self.font_registered = True
            logger.info(f"Loaded CID font: {cid_font_name}")
            
        except Exception as e:
            logger.error(f"Failed to register font: {e}")
            self.font_name = "Helvetica"

    def generate(self) -> tuple[str, dict]:
        """
        生成箱贴PDF
        Returns:
            tuple: (output_path, stats_dict)
        """
        if not self.boxes:
            logger.warning("No boxes data provided to generate labels.")
            return self.output_path, {}

        c = canvas.Canvas(self.output_path, pagesize=A4)
        total_boxes = len(self.boxes)
        logger.info(f"Generating labels for {total_boxes} boxes...")
        
        # 统计数据
        stats = {
            "box_count": total_boxes,
            "store_count": len(set(b['store_code'] for b in self.boxes)),
            "total_qty": sum(b['total_qty'] for b in self.boxes),
            "pt_count": len(set(b['pattern'] for b in self.boxes)),
            "sku_count": 0
        }
        
        # SKU count needs careful calculation (unique maker_codes?)
        # Or just sum of line items? Let's count unique maker codes across all boxes
        all_maker_codes = set()
        for b in self.boxes:
            for item in b['items']:
                all_maker_codes.add(item['maker_code'])
        stats['sku_count'] = len(all_maker_codes)
        
        # 分页生成 (每页4张)
        for page_idx in range(0, total_boxes, 4):
            page_boxes = self.boxes[page_idx:page_idx+4]
            self._draw_page(c, page_boxes, page_idx, total_boxes)
            c.showPage()
        
        c.save()
        logger.info(f"PDF generated at: {self.output_path}")
        return self.output_path, stats
    
    def _draw_page(self, c, boxes, start_idx, total):
        """绘制单页 (最多4张)"""
        # A4: 210 x 297 mm
        # Labels: 100 x 120 mm
        # Layout: 2x2 with gaps
        
        GAP_X = 4 * mm
        GAP_Y = 4 * mm
        
        # Calculate starting position to center the block of 2x2 labels
        total_content_width = 2 * BoxLabelConfig.LABEL_WIDTH + GAP_X
        total_content_height = 2 * BoxLabelConfig.LABEL_HEIGHT + GAP_Y
        
        start_x = (BoxLabelConfig.PAGE_WIDTH - total_content_width) / 2
        start_y = (BoxLabelConfig.PAGE_HEIGHT - total_content_height) / 2
        
        positions = [
            (0, 1),  # Top-Left
            (1, 1),  # Top-Right
            (0, 0),  # Bottom-Left
            (1, 0)   # Bottom-Right
        ]
        
        for i, box in enumerate(boxes):
            if i >= 4: break
            col, row = positions[i]
            x = start_x + col * (BoxLabelConfig.LABEL_WIDTH + GAP_X)
            y = start_y + row * (BoxLabelConfig.LABEL_HEIGHT + GAP_Y)
            
            self._draw_single_label(c, x, y, box, start_idx + i + 1, total)
            
    def _draw_single_label(self, c, x, y, box, current_no, total):
        """绘制单个箱贴"""
        width = BoxLabelConfig.LABEL_WIDTH
        height = BoxLabelConfig.LABEL_HEIGHT
        padding = 5 * mm
        
        c.saveState()
        c.translate(x, y)
        
        # 1. 边框 (方便裁剪)
        c.setLineWidth(1)
        c.setStrokeColor(colors.black)
        c.rect(0, 0, width, height)
        
        # 2. 头部信息区 (高度约 30mm)
        # 字体设置
        c.setFont(self.font_name, 10)
        
        # 第一行: 店铺番号 店铺名称
        store_text = f"{box['store_code']}   {box['store_name']}"
        c.drawString(padding, height - 10*mm, store_text)
        
        # 第二行: 日期 (右对齐)
        c.setFont(self.font_name, 8)
        date_y = height - 15*mm
        c.drawRightString(width - padding, date_y, f"店着日: {box['store_date']}")
        # 出区日暂时留空
        # if box.get('delivery_date'):
        #      c.drawRightString(width - padding, date_y - 4*mm, f"出区日: {box['delivery_date']}")
        
        # 第三行: 箱ID
        # 箱ID: ★E-{管理No}-{箱号}
        # 如果箱号是 C-001 这种格式，可能需要处理一下
        ctn_clean = str(box['ctn_no']).replace('C-', '')
        try:
            ctn_clean_int = int(ctn_clean)
            ctn_fmt = f"{ctn_clean_int:05d}"
        except:
            ctn_fmt = ctn_clean
            
        box_id = f"★E-{box['kanri_no']}-{ctn_fmt}"
        c.setFont(self.font_name, 9)
        c.drawString(padding, height - 25*mm, f"箱ID: {box_id}")
        
        # 分割线
        header_height = 30 * mm
        c.line(0, height - header_height, width, height - header_height)
        
        # 3. 明细表格区
        # 表头: 部门 | メーカー品番 | 品名 | 数量
        table_data = [['部門', 'メーカー品番', '品名', '数量']]
        
        # 填充数据
        # 动态计算可用高度: 120 - 30(Head) - 20(Foot) = 70mm
        # 行高 4.5mm => ~14-15行
        max_rows = 14
        items = box['items']
        
        for item in items[:max_rows]:
            # 修正メーカー品番: Brand(Dept) - Product - Size - Color
            # item['maker_code'] is currently Product-Color-Size (from BoxSettingReader)
            # We need to parse it and reconstruct.
            # However, if parsing fails, fallback to original logic.
            
            maker_code = item['maker_code']
            try:
                parts = maker_code.split('-')
                if len(parts) >= 3:
                    product = parts[0]
                    color = parts[1]
                    size = parts[2]
                    # Target: Brand-Product-Size-Color
                    # Brand comes from box['dept']
                    maker_code = f"{box['dept']}-{product}-{size}-{color}"
            except Exception:
                pass # Keep original if parsing fails

            table_data.append([
                "", # 部门列留空
                maker_code,
                item['product_name'],
                str(item['qty'])
            ])
            
        if len(items) > max_rows:
            table_data.append(['...', '...', '...', '...'])
            
        # 创建表格
        # 列宽分配: 12mm, 38mm, 35mm, 10mm (Total 95mm < 100mm)
        col_widths = [12*mm, 38*mm, 35*mm, 10*mm]
        t = Table(table_data, colWidths=col_widths, rowHeights=4.5*mm)
        
        t.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), self.font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 7),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (-1, 0), (-1, -1), 'CENTER'), # 数量居中
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            # 表头背景
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ]))
        
        # 计算表格位置
        w, h = t.wrap(width, height)
        # 底部留 20mm + buffer
        table_bottom_y = height - header_height - h - 2*mm
        if table_bottom_y < 20*mm:
            table_bottom_y = 20*mm # 防止覆盖Footer
            
        t.drawOn(c, 2.5*mm, table_bottom_y)
        
        # 4. 底部汇总区 (高度 20mm)
        footer_y = 20 * mm
        c.line(0, footer_y, width, footer_y)
        
        c.setFont(self.font_name, 9)
        
        # 左侧：C/No.
        c.drawString(padding, 13*mm, f"C/No. {box['ctn_no']}")
        
        # 右侧：入数
        c.drawRightString(width - padding, 13*mm, f"入数  {box['total_qty']} PCS")
        
        # 底部中间：页码/总箱数
        c.drawCentredString(width / 2, 5*mm, f"{current_no} / {total}")
        
        c.restoreState()

if __name__ == "__main__":
    # 测试代码
    pass
