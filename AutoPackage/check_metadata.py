#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全面检查配分表开头所有单元格
"""
import xlrd

def check_all_cells():
    """检查配分表所有开头单元格"""
    file_path = "配分表（15305078）_20251205100542.xls"
    
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)  # 第一个sheet
    
    print("=" * 80)
    print("配分表第一个sheet的所有非空单元格（前10行）")
    print("=" * 80)
    
    # 检查前10行所有列
    for row in range(min(10, sheet.nrows)):
        has_data = False
        row_data = []
        for col in range(sheet.ncols):
            cell = sheet.cell(row, col)
            if cell.value:
                has_data = True
                row_data.append(f"[{col}]{cell.value}")
        
        if has_data:
            print(f"\n行{row}: {' | '.join(row_data)}")

if __name__ == "__main__":
    check_all_cells()
