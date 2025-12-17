#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用 xlrd 分析 Excel 文件结构的脚本
"""
import sys

def analyze_excel_with_xlrd(file_path):
    """使用 xlrd 分析 Excel 文件的结构"""
    try:
        import xlrd
        print(f"\n{'='*60}")
        print(f"分析文件: {file_path}")
        print(f"{'='*60}\n")
        
        wb = xlrd.open_workbook(file_path, formatting_info=True)
        print(f"Sheet 数量: {wb.nsheets}")
        print(f"Sheet 名称: {wb.sheet_names()}\n")
        
        for sheet_idx in range(wb.nsheets):
            sheet = wb.sheet_by_index(sheet_idx)
            print(f"\n{'-'*60}")
            print(f"Sheet: {sheet.name}")
            print(f"{'-'*60}")
            print(f"行数: {sheet.nrows}")
            print(f"列数: {sheet.ncols}")
            
            # 显示前 30 行数据
            print(f"\n前 30 行数据预览:")
            max_preview_rows = min(30, sheet.nrows)
            for row_idx in range(max_preview_rows):
                row_data = []
                for col_idx in range(min(20, sheet.ncols)):
                    cell = sheet.cell(row_idx, col_idx)
                    if cell.value:
                        row_data.append(f"[列{col_idx+1}]{cell.value}")
                if row_data:
                    print(f"行 {row_idx+1}: {' | '.join(row_data)}")
            
    except ImportError:
        print("xlrd 未安装，尝试安装...")
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "xlrd"], check=True)
        print("安装完成，重新运行脚本")
        analyze_excel_with_xlrd(file_path)
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        analyze_excel_with_xlrd(sys.argv[1])
    else:
        print("请提供 Excel 文件路径")
