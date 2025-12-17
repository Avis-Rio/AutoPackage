#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel读取模块 - 读取配分表数据
"""
import xlrd
import openpyxl
from config import AllocationTableConfig
from typing import Dict, List, Tuple


class AllocationTableReader:
    """配分表读取器"""
    
    def __init__(self, file_path: str):
        """
        初始化读取器
        
        Args:
            file_path: 配分表文件路径
        """
        self.file_path = file_path
        self.workbook = None
        self.metadata = {}
        self.products_data = []  # 存储所有品番的数据
        self.engine = 'xlrd' # 'xlrd' or 'openpyxl'
        
    def read(self) -> Dict:
        """
        读取配分表文件
        
        Returns:
            包含所有数据的字典
        """
        try:
            if self.file_path.lower().endswith('.xlsx'):
                return self._read_xlsx()
            else:
                return self._read_xls()
                
        except Exception as e:
            raise Exception(f"读取配分表文件失败: {e}")

    def _read_xls(self) -> Dict:
        """使用xlrd读取.xls文件"""
        self.engine = 'xlrd'
        self.workbook = xlrd.open_workbook(self.file_path, formatting_info=False)
        
        # 读取所有品番sheet（跳过可能的汇总sheet）
        for sheet_idx in range(self.workbook.nsheets):
            sheet = self.workbook.sheet_by_index(sheet_idx)
            sheet_name = sheet.name
            
            # 跳过空sheet或太小的sheet
            if sheet.nrows < AllocationTableConfig.DATA_START_ROW:
                continue
            
            # 读取该品番的数据
            product_data = self._read_product_sheet(sheet, sheet_name)
            if product_data:
                self.products_data.append(product_data)
        
        return {
            'metadata': self.metadata,
            'products': self.products_data
        }

    def _read_xlsx(self) -> Dict:
        """使用openpyxl读取.xlsx文件"""
        self.engine = 'openpyxl'
        self.workbook = openpyxl.load_workbook(self.file_path, data_only=True)
        
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            
            # 跳过空sheet或太小的sheet
            if sheet.max_row < AllocationTableConfig.DATA_START_ROW:
                continue
                
            # 读取该品番的数据
            product_data = self._read_product_sheet(sheet, sheet_name)
            if product_data:
                self.products_data.append(product_data)
                
        return {
            'metadata': self.metadata,
            'products': self.products_data
        }
    
    def _get_cell_value(self, sheet, row, col):
        """获取单元格值，屏蔽不同库的差异"""
        if self.engine == 'xlrd':
            try:
                if row >= sheet.nrows or col >= sheet.ncols:
                    return None
                return sheet.cell(row, col).value
            except:
                return None
        else: # openpyxl
            try:
                # openpyxl is 1-based
                return sheet.cell(row=row+1, column=col+1).value
            except:
                return None

    def _get_sheet_dims(self, sheet):
        """获取sheet的维度(rows, cols)"""
        if self.engine == 'xlrd':
            return sheet.nrows, sheet.ncols
        else:
            return sheet.max_row, sheet.max_column

    def _read_product_sheet(self, sheet, sheet_name: str) -> Dict:
        """
        读取单个品番sheet的数据
        
        Args:
            sheet: sheet对象
            sheet_name: sheet名称（品番）
            
        Returns:
            该品番的数据字典
        """
        try:
            # 提取元数据
            metadata = self._extract_metadata(sheet)
            
            # 读取表头（颜色和尺码信息）
            sku_columns = self._read_sku_columns(sheet)
            
            # 读取店铺数据
            stores_data = self._read_stores_data(sheet, sku_columns)
            
            return {
                'product_code': sheet_name,  # 品番
                'metadata': metadata,
                'sku_columns': sku_columns,   # SKU列信息
                'stores': stores_data         # 店铺数据
            }
            
        except Exception as e:
            print(f"警告: 读取sheet '{sheet_name}' 失败: {e}")
            return None
    
    def _extract_metadata(self, sheet) -> Dict:
        """提取元数据（管理No、納期等）"""
        metadata = {}
        nrows, ncols = self._get_sheet_dims(sheet)
        
        try:
            # 提取カンパニー
            if nrows > AllocationTableConfig.COMPANY_ROW:
                val = self._get_cell_value(sheet, 
                    AllocationTableConfig.COMPANY_ROW,
                    AllocationTableConfig.COMPANY_COL)
                if val:
                    metadata['company'] = val
            
            # 提取納期
            if nrows > AllocationTableConfig.DELIVERY_DATE_ROW:
                val = self._get_cell_value(sheet,
                    AllocationTableConfig.DELIVERY_DATE_ROW,
                    AllocationTableConfig.DELIVERY_DATE_COL)
                if val:
                    metadata['delivery_date'] = val
            
            # 提取品番（从元数据区域）
            if nrows > AllocationTableConfig.PRODUCT_CODE_ROW:
                val = self._get_cell_value(sheet,
                    AllocationTableConfig.PRODUCT_CODE_ROW,
                    AllocationTableConfig.PRODUCT_CODE_COL)
                if val:
                    metadata['product_code'] = val
            
            # 提取管理No（直接从已知位置读取）
            if (nrows > AllocationTableConfig.KANRI_NO_LABEL_ROW and 
                ncols > AllocationTableConfig.KANRI_NO_VALUE_COL):
                val = self._get_cell_value(sheet,
                    AllocationTableConfig.KANRI_NO_LABEL_ROW,
                    AllocationTableConfig.KANRI_NO_VALUE_COL)
                if val:
                    # 转换为字符串，如果是数字则转为整数字符串
                    if isinstance(val, float):
                        val = int(val)
                    metadata['kanri_no'] = str(val)
            
            # 提取店着日（直接从已知位置读取）
            if (nrows > AllocationTableConfig.STORE_DATE_LABEL_ROW and
                ncols > AllocationTableConfig.STORE_DATE_VALUE_COL):
                val = self._get_cell_value(sheet,
                    AllocationTableConfig.STORE_DATE_LABEL_ROW,
                    AllocationTableConfig.STORE_DATE_VALUE_COL)
                if val:
                    metadata['store_date'] = str(val)
            
        except Exception as e:
            print(f"警告: 提取元数据失败: {e}")
        
        # 使用第一个sheet的元数据作为全局元数据
        if not self.metadata:
            self.metadata = metadata.copy()
        
        return metadata
    
    def _read_sku_columns(self, sheet) -> List[Dict]:
        """
        读取SKU列信息（カラー和サイズ的组合）
        """
        sku_columns = []
        nrows, ncols = self._get_sheet_dims(sheet)
        
        # 读取カラー行（COL_FIRST_COLOR之后的列）
        color_row = AllocationTableConfig.HEADER_ROW
        size_row = AllocationTableConfig.SIZE_ROW
        
        # 从第一个颜色列开始遍历
        col_idx = AllocationTableConfig.COL_FIRST_COLOR
        current_color = None
        
        while col_idx < ncols:
            # 读取颜色
            color_val = self._get_cell_value(sheet, color_row, col_idx)
            if color_val:
                # 新的颜色值
                color_str = str(color_val).strip()
                # 过滤掉非数据列（如"カラー"标签列）
                if color_str and color_str != "カラー":
                    current_color = color_str
            
            # 读取尺码
            size_val = self._get_cell_value(sheet, size_row, col_idx)
            size_value = str(size_val).strip() if size_val else ""
            
            # 过滤掉无效的尺码值（如"サイズ"、"合計"等）
            invalid_sizes = ["サイズ", "合計", "合计", "小計", "Total"]
            is_valid_size = size_value and size_value not in invalid_sizes
            
            # 如果有颜色和有效尺码，记录这个SKU列
            if current_color and is_valid_size:
                sku_columns.append({
                    'column_index': col_idx,
                    'color': current_color,
                    'size': size_value
                })
            
            col_idx += 1
        
        return sku_columns
    
    def _read_stores_data(self, sheet, sku_columns: List[Dict]) -> List[Dict]:
        """
        读取店铺数据
        """
        stores_data = []
        nrows, ncols = self._get_sheet_dims(sheet)
        
        # 从数据开始行读取
        for row_idx in range(AllocationTableConfig.DATA_START_ROW, nrows):
            # 读取店铺基本信息
            no_val = self._get_cell_value(sheet, row_idx, AllocationTableConfig.COL_NO)
            store_code_val = self._get_cell_value(sheet, row_idx, AllocationTableConfig.COL_STORE_CODE)
            store_name_val = self._get_cell_value(sheet, row_idx, AllocationTableConfig.COL_STORE_NAME)
            type_val = self._get_cell_value(sheet, row_idx, AllocationTableConfig.COL_TYPE)
            
            # 如果没有店铺代码或名称，跳过这行（可能是汇总行）
            if not store_code_val or not store_name_val:
                continue
            
            # 读取该店铺的SKU配货数量
            sku_quantities = {}
            for sku_info in sku_columns:
                col_idx = sku_info['column_index']
                qty_val = self._get_cell_value(sheet, row_idx, col_idx)
                qty = 0
                if qty_val:
                    try:
                        qty = float(qty_val)
                    except:
                        qty = 0
                
                # 构建SKU key (color_size)
                sku_key = f"{sku_info['color']}_{sku_info['size']}"
                sku_quantities[sku_key] = qty
            
            # 只有当至少有一个SKU数量大于0时才记录这个店铺
            if any(qty > 0 for qty in sku_quantities.values()):
                store_data = {
                    'no': no_val if no_val else "",
                    'type': str(int(type_val)) if isinstance(type_val, (int, float)) and type_val else str(type_val) if type_val else "",
                    'store_code': str(int(store_code_val)) if isinstance(store_code_val, float) else str(store_code_val),
                    'store_name': str(store_name_val),
                    'sku_quantities': sku_quantities
                }
                
                # 读取ランク（如果有）
                if AllocationTableConfig.COL_RANK < ncols:
                    rank_val = self._get_cell_value(sheet, row_idx, AllocationTableConfig.COL_RANK)
                    if rank_val:
                        store_data['rank'] = str(rank_val)
                
                stores_data.append(store_data)
        
        return stores_data
    
    def close(self):
        """关闭工作簿"""
        if self.workbook:
            if self.engine == 'xlrd':
                self.workbook.release_resources()
            elif self.engine == 'openpyxl':
                self.workbook.close()
            self.workbook = None
