#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置文件 - 定义常量和文件结构映射
"""

# 配分表文件结构
class AllocationTableConfig:
    """配分表文件结构配置"""
    # 元数据行
    METADATA_START_ROW = 0  # 元数据开始行（0-indexed）
    METADATA_END_ROW = 7    # 元数据结束行
    
    # 表头行
    HEADER_ROW = 7          # No, タイプ, 得意先コード等（0-indexed）
    SIZE_ROW = 8            # サイズ行（0-indexed）
    DATA_START_ROW = 9      # 数据开始行（0-indexed）
    
    # 列索引
    COL_NO = 0              # No列
    COL_TYPE = 1            # タイプ列
    COL_STORE_CODE = 2      # 得意先コード列
    COL_RANK = 3            # ランク列
    COL_PRIORITY = 4        # 優先順位列
    COL_STORE_NAME = 5      # 得意先名列
    COL_FIRST_COLOR = 6     # 第一个カラー列（之后是尺码数据）
    
    # 元数据提取
    COMPANY_ROW = 3         # カンパニー行（0-indexed）
    COMPANY_COL = 7         # カンパニー列
    
    DELIVERY_DATE_ROW = 3   # 納期行
    DELIVERY_DATE_COL = 15  # 納期列
    
    PRODUCT_CODE_ROW = 5    # 品番行
    PRODUCT_CODE_COL = 18   # 品番列
    
    # 管理No提取（列31是标签，列34是值）
    KANRI_NO_LABEL_ROW = 3
    KANRI_NO_LABEL_COL = 31
    KANRI_NO_VALUE_COL = 34
    
    # 店着日提取（列31是标签，列34是值）
    STORE_DATE_LABEL_ROW = 5
    STORE_DATE_LABEL_COL = 31
    STORE_DATE_VALUE_COL = 34


# 模板文件结构
class TemplateConfig:
    """模板文件结构配置"""
    # Sheet名称
    PRODUCT_LIST_SHEET = "商品一覧"
    PT_TEMPLATE_SHEET = "PT-1"
    
    # 商品一覧页结构
    PRODUCT_LIST_HEADER_ROW = 1  # 表头行（0-indexed）
    PRODUCT_LIST_DATA_START = 2  # 数据开始行
    PRODUCT_LIST_COL_NO = 0      # No列
    PRODUCT_LIST_COL_CODE = 1    # 商品コード列
    PRODUCT_LIST_COL_COLOR = 2   # カラー列
    PRODUCT_LIST_COL_SIZE = 3    # サイズ列
    PRODUCT_LIST_COL_JAN = 4     # JAN列
    PRODUCT_LIST_COL_ORDER_QTY = 5 # 発注数列
    PRODUCT_LIST_COL_SHIP_QTY = 6  # 出荷数列
    PRODUCT_LIST_COL_DECREASE = 7  # 減産数列
    PRODUCT_LIST_COL_INCREASE = 8  # 増産数列
    
    # 商品一覧页的管理No位置
    PRODUCT_LIST_KANRI_NO_ROW = 0
    PRODUCT_LIST_KANRI_NO_COL = 2
    
    # PT页表头结构
    PT_HEADER_ROW_1 = 0         # 管理No、No.行
    PT_HEADER_ROW_2 = 1         # パターン、品番行
    PT_HEADER_ROW_3 = 2         # 枚数、カラー行
    PT_HEADER_ROW_4 = 3         # 納期、サイズ行
    PT_DATA_HEADER_ROW = 4      # 数据列表头（No., タイプ, ランク...）
    PT_DATA_START_ROW = 5       # 数据开始行
    
    # PT页列索引
    PT_COL_NO = 0               # No.列
    PT_COL_TYPE = 1             # タイプ列
    PT_COL_RANK = 2             # ランク列
    PT_COL_STORE_CODE = 3       # コード列
    PT_COL_STORE_NAME = 4       # 店舗名列
    PT_COL_CTN_NO = 5           # CTN_NO列
    PT_COL_PATTERN = 6          # パターン列
    PT_COL_TOTAL = 7            # 合計列
    PT_COL_FIRST_SKU = 8        # 第一个SKU数据列
    
    # PT页表头固定值位置
    PT_KANRI_NO_COL = 0         # 管理No列（行0）
    PT_PATTERN_NAME_COL = 4     # パターン名列（行1）
    PT_TOTAL_QTY_COL = 4        # 枚数列（行2）
    PT_DELIVERY_DATE_COL = 4    # 納期列（行3）
    PT_NO_LABEL_COL = 5         # "No."标签列（行0）
    PT_PRODUCT_CODE_LABEL_COL = 5  # "品番"标签列（行1）
    PT_COLOR_LABEL_COL = 5      # "カラー"标签列（行2）
    PT_SIZE_LABEL_COL = 5       # "サイズ"标签列（行3）
    
    # SKU序列开始列
    PT_SKU_START_COL = 8        # SKU数据从第9列开始（0-indexed为8）


# 明细表文件结构
class DetailTableConfig:
    """明细表文件结构配置"""
    # 列名
    COL_PRODUCT_CODE = '品番'
    COL_COLOR = 'カラー'
    COL_SIZE = 'サイズ'
    COL_JAN = 'JAN'


# 受渡伝票配置
class DeliveryNoteConfig:
    """受渡伝票生成配置"""
    # 模板路径
    TEMPLATE_NAME = "③受渡伝票_模板（上传系统资料）.xls"
    
    # 写入起始位置
    WRITE_START_ROW = 7         # B8 -> row 7 (0-indexed)
    WRITE_START_COL = 1         # B8 -> col 1 (0-indexed)
    
    # 输出列映射 (相对于B列的偏移)
    # B列 (col 1) -> 受渡伝票NO
    # C列 (col 2) -> ブランド (Brand)
    # D列 (col 3) -> 店舗コード (Store Code)
    # E列 (col 4) -> 品番 (Product Code)
    # F列 (col 5) -> サイズ (Size)
    # G列 (col 6) -> カラー (Color)
    # H列 (col 7) -> 数量 (Quantity)
    
    COL_INDEX_SLIP_NO = 1       # B列
    COL_INDEX_BRAND = 2         # C列
    COL_INDEX_STORE_CODE = 3    # D列
    COL_INDEX_PRODUCT_CODE = 4  # E列
    COL_INDEX_SIZE = 5          # F列
    COL_INDEX_COLOR = 6         # G列
    COL_INDEX_QTY = 7           # H列


# 配分表文件结构
class AllocationConfig:
    """配分表配置"""
    TEMPLATE_NAME = "①箱设定_模板（配分表用）.xlsx"

# ... (AllocationTableConfig remains same)

# ...

# アソート明細配置
class AssortmentConfig:
    """アソート明細生成配置"""
    # 模板路径
    TEMPLATE_NAME = "②アソート明細_模板.xlsx"
    
    # 写入起始位置
    WRITE_START_ROW = 2         # Row 3 (0-indexed)
    WRITE_START_COL = 1         # B列 (0-indexed)
    
    # 输出列索引 (0-indexed)
    COL_INDEX_DELIVERY_CODE = 1      # B列: 届け先コード
    COL_INDEX_DELIVERY_NAME = 2      # C列: 届け先名
    COL_INDEX_SLIP_NO = 3            # D列: 受渡伝票
    COL_INDEX_JAN = 4                # E列: JANコード
    COL_INDEX_MANUFACTURER_CODE = 5  # F列: メーカー品番
    COL_INDEX_QTY = 6                # G列: 汇总(数量)


# 各店铺明细配置
class StoreDetailConfig:
    """各店铺明细生成配置"""
    # 模板路径
    TEMPLATE_NAME = "④各店铺明细_模板.xlsx"
    
    # 写入起始位置
    WRITE_START_ROW = 7         # B8 -> row 7 (0-indexed)
    WRITE_START_COL = 1         # B列 (0-indexed)
    
    # 列索引 (0-indexed)
    COL_INDEX_SLIP_NO = 1       # B列
    COL_INDEX_BRAND = 2         # C列
    COL_INDEX_STORE_CODE = 3    # D列
    COL_INDEX_PRODUCT_CODE = 4  # E列
    COL_INDEX_SIZE = 5          # F列
    COL_INDEX_COLOR = 6         # G列
    COL_INDEX_QTY = 7           # H列



# 文件路径配置
class FileConfig:
    """文件路径配置"""
    DEFAULT_TEMPLATE_NAME = "①箱设定_模板（配分表用）.xlsx"
    OUTPUT_FILE_PREFIX = "【箱設定　上海】"
    OUTPUT_FILE_SUFFIX = "_振分.xlsx"


# 日志配置
class LogConfig:
    """日志配置"""
    LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
    DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
