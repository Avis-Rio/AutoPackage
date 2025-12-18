#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据转换模块 - 将配分表数据转换为箱設定格式
"""
from typing import Dict, List, Tuple
from collections import defaultdict
import re


class DataTransformer:
    """数据转换器"""
    
    def __init__(self, allocation_data: Dict, jan_map: Dict[Tuple[str, str, str], str] = None):
        """
        初始化转换器
        
        Args:
            allocation_data: 从AllocationTableReader读取的数据
            jan_map: JANCODE映射字典 {(品番, 颜色, 尺码): JANCODE}
        """
        self.allocation_data = allocation_data
        self.jan_map = jan_map or {}
        self.all_skus = []  # 所有唯一的SKU列表（排序后）
        self.pt_groups = []  # PT分组结果
        self.ctn_counter = 1  # 箱号计数器
        
    def transform(self) -> Dict:
        """
        执行转换
        
        Returns:
            转换后的数据，包含PT分组和SKU信息
        """
        # 1. 聚合所有SKU
        self._aggregate_skus()
        
        # 2. 对SKU排序
        self._sort_skus()
        
        # 3. 注入 JANCODE
        self._inject_jancodes()
        
        # 4. 计算 SKU 总数
        self._calculate_sku_totals()
        
        # 5. 分析店铺配比并分组PT
        self._group_by_pattern()
        
        # 5. 分配箱号
        self._assign_ctn_numbers()
        
        return {
            'metadata': self.allocation_data['metadata'],
            'skus': self.all_skus,
            'pt_groups': self.pt_groups
        }
    
    def _clean_product_code(self, code):
        """去除品番中的括号及内容，例如 '14003(2)' -> '14003'"""
        # 将全角括号转为半角
        code = str(code).replace('（', '(').replace('）', ')')
        # 去除 (..) 内容
        return re.sub(r'\(.*?\)', '', code).strip()

    def _inject_jancodes(self):
        """注入 JANCODE 到 SKU 信息中"""
        self.logs = []
        match_count = 0
        fail_count = 0
        
        self.logs.append(f"开始匹配 JANCODE (明细表共 {len(self.jan_map)} 条)")
        if self.jan_map:
            self.logs.append(f"明细表键样例(前3): {list(self.jan_map.keys())[:3]}")

        for sku in self.all_skus:
            # 基础清理 & 品番去括号
            original_p_code = str(sku['product_code']).strip()
            p_code = self._clean_product_code(original_p_code)
            
            # 更新 sku 中的 product_code (写入模板时也生效)
            if p_code != original_p_code:
                # self.logs.append(f"品番清洗: {original_p_code} -> {p_code}")
                sku['product_code'] = p_code
            
            color = str(sku['color']).strip()
            size = str(sku['size']).strip()
            
            # 1. 精确匹配
            key = (p_code, color, size)
            jan = self.jan_map.get(key, "")
            
            # 2. 模糊匹配：尝试去除颜色前导零 (例如 '003' -> '3')
            if not jan:
                color_stripped = color.lstrip('0')
                # 如果全是0，变成空串了，要恢复成'0'
                if not color_stripped and '0' in color: color_stripped = "0"
                
                if color != color_stripped:
                    key_retry = (p_code, color_stripped, size)
                    jan = self.jan_map.get(key_retry, "")
                    if jan:
                        self.logs.append(f"模糊匹配成功(去零): {key} -> {key_retry}")

            # 3. 模糊匹配：尝试去除尺码前导零 (例如 '09' -> '9')
            if not jan:
                size_stripped = size.lstrip('0')
                if not size_stripped and '0' in size: size_stripped = "0"
                
                if size != size_stripped:
                    key_retry = (p_code, color, size_stripped)
                    jan = self.jan_map.get(key_retry, "")
                    if jan:
                        self.logs.append(f"模糊匹配成功(尺码去零): {key} -> {key_retry}")
                        
            # 4. 组合拳 (颜色尺码都去零)
            if not jan:
                color_stripped = color.lstrip('0') or "0"
                size_stripped = size.lstrip('0') or "0"
                key_retry = (p_code, color_stripped, size_stripped)
                jan = self.jan_map.get(key_retry, "")
                if jan:
                     self.logs.append(f"模糊匹配成功(双去零): {key} -> {key_retry}")

            if jan:
                sku['jan_code'] = jan
                match_count += 1
            else:
                fail_count += 1
                if fail_count <= 5:
                    self.logs.append(f"匹配失败: {key}")
        
        self.logs.append(f"JANCODE 匹配结果: 成功 {match_count}, 失败 {fail_count}")

    def _calculate_sku_totals(self):
        """计算每个SKU的全局总数量"""
        # 初始化计数器
        sku_totals = defaultdict(int) # key: (product_code, color, size)
        
        # 调试日志
        debug_count = 0
        
        for product_data in self.allocation_data['products']:
            # 使用清洗后的品番
            product_code = self._clean_product_code(product_data['product_code'])
            
            for store in product_data['stores']:
                for sku_key, qty in store['sku_quantities'].items():
                    # sku_key 是 "Color_Size"
                    if '_' in sku_key:
                        parts = sku_key.rsplit('_', 1)
                        color = parts[0]
                        size = parts[1]
                        
                        key = (product_code, color, size)
                        sku_totals[key] += qty
                        
                        if debug_count < 3:
                            # print(f"DEBUG: Accumulate {key} += {qty}")
                            debug_count += 1
        
        # 将总数注入到 self.all_skus
        inject_count = 0
        for sku in self.all_skus:
            key = (sku['product_code'], sku['color'], sku['size'])
            total = sku_totals.get(key, 0)
            sku['total_qty'] = total
            if total > 0:
                inject_count += 1
                
        self.logs.append(f"SKU总数计算完成: {inject_count} 个SKU有数量")
    
    def _aggregate_skus(self):
        """聚合所有品番中的SKU"""
        sku_set = set()
        
        for product_data in self.allocation_data['products']:
            product_code = product_data['product_code']
            
            # 从SKU列信息中提取
            for sku_col in product_data['sku_columns']:
                sku = {
                    'product_code': product_code,
                    'color': sku_col['color'],
                    'size': sku_col['size']
                }
                # 使用tuple作为set的key
                sku_tuple = (product_code, sku_col['color'], sku_col['size'])
                sku_set.add(sku_tuple)
        
        # 转换为列表
        self.all_skus = [
            {'product_code': sku[0], 'color': sku[1], 'size': sku[2]}
            for sku in sku_set
        ]
    
    def _sort_skus(self):
        """对SKU按品番、カラー、サイズ排序"""
        self.all_skus.sort(key=lambda x: (
            str(x['product_code']),
            str(x['color']),
            str(x['size'])
        ))
    
    def _group_by_pattern(self):
        """根据配比模式对店铺进行分组"""
        # 1. 首先聚合同一店铺在所有品番下的数据
        merged_stores_map = {}
        
        for product_data in self.allocation_data['products']:
            # 使用清洗后的品番，确保与sku中的key一致
            product_code = self._clean_product_code(product_data['product_code'])
            
            for store in product_data['stores']:
                # 修改聚合逻辑：仅使用 store_code 作为唯一标识，忽略 type
                # 这样同一店铺即使有不同的 type 也会被合并到一起
                store_key = store['store_code']
                
                if store_key not in merged_stores_map:
                    merged_stores_map[store_key] = {
                        'no': store['no'],
                        'type': store['type'], # 保留第一个遇到的type
                        'store_code': store['store_code'],
                        'store_name': store['store_name'],
                        'rank': store.get('rank', ''),
                        'raw_sku_quantities': {}  # 临时存储：{"product_color_size": qty}
                    }
                
                # 合并该店铺在该品番下的SKU数量
                for sku_short_key, qty in store['sku_quantities'].items():
                    # sku_short_key 是 "Color_Size"
                    # 组合成全局唯一的 "Product_Color_Size"
                    full_key = f"{product_code}_{sku_short_key}"
                    
                    # 累加数量
                    current_qty = merged_stores_map[store_key]['raw_sku_quantities'].get(full_key, 0)
                    merged_stores_map[store_key]['raw_sku_quantities'][full_key] = current_qty + qty

        # 2. 为聚合后的店铺构建全局配比向量
        all_stores = []
        
        for store_data in merged_stores_map.values():
            pattern_vector = []
            full_quantities = {}
            
            # 遍历所有已排序的全局SKU列表
            for sku in self.all_skus:
                full_key = f"{sku['product_code']}_{sku['color']}_{sku['size']}"
                
                # 获取数量，默认为0
                qty = store_data['raw_sku_quantities'].get(full_key, 0)
                
                pattern_vector.append(qty)
                full_quantities[full_key] = qty
            
            # 更新店铺数据结构
            store_data['pattern_vector'] = pattern_vector
            store_data['sku_quantities'] = full_quantities
            del store_data['raw_sku_quantities']  # 清理临时数据
            
            all_stores.append(store_data)
        
        # 3. 按配比模式分组
        pattern_groups = defaultdict(list)
        
        for store in all_stores:
            # 使用配比向量的tuple作为分组key
            # 注意：不再包含 type，完全由配比向量决定分组
            pattern_key = tuple(store['pattern_vector'])
            pattern_groups[pattern_key].append(store)
        
        # 4. 转换为PT组列表
        pt_index = 1
        for pattern_key, stores in pattern_groups.items():
            # 计算该PT的总枚数 (应为单个店铺的配货总数)
            # 因为同组内所有店铺的配货模式都一样，所以取第一个店铺的总数即可
            # 注意：之前的逻辑是 sum(sum(store...)) 计算了该组所有店铺的总和，这是错误的
            
            if not stores:
                continue
                
            # 取第一个店铺计算单店总配货量
            first_store = stores[0]
            pt_total_qty = sum(first_store['sku_quantities'].values())
            
            # 如果总数量为0，跳过（可能是空配比）
            if pt_total_qty == 0:
                continue

            self.pt_groups.append({
                'pt_name': f'PT-{pt_index}',
                'pattern_vector': list(pattern_key),
                'stores': stores,
                'total_qty': int(pt_total_qty) # 这里存储的是单店配货总数，用于表头显示
            })
            pt_index += 1
    
    def _build_pattern_vector(self, product_code: str, sku_quantities: Dict) -> List[float]:
        """
        构建配比向量（按照排序后的SKU顺序）
        
        Args:
            product_code: 当前店铺数据所属的品番
            sku_quantities: 店铺的SKU配货数量（key格式: "color_size"）
            
        Returns:
            配比数量列表，顺序与self.all_skus一致
        """
        pattern = []
        
        for sku in self.all_skus:
            # 只考虑当前品番的SKU
            sku_key = f"{sku['color']}_{sku['size']}"
            
            # 如果是当前品番且有配货数量，使用该数量；否则为0
            if sku['product_code'] == product_code and sku_key in sku_quantities:
                qty = sku_quantities[sku_key]
            else:
                qty = 0
            
            pattern.append(qty)
        
        return pattern
    
    def _build_full_sku_quantities(self, product_code: str, sku_quantities: Dict) -> Dict:
        """
        构建完整的SKU数量字典（包含所有SKU，按照排序后的顺序）
        
        Args:
            product_code: 当前店铺数据所属的品番
            sku_quantities: 店铺的SKU配货数量
            
        Returns:
            完整的SKU数量字典，key格式: "product_color_size"
        """
        full_quantities = {}
        
        for sku in self.all_skus:
            sku_full_key = f"{sku['product_code']}_{sku['color']}_{sku['size']}"
            sku_key = f"{sku['color']}_{sku['size']}"
            
            # 如果是当前品番且有配货，使用该数量
            if sku['product_code'] == product_code and sku_key in sku_quantities:
                full_quantities[sku_full_key] = sku_quantities[sku_key]
            else:
                full_quantities[sku_full_key] = 0
        
        return full_quantities
    
    def _assign_ctn_numbers(self):
        """为每个PT组中的店铺分配箱号"""
        for pt_group in self.pt_groups:
            for store in pt_group['stores']:
                # 每个店铺的每个タイプ分配一个箱号
                store['ctn_no'] = self.ctn_counter
                self.ctn_counter += 1
                
                # 计算该箱的合计数量
                store['total_qty'] = int(sum(store['sku_quantities'].values()))
