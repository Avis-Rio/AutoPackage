#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试脚本 - 验证配分表转换功能
"""
import os
from excel_reader import AllocationTableReader
from data_transformer import DataTransformer
from template_writer import TemplateWriter


def test_conversion():
    """测试转换功能"""
    print("=" * 60)
    print("配分表转换功能测试")
    print("=" * 60)
    
    # 文件路径
    allocation_file = "配分表（15305078）_20251205100542.xls"
    template_file = "template.xlsx"
    output_file = "test_output.xlsx"
    
    # 检查文件是否存在
    if not os.path.exists(allocation_file):
        print(f"错误: 找不到配分表文件: {allocation_file}")
        return
    
    if not os.path.exists(template_file):
        print(f"错误: 找不到模板文件: {template_file}")
        return
    
    try:
        # 1. 读取配分表
        print(f"\n步骤 1: 读取配分表...")
        print(f"  文件: {allocation_file}")
        
        reader = AllocationTableReader(allocation_file)
        allocation_data = reader.read()
        reader.close()
        
        products_count = len(allocation_data['products'])
        print(f"  ✓ 成功读取 {products_count} 个品番")
        
        # 显示元数据
        metadata = allocation_data.get('metadata', {})
        if metadata:
            print(f"\n  元数据信息:")
            if 'kanri_no' in metadata:
                print(f"    - 管理No: {metadata['kanri_no']}")
            if 'store_date' in metadata:
                print(f"    - 店着日: {metadata['store_date']}")
            if 'delivery_date' in metadata:
                print(f"    - 納期: {metadata['delivery_date']}")
        
        # 统计店铺总数
        total_stores = sum(
            len(product['stores']) 
            for product in allocation_data['products']
        )
        print(f"  ✓ 店铺数据总数: {total_stores}")

        
        # 2. 数据转换
        print(f"\n步骤 2: 数据转换...")
        
        transformer = DataTransformer(allocation_data)
        transformed_data = transformer.transform()
        
        sku_count = len(transformed_data['skus'])
        pt_count = len(transformed_data['pt_groups'])
        
        print(f"  ✓ SKU总数: {sku_count}")
        print(f"  ✓ PT分组数: {pt_count}")
        
        # 显示每个PT的详情
        print(f"\n  PT分组详情:")
        for pt_group in transformed_data['pt_groups']:
            store_count = len(pt_group['stores'])
            total_qty = pt_group['total_qty']
            print(f"    - {pt_group['pt_name']}: {store_count} 个店铺, {total_qty} 件商品")
        
        # 显示前10个SKU
        print(f"\n  前10个SKU（已排序）:")
        for idx, sku in enumerate(transformed_data['skus'][:10]):
            print(f"    {idx+1}. 品番:{sku['product_code']} カラー:{sku['color']} サイズ:{sku['size']}")
        
        # 3. 生成输出文件
        print(f"\n步骤 3: 生成输出文件...")
        print(f"  模板: {template_file}")
        print(f"  输出: {output_file}")
        
        writer = TemplateWriter(template_file, output_file)
        output_path = writer.write(transformed_data)
        
        print(f"  ✓ 成功生成文件: {output_path}")
        
        # 验证输出文件
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"  ✓ 文件大小: {file_size:,} bytes")
        
        print("\n" + "=" * 60)
        print("✅ 测试完成！")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_conversion()
