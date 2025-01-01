import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# 父体标题
title = 'Chic Silhouettes: Stylish Phone Cases for the Modern Woman for iPhone.'
# 主题
theme = 'Multi'
# 输入文件
input_file = 'multi_source.xlsx'
# 划线价
standard_price = 16.9
# 卖价格
sale_price = 12.9


def simplify_model_name(size):
    """简化型号名称"""
    # 移除所有空格
    size = size.replace(' ', '')
    
    # 常见型号的简化规则
    replacements = {
        'iPhone': 'IP',
        'ProMax': 'PM',
        'Pro': 'P',
        'Plus': '+',
    }
    
    for full, short in replacements.items():
        size = size.replace(full, short)
    
    return size

def generate_skus(input_file):
    """
    从输入文件生成SKU信息
    返回一个包含所有变体信息的DataFrame
    """
    # 读取输入文件，跳过可能的标题行
    df = pd.read_excel(input_file, header=None)  # 不要自动设置标题行
    
    # 打印原始数据，看看实际内容
    print("Raw data from first row:", df.iloc[0, 1:].values)
    print("Raw data from second row:", df.iloc[1, 1:].values)
    print("Raw data type of second row:", type(df.iloc[1, 1:].values[0]))
    
    # 获取颜色（第一行）和尺寸（第一列，从第三行开始）
    colors = [str(col).strip().title() for col in df.iloc[0, 1:] if pd.notna(col)]
    sizes = [size for size in df.iloc[2:, 0].tolist() if pd.notna(size)]
    
    # 获取每个颜色对应的工厂编号（在第二行）
    factory_codes = {}
    for col_idx, color in enumerate(colors):
        factory_num = str(df.iloc[1, col_idx + 1])
        if factory_num.isdigit() and len(factory_num) == 1:
            factory_num = f'0{factory_num}'
        factory_codes[color] = factory_num
    
    # 准备存储所有SKU信息的列表
    sku_data = []
    
    # 生成基础SKU（父体）
    base_sku = {
        "item_name": f"{title} - {theme}",  # 父体不需要工厂编号
        'item_sku': f'{theme}-BASE',
        'parent_child': 'Parent',
        'feed_product_type': 'cellularphonecase',
        'update_delete': 'PartialUpdate',
        'brand_name': 'ChiCaseVer',
        'manufacturer': 'ChiCaseVer',
        'item_type': 'cell-phone-basic-cases',
        'special_features1': 'All Colors and Sizes',
        'material_type': 'Plastic',
        'compatible_devices': 'iPhone',
        'compatible_phone_models1': 'iPhone',
        'variation_theme': 'SizeName-ColorName'
    }
    sku_data.append(base_sku)
    
    # 为每个颜色和尺寸组合生成变体SKU
    for color in colors:
        factory_code = factory_codes[color]  # 获取当前颜色对应的工厂编号
        for size in sizes:
            simplified_size = simplify_model_name(size)
            child_sku = {
                'item_sku': f'{theme}-{color}-{simplified_size}'.replace(' ', ''),
                'parent_child': 'Child',
                'feed_product_type': 'cellularphonecase',
                'parent_sku': f'{theme}-BASE',
                'relationship_type': 'variation',
                'brand_name': 'ChiCaseVer',
                'update_delete': 'PartialUpdate',
                'item_name': f'Phone Case - {factory_code} - {theme} - {color} - {simplified_size}',
                'manufacturer': 'ChiCaseVer',
                'product_description': 'High quality phone case compatible with IPhone',
                'item_type': 'cell-phone-basic-cases',
                'bullet_point1': 'Perfect fit and protection',
                'bullet_point2': 'Premium quality materials',
                'bullet_point3': 'Easy installation',
                'bullet_point4': 'Full access to all ports and buttons',
                'bullet_point5': 'Slim and lightweight design',
                'included_components': '1 x Phone Case',
                'special_features1': 'Compatible with iPhone',
                'color_name': color,
                'color_map': color,
                'size_name': simplified_size,
                'material_type': 'Plastic',
                'pattern_name': 'Solid',
                'compatible_phone_models1': size,
                'theme': theme,
                'form_factor': 'Phone Case',
                'fulfillment_center_id': 'AMAZON_NA',
                'batteries_required': 'No',
                'standard_price': standard_price,
                'sale_price': sale_price,
                'list_price': 12.9,
                'sale_from_date': (datetime.now() - timedelta(days=2)).strftime('%Y-%m-%d'),
                'sale_end_date': (datetime.now() + timedelta(days=365)).strftime('%Y-%m-%d'),
                'variation_theme': 'SizeName-ColorName',
                'item_height': '1',
                'item_height_unit_of_measure': 'CM',
                'item_width': '16.5',
                'item_width_unit_of_measure': 'CM',
                'item_length': '8',
                'item_length_unit_of_measure': 'CM',
                'item_weight': '0.05',
                'item_weight_unit_of_measure': 'KG',
                'package_height': '1',
                'package_height_unit_of_measure': 'CM',
                'package_width': '20',
                'package_width_unit_of_measure': 'CM',
                'package_length': '13',
                'package_length_unit_of_measure': 'CM',
                'package_weight': '0.05',
                'package_weight_unit_of_measure': 'KG',
            }
            # 只添加有效的字段，不添加None或NaN值
            child_sku = {k: v for k, v in child_sku.items() if pd.notna(v)}
            sku_data.append(child_sku)
    
    return pd.DataFrame(sku_data)

def create_amazon_template(sku_df, template_file, output_file):
    """
    将SKU信息写入亚马逊模板格式
    """
    # 读取模板文件，指定header=2来从第3行读取列名
    template = pd.read_excel(template_file, sheet_name='Template', header=2)
    
    # 找到sku_df中存在且template中也存在的列
    common_columns = list(set(template.columns) & set(sku_df.columns))
    
    # 修改这部分: 确保template有足够的行数来容纳sku_df的数据
    if len(template) < len(sku_df):
        # 创建一个与template列相同的空DataFrame
        template = pd.DataFrame(columns=template.columns)
        # 直接设置所需的行数
        template = template.reindex(range(len(sku_df)))
        
    # 确保sku_df中的所有列都为字符串类型，并将'nan'替换为空字符串
    sku_df = sku_df.fillna('')  # 首先将NaN替换为空字符串
    sku_df = sku_df.astype(str)  # 然后转换为字符串
    sku_df = sku_df.replace('nan', '')  # 最后确保没有'nan'字符串
    
    # 只更新共同存在的列的数据
    template.loc[:len(sku_df)-1, common_columns] = sku_df[common_columns].values
    
    # 保存到新文件，保持所有sheet不变
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        xls = pd.ExcelFile(template_file)
        for sheet_name in xls.sheet_names:
            if sheet_name == 'Template':
                template.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.read_excel(template_file, sheet_name=sheet_name).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )

def main():
    # 文件路径
    template_file = 'amazon_template.xlsx'
    
    # 从输入文件名获取基础名称（移除.xlsx扩展名）
    base_name = input_file.rsplit('.', 1)[0]
    
    # 使用基础名称构建输出文件名
    output_file = f'generated_amazon_upload_{base_name}.xlsx'
    debug_file = f'generated_skus_{base_name}.xlsx'
    
    # 生成SKU信息
    sku_df = generate_skus(input_file)
    
    # 保存中间结果，用于调试
    sku_df = sku_df.astype(str)  # 将所有数据转换为字符串类型
    sku_df.to_excel(debug_file, index=False)
    
    # 创建亚马逊模板
    create_amazon_template(sku_df, template_file, output_file)
    
    print("转换完成！生成的文件：", output_file)
    print("中间SKU信息保存在：", debug_file)
    

if __name__ == "__main__":
    main() 