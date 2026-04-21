"""
检查特定关联型号的详细信息
"""
import sqlite3

def check_specific_aliases(db_path="cable_products_v4.db"):
    """检查特定关联型号"""
    
    print("=" * 80)
    print("检查特定关联型号详情")
    print("=" * 80)
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 检查所有可能的重复情况
        print("\n1. 检查所有关联型号与产品型号的关系...")
        cursor.execute("""
            SELECT 
                ps.spec_id,
                ps.product_model,
                asm.alias_name,
                asm.source,
                asm.confidence,
                CASE 
                    WHEN ps.product_model IS NULL THEN '产品型号为空'
                    WHEN ps.product_model = '' THEN '产品型号为空字符串'
                    WHEN ps.product_model = asm.alias_name THEN '完全相同'
                    WHEN TRIM(ps.product_model) = TRIM(asm.alias_name) THEN '去空格后相同'
                    ELSE '不同'
                END as comparison
            FROM product_specs ps
            JOIN alias_spec_mapping asm ON ps.spec_id = asm.spec_id
            ORDER BY comparison DESC, ps.spec_id
        """)
        
        all_results = cursor.fetchall()
        
        print(f"\n找到 {len(all_results)} 条关联型号记录\n")
        
        # 按比较结果分组
        same_count = 0
        null_count = 0
        different_count = 0
        
        print("-" * 80)
        print("详细列表：")
        print("-" * 80)
        
        for spec_id, product_model, alias_name, source, confidence, comparison in all_results:
            if comparison in ['完全相同', '去空格后相同']:
                same_count += 1
                print(f"\n⚠ {comparison}")
                print(f"   规格编号: {spec_id}")
                print(f"   产品型号: '{product_model}'")
                print(f"   关联型号: '{alias_name}'")
                print(f"   来源: {source}")
            elif comparison in ['产品型号为空', '产品型号为空字符串']:
                null_count += 1
                if null_count <= 5:  # 只显示前5个
                    print(f"\nℹ {comparison}")
                    print(f"   规格编号: {spec_id}")
                    print(f"   产品型号: '{product_model}'")
                    print(f"   关联型号: '{alias_name}'")
                    print(f"   来源: {source}")
            else:
                different_count += 1
        
        print("\n" + "=" * 80)
        print("统计结果：")
        print("=" * 80)
        print(f"完全相同或去空格后相同: {same_count} 条")
        print(f"产品型号为空: {null_count} 条")
        print(f"不同（正常）: {different_count} 条")
        print(f"总计: {len(all_results)} 条")
        
        # 2. 特别检查 YJLV62.LV 和 YJLC23.LV
        print("\n" + "=" * 80)
        print("特别检查指定的型号：")
        print("=" * 80)
        
        for alias in ['YJLV62.LV', 'YJLC23.LV']:
            print(f"\n检查关联型号: {alias}")
            cursor.execute("""
                SELECT 
                    ps.spec_id,
                    ps.product_model,
                    asm.alias_name,
                    asm.source,
                    asm.confidence,
                    asm.created_date
                FROM product_specs ps
                JOIN alias_spec_mapping asm ON ps.spec_id = asm.spec_id
                WHERE asm.alias_name = ?
            """, (alias,))
            
            results = cursor.fetchall()
            if results:
                for spec_id, product_model, alias_name, source, confidence, created_date in results:
                    print(f"  规格编号: {spec_id}")
                    print(f"  产品型号: '{product_model}'")
                    print(f"  关联型号: '{alias_name}'")
                    print(f"  来源: {source}")
                    print(f"  置信度: {confidence}")
                    print(f"  创建时间: {created_date}")
                    
                    if product_model == alias_name:
                        print(f"  ⚠ 状态: 与产品型号相同，需要删除")
                    elif product_model is None or product_model == '':
                        print(f"  ℹ 状态: 产品型号为空，这是有效的关联型号")
                    else:
                        print(f"  ✓ 状态: 与产品型号不同，这是有效的关联型号")
            else:
                print(f"  未找到该关联型号")
        
        conn.close()
        
        return same_count
        
    except Exception as e:
        print(f"\n✗ 检查失败：{str(e)}")
        import traceback
        traceback.print_exc()
        return -1

if __name__ == "__main__":
    import sys
    same_count = check_specific_aliases()
    
    if same_count > 0:
        print(f"\n发现 {same_count} 条需要清理的记录")
        sys.exit(1)
    elif same_count == 0:
        print("\n✓ 没有发现需要清理的记录")
        sys.exit(0)
    else:
        sys.exit(1)
