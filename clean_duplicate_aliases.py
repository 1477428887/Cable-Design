"""
清理产品列表中关联型号与产品名称相同的记录
"""
import sqlite3
import sys
import os

def clean_duplicate_aliases(db_path="cable_products_v4.db"):
    """删除关联型号与产品名称相同的映射记录"""
    
    print("=" * 70)
    print("清理重复关联型号记录")
    print("=" * 70)
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 1. 查找所有产品及其关联型号
        print("\n正在分析数据库...")
        cursor.execute("""
            SELECT 
                ps.spec_id,
                ps.product_model,
                asm.alias_name,
                asm.id as mapping_id
            FROM product_specs ps
            LEFT JOIN alias_spec_mapping asm ON ps.spec_id = asm.spec_id
            WHERE ps.product_model IS NOT NULL 
            AND ps.product_model != ''
            AND asm.alias_name IS NOT NULL
        """)
        
        all_mappings = cursor.fetchall()
        print(f"✓ 找到 {len(all_mappings)} 条关联型号记录")
        
        # 2. 找出需要删除的记录（关联型号与产品名称相同）
        duplicate_mappings = []
        for spec_id, product_model, alias_name, mapping_id in all_mappings:
            if product_model and alias_name and product_model.strip() == alias_name.strip():
                duplicate_mappings.append({
                    'spec_id': spec_id,
                    'product_model': product_model,
                    'alias_name': alias_name,
                    'mapping_id': mapping_id
                })
        
        if not duplicate_mappings:
            print("\n✓ 没有找到需要清理的重复记录")
            conn.close()
            return True
        
        print(f"\n找到 {len(duplicate_mappings)} 条重复记录需要删除：")
        print("-" * 70)
        
        # 显示前10条示例
        for i, dup in enumerate(duplicate_mappings[:10]):
            print(f"{i+1}. 规格编号: {dup['spec_id']}")
            print(f"   产品型号: {dup['product_model']}")
            print(f"   重复的关联型号: {dup['alias_name']}")
            print()
        
        if len(duplicate_mappings) > 10:
            print(f"... 还有 {len(duplicate_mappings) - 10} 条记录")
        
        print("-" * 70)
        
        # 3. 确认删除
        response = input(f"\n是否删除这 {len(duplicate_mappings)} 条重复记录？(y/n): ")
        if response.lower() != 'y':
            print("操作已取消")
            conn.close()
            return False
        
        # 4. 执行删除
        print("\n正在删除重复记录...")
        deleted_count = 0
        
        for dup in duplicate_mappings:
            cursor.execute("""
                DELETE FROM alias_spec_mapping 
                WHERE id = ?
            """, (dup['mapping_id'],))
            deleted_count += 1
        
        conn.commit()
        
        print(f"✓ 成功删除 {deleted_count} 条重复记录")
        
        # 5. 验证结果
        print("\n正在验证清理结果...")
        cursor.execute("""
            SELECT COUNT(*) 
            FROM product_specs ps
            JOIN alias_spec_mapping asm ON ps.spec_id = asm.spec_id
            WHERE ps.product_model = asm.alias_name
        """)
        
        remaining = cursor.fetchone()[0]
        if remaining == 0:
            print("✓ 验证通过：所有重复记录已清理")
        else:
            print(f"⚠ 警告：仍有 {remaining} 条重复记录")
        
        # 6. 显示清理后的统计信息
        print("\n" + "=" * 70)
        print("清理后的统计信息")
        print("=" * 70)
        
        cursor.execute("SELECT COUNT(*) FROM product_specs")
        total_specs = cursor.fetchone()[0]
        print(f"产品规格总数: {total_specs}")
        
        cursor.execute("SELECT COUNT(*) FROM alias_spec_mapping")
        total_mappings = cursor.fetchone()[0]
        print(f"关联型号映射总数: {total_mappings}")
        
        cursor.execute("""
            SELECT COUNT(DISTINCT spec_id) 
            FROM alias_spec_mapping
        """)
        specs_with_aliases = cursor.fetchone()[0]
        print(f"有关联型号的产品数: {specs_with_aliases}")
        
        conn.close()
        
        print("\n" + "=" * 70)
        print("清理完成！")
        print("=" * 70)
        
        return True
        
    except Exception as e:
        print(f"\n✗ 清理失败：{str(e)}")
        import traceback
        traceback.print_exc()
        return False

def preview_duplicates(db_path="cable_products_v4.db"):
    """预览需要清理的重复记录（不执行删除）"""
    
    print("=" * 70)
    print("预览重复关联型号记录")
    print("=" * 70)
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
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
            WHERE ps.product_model = asm.alias_name
            ORDER BY ps.spec_id
        """)
        
        duplicates = cursor.fetchall()
        
        if not duplicates:
            print("\n✓ 没有找到重复记录")
            conn.close()
            return
        
        print(f"\n找到 {len(duplicates)} 条重复记录：")
        print("-" * 70)
        
        for i, (spec_id, product_model, alias_name, source, confidence, created_date) in enumerate(duplicates, 1):
            print(f"{i}. 规格编号: {spec_id}")
            print(f"   产品型号: {product_model}")
            print(f"   重复的关联型号: {alias_name}")
            print(f"   来源: {source}")
            print(f"   置信度: {confidence}")
            print(f"   创建时间: {created_date}")
            print()
        
        conn.close()
        
    except Exception as e:
        print(f"\n✗ 预览失败：{str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print("\n请选择操作：")
    print("1. 预览重复记录（不删除）")
    print("2. 清理重复记录（执行删除）")
    
    choice = input("\n请输入选项 (1/2): ").strip()
    
    if choice == "1":
        preview_duplicates()
    elif choice == "2":
        success = clean_duplicate_aliases()
        sys.exit(0 if success else 1)
    else:
        print("无效的选项")
        sys.exit(1)
