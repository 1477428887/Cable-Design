"""
验证关联型号清理结果
"""
import sqlite3

def verify_cleanup(db_path="cable_products_v4.db"):
    """验证清理后的数据库状态"""
    
    print("=" * 70)
    print("验证关联型号清理结果")
    print("=" * 70)
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 1. 检查是否还有重复记录
        print("\n1. 检查重复记录...")
        cursor.execute("""
            SELECT COUNT(*) 
            FROM product_specs ps
            JOIN alias_spec_mapping asm ON ps.spec_id = asm.spec_id
            WHERE ps.product_model = asm.alias_name
        """)
        
        duplicate_count = cursor.fetchone()[0]
        if duplicate_count == 0:
            print("   ✓ 没有发现重复记录")
        else:
            print(f"   ✗ 仍有 {duplicate_count} 条重复记录")
        
        # 2. 统计清理后的数据
        print("\n2. 数据库统计信息...")
        
        cursor.execute("SELECT COUNT(*) FROM product_specs")
        total_specs = cursor.fetchone()[0]
        print(f"   产品规格总数: {total_specs}")
        
        cursor.execute("SELECT COUNT(*) FROM alias_spec_mapping")
        total_mappings = cursor.fetchone()[0]
        print(f"   关联型号映射总数: {total_mappings}")
        
        cursor.execute("""
            SELECT COUNT(DISTINCT spec_id) 
            FROM alias_spec_mapping
        """)
        specs_with_aliases = cursor.fetchone()[0]
        print(f"   有关联型号的产品数: {specs_with_aliases}")
        
        # 3. 显示一些有关联型号的产品示例
        print("\n3. 有效关联型号示例（前5条）...")
        cursor.execute("""
            SELECT 
                ps.spec_id,
                ps.product_model,
                GROUP_CONCAT(asm.alias_name, ', ') as aliases
            FROM product_specs ps
            JOIN alias_spec_mapping asm ON ps.spec_id = asm.spec_id
            GROUP BY ps.spec_id, ps.product_model
            LIMIT 5
        """)
        
        examples = cursor.fetchall()
        if examples:
            for i, (spec_id, product_model, aliases) in enumerate(examples, 1):
                print(f"\n   {i}. 产品型号: {product_model}")
                print(f"      规格编号: {spec_id}")
                print(f"      关联型号: {aliases}")
        else:
            print("   没有找到有关联型号的产品")
        
        # 4. 检查数据完整性
        print("\n4. 数据完整性检查...")
        
        # 检查是否有孤立的映射记录
        cursor.execute("""
            SELECT COUNT(*) 
            FROM alias_spec_mapping asm
            LEFT JOIN product_specs ps ON asm.spec_id = ps.spec_id
            WHERE ps.spec_id IS NULL
        """)
        orphan_mappings = cursor.fetchone()[0]
        
        if orphan_mappings == 0:
            print("   ✓ 没有孤立的映射记录")
        else:
            print(f"   ⚠ 发现 {orphan_mappings} 条孤立的映射记录")
        
        conn.close()
        
        print("\n" + "=" * 70)
        if duplicate_count == 0 and orphan_mappings == 0:
            print("✓ 验证通过：数据库状态正常")
        else:
            print("⚠ 验证发现问题，请检查")
        print("=" * 70)
        
        return duplicate_count == 0 and orphan_mappings == 0
        
    except Exception as e:
        print(f"\n✗ 验证失败：{str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    import sys
    success = verify_cleanup()
    sys.exit(0 if success else 1)
