
"""
更新产品使用次数统计
根据项目清单中的报价型号统计使用次数
"""
import sqlite3
import json
from collections import defaultdict

def load_project_lists(config_path="cable_config.json"):
    """从配置文件加载项目清单数据"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config.get("project_lists", {})
    except Exception as e:
        print(f"✗ 加载配置文件失败：{str(e)}")
        return {}

def count_model_usage(project_lists):
    """统计每个型号在不同项目中的使用次数"""
    # 使用字典记录每个型号在哪些项目中出现过
    model_projects = defaultdict(set)
    
    for project_code, project_data in project_lists.items():
        list_data = project_data.get("data", [])
        
        # 记录该项目中出现的所有型号（去重）
        project_models = set()
        for item in list_data:
            model = item.get("报价型号", "").strip()
            if model:
                project_models.add(model)
        
        # 将该项目添加到每个型号的项目集合中
        for model in project_models:
            model_projects[model].add(project_code)
    
    # 计算每个型号的使用次数（出现在多少个不同项目中）
    model_usage_count = {}
    for model, projects in model_projects.items():
        model_usage_count[model] = len(projects)
    
    return model_usage_count, model_projects

def find_matching_specs(cursor, model_name):
    """查找与型号匹配的产品规格"""
    matching_specs = []
    
    # 1. 直接匹配 product_model
    cursor.execute("""
        SELECT spec_id, product_model 
        FROM product_specs 
        WHERE product_model = ?
    """, (model_name,))
    
    result = cursor.fetchone()
    if result:
        matching_specs.append({
            'spec_id': result[0],
            'match_type': 'product_model',
            'match_value': result[1]
        })
    
    # 2. 匹配 alias_name
    cursor.execute("""
        SELECT DISTINCT asm.spec_id, asm.alias_name
        FROM alias_spec_mapping asm
        WHERE asm.alias_name = ?
    """, (model_name,))
    
    for row in cursor.fetchall():
        matching_specs.append({
            'spec_id': row[0],
            'match_type': 'alias',
            'match_value': row[1]
        })
    
    return matching_specs

def clear_all_usage_counts(db_path="cable_products_v4.db"):
    """清空所有产品的使用次数"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        cursor.execute("UPDATE product_specs SET usage_count = 0")
        conn.commit()
        
        affected = cursor.rowcount
        conn.close()
        
        return affected
    except Exception as e:
        print(f"✗ 清空使用次数失败：{str(e)}")
        return 0

def update_usage_counts(db_path="cable_products_v4.db", config_path="cable_config.json"):
    """更新产品使用次数"""
    
    print("=" * 80)
    print("更新产品使用次数统计")
    print("=" * 80)
    
    try:
        # 1. 加载项目清单数据
        print("\n1. 加载项目清单数据...")
        project_lists = load_project_lists(config_path)
        
        if not project_lists:
            print("✗ 没有找到项目清单数据")
            return False
        
        print(f"✓ 找到 {len(project_lists)} 个项目")
        
        # 2. 统计型号使用次数
        print("\n2. 统计型号使用次数...")
        model_usage_count, model_projects = count_model_usage(project_lists)
        
        print(f"✓ 统计到 {len(model_usage_count)} 个不同的型号")
        
        # 显示使用次数最多的前10个型号
        print("\n使用次数最多的型号（前10）：")
        sorted_models = sorted(model_usage_count.items(), key=lambda x: x[1], reverse=True)[:10]
        for i, (model, count) in enumerate(sorted_models, 1):
            projects = list(model_projects[model])[:3]
            projects_str = ", ".join(projects)
            if len(model_projects[model]) > 3:
                projects_str += f" (+{len(model_projects[model])-3})"
            print(f"   {i}. {model}: {count}次 (项目: {projects_str})")
        
        # 3. 连接数据库
        print("\n3. 连接数据库...")
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 4. 清空现有的使用次数
        print("\n4. 清空现有的使用次数...")
        cleared = clear_all_usage_counts(db_path)
        print(f"✓ 清空了 {cleared} 条记录的使用次数")
        
        # 重新连接（因为clear函数关闭了连接）
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 5. 更新使用次数
        print("\n5. 更新使用次数...")
        updated_specs = set()
        unmatched_models = []
        match_details = []
        
        for model, count in model_usage_count.items():
            matching_specs = find_matching_specs(cursor, model)
            
            if matching_specs:
                for spec_info in matching_specs:
                    spec_id = spec_info['spec_id']
                    
                    # 更新使用次数
                    cursor.execute("""
                        UPDATE product_specs 
                        SET usage_count = usage_count + ?
                        WHERE spec_id = ?
                    """, (count, spec_id))
                    
                    updated_specs.add(spec_id)
                    match_details.append({
                        'model': model,
                        'spec_id': spec_id,
                        'count': count,
                        'match_type': spec_info['match_type']
                    })
            else:
                unmatched_models.append((model, count))
        
        conn.commit()
        
        print(f"✓ 更新了 {len(updated_specs)} 个产品规格的使用次数")
        
        # 6. 显示匹配详情
        print("\n6. 匹配详情（前10条）：")
        for i, detail in enumerate(match_details[:10], 1):
            print(f"   {i}. 型号: {detail['model']}")
            print(f"      规格ID: {detail['spec_id']}")
            print(f"      使用次数: {detail['count']}")
            print(f"      匹配方式: {detail['match_type']}")
        
        if len(match_details) > 10:
            print(f"   ... 还有 {len(match_details) - 10} 条匹配记录")
        
        # 7. 显示未匹配的型号
        if unmatched_models:
            print(f"\n7. 未匹配到产品规格的型号（共{len(unmatched_models)}个）：")
            for i, (model, count) in enumerate(unmatched_models[:10], 1):
                print(f"   {i}. {model} (使用{count}次)")
            
            if len(unmatched_models) > 10:
                print(f"   ... 还有 {len(unmatched_models) - 10} 个未匹配型号")
        else:
            print("\n7. ✓ 所有型号都已匹配到产品规格")
        
        # 8. 验证结果
        print("\n8. 验证结果...")
        cursor.execute("""
            SELECT COUNT(*), SUM(usage_count), MAX(usage_count)
            FROM product_specs
            WHERE usage_count > 0
        """)
        
        result = cursor.fetchone()
        specs_with_usage = result[0]
        total_usage = result[1]
        max_usage = result[2]
        
        print(f"✓ 有使用记录的产品: {specs_with_usage}")
        print(f"✓ 总使用次数: {total_usage}")
        print(f"✓ 最大使用次数: {max_usage}")
        
        conn.close()
        
        print("\n" + "=" * 80)
        print("✓ 更新完成！")
        print("=" * 80)
        
        return True
        
    except Exception as e:
        print(f"\n✗ 更新失败：{str(e)}")
        import traceback
        traceback.print_exc()
        return False

def preview_usage_statistics(config_path="cable_config.json"):
    """预览使用次数统计（不更新数据库）"""
    
    print("=" * 80)
    print("预览使用次数统计")
    print("=" * 80)
    
    try:
        # 加载项目清单数据
        print("\n加载项目清单数据...")
        project_lists = load_project_lists(config_path)
        
        if not project_lists:
            print("✗ 没有找到项目清单数据")
            return
        
        print(f"✓ 找到 {len(project_lists)} 个项目")
        
        # 统计型号使用次数
        print("\n统计型号使用次数...")
        model_usage_count, model_projects = count_model_usage(project_lists)
        
        print(f"✓ 统计到 {len(model_usage_count)} 个不同的型号")
        
        # 显示详细统计
        print("\n" + "=" * 80)
        print("使用次数统计（按次数降序）")
        print("=" * 80)
        
        sorted_models = sorted(model_usage_count.items(), key=lambda x: x[1], reverse=True)
        
        for i, (model, count) in enumerate(sorted_models, 1):
            projects = list(model_projects[model])
            print(f"\n{i}. 型号: {model}")
            print(f"   使用次数: {count}")
            print(f"   出现在项目: {', '.join(projects)}")
        
        print("\n" + "=" * 80)
        print("统计完成")
        print("=" * 80)
        
    except Exception as e:
        print(f"\n✗ 预览失败：{str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    import sys
    
    print("\n请选择操作：")
    print("1. 预览使用次数统计（不更新数据库）")
    print("2. 更新产品使用次数（会清空现有数据并重新统计）")
    
    choice = input("\n请输入选项 (1/2): ").strip()
    
    if choice == "1":
        preview_usage_statistics()
        sys.exit(0)
    elif choice == "2":
        response = input("\n⚠ 此操作会清空所有产品的现有使用次数，确认继续？(y/n): ")
        if response.lower() == 'y':
            success = update_usage_counts()
            sys.exit(0 if success else 1)
        else:
            print("操作已取消")
            sys.exit(0)
    else:
        print("无效的选项")
        sys.exit(1)
