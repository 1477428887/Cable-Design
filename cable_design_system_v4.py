#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电缆设计结构库软件 V4.0 - 精简版（修复版）
作者：陈颖
功能：项目文件夹生成 + 基于参数卡片的电缆编码系统
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import pandas as pd
from datetime import datetime
import json
import sqlite3
import hashlib
import re

class CableProduct:
    """电缆产品参数卡片"""
    
    def __init__(self, **kwargs):
        # 10个核心字段
        self.category = kwargs.get('category', '')  # 产品大类
        self.voltage_rating = kwargs.get('voltage_rating', '')  # 额定电压
        self.conductor = kwargs.get('conductor', '')  # 导体材质与结构
        self.insulation = kwargs.get('insulation', '')  # 绝缘材料
        self.shield_type = kwargs.get('shield_type', 'None')  # 屏蔽类型
        self.inner_sheath = kwargs.get('inner_sheath', 'None')  # 内护套
        self.armor = kwargs.get('armor', 'None')  # 铠装类型
        self.outer_sheath = kwargs.get('outer_sheath', '')  # 外护套材料
        self.is_fire_resistant = kwargs.get('is_fire_resistant', False)  # 是否耐火
        self.special_performance = kwargs.get('special_performance', [])  # 特殊性能标签
        
        # 辅助字段
        self.model_name = kwargs.get('model_name', '')  # 型号名称（可选）
        self.description = kwargs.get('description', '')  # 描述
    
    def to_dict(self):
        """转换为字典"""
        return {
            'category': self.category,
            'voltage_rating': self.voltage_rating,
            'conductor': self.conductor,
            'insulation': self.insulation,
            'shield_type': self.shield_type,
            'inner_sheath': self.inner_sheath,
            'armor': self.armor,
            'outer_sheath': self.outer_sheath,
            'is_fire_resistant': self.is_fire_resistant,
            'special_performance': sorted(self.special_performance) if self.special_performance else [],
            'model_name': self.model_name,
            'description': self.description
        }
    
    def get_signature(self):
        """获取参数卡片的唯一签名（用于判断是否相同）"""
        # 只使用核心参数生成签名
        core_params = {
            'category': self.category,
            'voltage_rating': self.voltage_rating,
            'conductor': self.conductor,
            'insulation': self.insulation,
            'shield_type': self.shield_type,
            'inner_sheath': self.inner_sheath,
            'armor': self.armor,
            'outer_sheath': self.outer_sheath,
            'is_fire_resistant': self.is_fire_resistant,
            'special_performance': sorted(self.special_performance) if self.special_performance else []
        }
        
        # 生成MD5哈希
        signature_str = json.dumps(core_params, sort_keys=True, ensure_ascii=False)
        return hashlib.md5(signature_str.encode('utf-8')).hexdigest()
    
    def get_structure_string(self):
        """生成结构字符串"""
        parts = []
        
        # 导体
        if self.conductor:
            parts.append(self.conductor)
        
        # 绝缘（考虑耐火）- 修复：对于"无"值不添加到结构中
        if self.insulation and self.insulation not in ['None', '无']:
            if self.is_fire_resistant:
                parts.append(f"MT/{self.insulation}")
            else:
                parts.append(self.insulation)
        
        # 屏蔽
        if self.shield_type and self.shield_type not in ['None', '无']:
            parts.append(self.shield_type)
        
        # 内护套
        if self.inner_sheath and self.inner_sheath not in ['None', '无']:
            parts.append(self.inner_sheath)
        
        # 铠装
        if self.armor and self.armor not in ['None', '无']:
            parts.append(self.armor)
        
        # 外护套
        if self.outer_sheath and self.outer_sheath not in ['None', '无']:
            parts.append(self.outer_sheath)
        
        return "/".join(parts)
    
    def validate(self):
        """验证必填字段"""
        required_fields = [
            'category', 'voltage_rating', 'conductor', 
            'insulation', 'shield_type', 'armor', 'outer_sheath'
        ]
        
        missing = []
        for field in required_fields:
            value = getattr(self, field, '')
            # 对于外护套，"无"是有效值
            if field == 'outer_sheath':
                if not value or value == '':
                    missing.append(field)
            # 对于裸铜线，绝缘材料可以为"无"
            elif field == 'insulation' and self.category == '裸铜线':
                if not value or value == '':
                    missing.append(field)
            else:
                if not value or value == '':
                    missing.append(field)
        
        return len(missing) == 0, missing

class CableCodeManagerV4:
    """基于三层数据模型的电缆编码管理器 V4.1"""
    
    def __init__(self, db_path="cable_products_v4.db"):
        self.db_path = db_path
        self.init_database()
        self.init_model_parsing_rules()
    
    def init_database(self):
        """初始化三层数据模型数据库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 第1层：参数卡片表（Product Specification）- 唯一实体
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS product_specs (
                spec_id TEXT PRIMARY KEY,
                param_hash TEXT UNIQUE NOT NULL,
                category TEXT NOT NULL,
                voltage_rating TEXT NOT NULL,
                conductor TEXT NOT NULL,
                insulation TEXT NOT NULL,
                shield_type TEXT NOT NULL,
                inner_sheath TEXT,
                armor TEXT NOT NULL,
                outer_sheath TEXT NOT NULL,
                is_fire_resistant BOOLEAN NOT NULL,
                special_performance TEXT,
                product_model TEXT,
                structure_string TEXT,
                quota_path TEXT,
                spec_path TEXT,
                created_date TEXT,
                modified_date TEXT,
                usage_count INTEGER DEFAULT 0
            )
        ''')
        
        # 第2层：型号别名表（Model Alias）- 外部标识符
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS model_aliases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                alias_name TEXT NOT NULL,
                normalized_alias TEXT NOT NULL,
                created_date TEXT,
                usage_count INTEGER DEFAULT 0
            )
        ''')
        
        # 第3层：映射关系表（Alias ↔ Spec）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS alias_spec_mapping (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                alias_name TEXT NOT NULL,
                spec_id TEXT NOT NULL,
                source TEXT NOT NULL,
                confidence REAL NOT NULL,
                remarks TEXT,
                created_date TEXT,
                last_used_date TEXT,
                usage_count INTEGER DEFAULT 0,
                FOREIGN KEY (spec_id) REFERENCES product_specs (spec_id)
            )
        ''')
        
        # 创建索引
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_param_hash ON product_specs(param_hash)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_spec_category ON product_specs(category)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_alias_name ON model_aliases(alias_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_normalized_alias ON model_aliases(normalized_alias)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_mapping_alias ON alias_spec_mapping(alias_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_mapping_spec ON alias_spec_mapping(spec_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_confidence ON alias_spec_mapping(confidence)')
        
        # 数据库迁移：添加product_model字段（如果不存在）
        try:
            cursor.execute("ALTER TABLE product_specs ADD COLUMN product_model TEXT")
            print("✓ 数据库迁移：添加product_model字段")
        except sqlite3.OperationalError:
            # 字段已存在，忽略错误
            pass
        
        conn.commit()
        conn.close()
    
    def init_model_parsing_rules(self):
        """初始化型号解析规则库"""
        self.parsing_rules = {
            # 前缀规则
            "prefix_rules": [
                {"pattern": r"^NH[-.]", "params": {"is_fire_resistant": True}, "confidence": 0.95},
                {"pattern": r"^WDZ", "params": {"special_performance": ["低烟无卤"]}, "confidence": 0.9},
                {"pattern": r"^ZR[-.]", "params": {"special_performance": ["ZR"]}, "confidence": 0.9},
            ],
            
            # 后缀规则
            "suffix_rules": [
                {"pattern": r"\.LV$", "params": {"voltage_rating": "0.6/1kV", "shield_type": "无"}, "confidence": 0.95},
                {"pattern": r"\.MV$", "params": {"shield_type": "CTS"}, "confidence": 0.95},
                {"pattern": r"-1kV$", "params": {"voltage_rating": "0.6/1kV"}, "confidence": 0.9},
                {"pattern": r"-3kV$", "params": {"voltage_rating": "1.8/3kV"}, "confidence": 0.9},
            ],
            
            # 中间标记规则
            "token_rules": [
                {"token": "YJ", "params": {"insulation": "XLPE"}, "confidence": 0.95},
                {"token": "V", "params": {"outer_sheath": "PVC"}, "confidence": 0.9},
                {"token": "Y", "params": {"outer_sheath": "HDPE"}, "confidence": 0.9},
                {"token": "L", "params": {"conductor": "AL"}, "confidence": 0.95},
                {"token": "22", "params": {"armor": "STA"}, "confidence": 0.95},
                {"token": "23", "params": {"armor": "STA", "outer_sheath": "HDPE"}, "confidence": 0.95},
                {"token": "32", "params": {"armor": "SWA"}, "confidence": 0.95},
                {"token": "62", "params": {"armor": "SSTA"}, "confidence": 0.95},
                {"token": "72", "params": {"armor": "AWA"}, "confidence": 0.95},
                {"token": "P", "params": {"shield_type": "CWS"}, "confidence": 0.9},
                {"token": "P2", "params": {"shield_type": "CTS"}, "confidence": 0.9},
                {"token": "S", "params": {"shield_type": "CWS"}, "confidence": 0.85},
            ],
            
            # 完整型号规则
            "complete_rules": [
                {"pattern": r"^YJV\.LV$", "params": {"category": "低压", "voltage_rating": "0.6/1kV", "conductor": "CU", "insulation": "XLPE", "shield_type": "无", "armor": "无", "outer_sheath": "PVC"}, "confidence": 1.0},
                {"pattern": r"^YJLV\.LV$", "params": {"category": "低压", "voltage_rating": "0.6/1kV", "conductor": "AL", "insulation": "XLPE", "shield_type": "无", "armor": "无", "outer_sheath": "PVC"}, "confidence": 1.0},
                {"pattern": r"^YJV\.MV$", "params": {"category": "中压", "conductor": "CU", "insulation": "XLPE", "shield_type": "CTS", "armor": "无", "outer_sheath": "PVC"}, "confidence": 1.0},
                {"pattern": r"^H1Z2Z2-K$", "params": {"category": "光伏缆", "voltage_rating": "DC 1500V", "conductor": "TAC", "insulation": "XLPO", "shield_type": "无", "armor": "无", "outer_sheath": "XLPO"}, "confidence": 1.0},
                {"pattern": r"^PABC$", "params": {"category": "裸铜线", "voltage_rating": "N/A", "conductor": "PABC"}, "confidence": 1.0},
                {"pattern": r"^HDBC$", "params": {"category": "裸铜线", "voltage_rating": "N/A", "conductor": "HDBC"}, "confidence": 1.0},
            ]
        }
    
    def calculate_param_hash(self, product: 'CableProduct'):
        """计算参数卡片的标准化哈希值"""
        # 只使用核心参数生成哈希
        core_params = {
            'category': product.category,
            'voltage_rating': product.voltage_rating,
            'conductor': product.conductor,
            'insulation': product.insulation,
            'shield_type': product.shield_type,
            'inner_sheath': product.inner_sheath or '无',  # 统一使用"无"
            'armor': product.armor,
            'outer_sheath': product.outer_sheath,
            'is_fire_resistant': product.is_fire_resistant,
            'special_performance': sorted(product.special_performance) if product.special_performance else []
        }
        
        # 生成标准化JSON字符串
        param_str = json.dumps(core_params, sort_keys=True, ensure_ascii=False)
        return hashlib.sha256(param_str.encode('utf-8')).hexdigest()
    
    def find_or_create_spec(self, product: 'CableProduct', model_aliases=None):
        """查找或创建参数卡片，支持型号别名关联"""
        param_hash = self.calculate_param_hash(product)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 检查是否已存在相同参数的规格
        cursor.execute("SELECT spec_id FROM product_specs WHERE param_hash = ?", (param_hash,))
        existing = cursor.fetchone()
        
        if existing:
            spec_id = existing[0]
            # 更新使用次数
            cursor.execute("UPDATE product_specs SET usage_count = usage_count + 1, modified_date = ? WHERE spec_id = ?", 
                         (datetime.now().isoformat(), spec_id))
        else:
            # 创建新的参数卡片 - 使用哈希编码
            param_hash = self.calculate_param_hash(product)
            hash_suffix = param_hash[:12].upper()
            spec_id = f"CBL-SPEC-{hash_suffix}"
            
            # 插入新规格
            product_dict = product.to_dict()
            special_performance_json = json.dumps(product_dict['special_performance'], ensure_ascii=False)
            
            cursor.execute('''
                INSERT INTO product_specs (
                    spec_id, param_hash, category, voltage_rating, conductor, insulation,
                    shield_type, inner_sheath, armor, outer_sheath, is_fire_resistant,
                    special_performance, product_model, structure_string, created_date, modified_date, usage_count
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                spec_id, param_hash, product.category, product.voltage_rating, product.conductor,
                product.insulation, product.shield_type, product.inner_sheath, product.armor,
                product.outer_sheath, product.is_fire_resistant, special_performance_json,
                product.model_name, product.get_structure_string(), datetime.now().isoformat(), datetime.now().isoformat(), 1
            ))
        
        # 处理型号别名关联
        if model_aliases:
            for alias in model_aliases:
                if alias.strip():
                    self.add_alias_mapping(alias.strip(), spec_id, "手动录入", 1.0, cursor=cursor)
        
        conn.commit()
        conn.close()
        return spec_id
    
    def add_alias_mapping(self, alias_name, spec_id, source, confidence, remarks=None, cursor=None):
        """添加型号别名映射"""
        should_close = cursor is None
        if cursor is None:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
        
        normalized_alias = self.normalize_alias(alias_name)
        current_time = datetime.now().isoformat()
        
        # 检查别名是否已存在
        cursor.execute("SELECT id FROM model_aliases WHERE normalized_alias = ?", (normalized_alias,))
        if not cursor.fetchone():
            cursor.execute('''
                INSERT INTO model_aliases (alias_name, normalized_alias, created_date, usage_count)
                VALUES (?, ?, ?, ?)
            ''', (alias_name, normalized_alias, current_time, 0))
        
        # 检查映射是否已存在
        cursor.execute("SELECT id FROM alias_spec_mapping WHERE alias_name = ? AND spec_id = ?", 
                      (alias_name, spec_id))
        if not cursor.fetchone():
            cursor.execute('''
                INSERT INTO alias_spec_mapping (
                    alias_name, spec_id, source, confidence, remarks, 
                    created_date, last_used_date, usage_count
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (alias_name, spec_id, source, confidence, remarks, current_time, current_time, 1))
        else:
            # 更新现有映射
            cursor.execute('''
                UPDATE alias_spec_mapping 
                SET confidence = MAX(confidence, ?), last_used_date = ?, usage_count = usage_count + 1
                WHERE alias_name = ? AND spec_id = ?
            ''', (confidence, current_time, alias_name, spec_id))
        
        if should_close:
            conn.commit()
            conn.close()
    
    def normalize_alias(self, alias_name):
        """标准化型号别名"""
        # 转大写，移除多余空格和特殊字符
        normalized = alias_name.upper().strip()
        normalized = re.sub(r'[-_\s]+', '-', normalized)
        return normalized
    
    def parse_model_alias(self, alias_name):
        """智能解析型号别名，返回可能的参数组合"""
        candidates = []
        normalized_alias = self.normalize_alias(alias_name)
        
        # 1. 尝试完整规则匹配
        for rule in self.parsing_rules["complete_rules"]:
            if re.match(rule["pattern"], normalized_alias):
                candidate = rule["params"].copy()
                candidate["confidence"] = rule["confidence"]
                candidate["source"] = "完整规则匹配"
                candidates.append(candidate)
        
        if candidates:
            return candidates
        
        # 2. 组合规则解析
        base_params = {
            "category": "",
            "voltage_rating": "",
            "conductor": "CU",  # 默认铜导体
            "insulation": "",
            "shield_type": "无",
            "inner_sheath": "None",
            "armor": "无",
            "outer_sheath": "",
            "is_fire_resistant": False,
            "special_performance": []
        }
        
        confidence_sum = 0.0
        match_count = 0
        
        # 应用前缀规则
        for rule in self.parsing_rules["prefix_rules"]:
            if re.search(rule["pattern"], normalized_alias):
                base_params.update(rule["params"])
                confidence_sum += rule["confidence"]
                match_count += 1
        
        # 应用后缀规则
        for rule in self.parsing_rules["suffix_rules"]:
            if re.search(rule["pattern"], normalized_alias):
                base_params.update(rule["params"])
                confidence_sum += rule["confidence"]
                match_count += 1
        
        # 应用标记规则
        for rule in self.parsing_rules["token_rules"]:
            if rule["token"] in normalized_alias:
                base_params.update(rule["params"])
                confidence_sum += rule["confidence"]
                match_count += 1
        
        if match_count > 0:
            base_params["confidence"] = min(confidence_sum / match_count, 1.0)
            base_params["source"] = "规则组合解析"
            candidates.append(base_params)
        
        return candidates
    
    def search_by_alias(self, alias_name):
        """根据型号别名搜索参数卡片"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 1. 精确匹配已有映射
        cursor.execute('''
            SELECT asm.spec_id, asm.confidence, asm.source, asm.remarks, asm.usage_count,
                   ps.category, ps.voltage_rating, ps.conductor, ps.insulation, ps.shield_type,
                   ps.armor, ps.outer_sheath, ps.is_fire_resistant, ps.special_performance,
                   ps.structure_string, ps.product_model
            FROM alias_spec_mapping asm
            JOIN product_specs ps ON asm.spec_id = ps.spec_id
            WHERE asm.alias_name = ?
            ORDER BY asm.last_used_date DESC, asm.created_date DESC, asm.confidence DESC, asm.usage_count DESC
        ''', (alias_name,))
        
        results = []
        for row in cursor.fetchall():
            spec_id, confidence, source, remarks, usage_count = row[:5]
            spec_data = row[5:]
            results.append({
                "spec_id": spec_id,
                "confidence": confidence,
                "source": source,
                "remarks": remarks,
                "usage_count": usage_count,
                "spec_data": spec_data,
                "match_type": "确定"
            })
        
        # 2. 搜索产品型号字段（精确匹配）
        cursor.execute('''
            SELECT spec_id, category, voltage_rating, conductor, insulation, shield_type,
                   armor, outer_sheath, is_fire_resistant, special_performance, structure_string,
                   product_model
            FROM product_specs
            WHERE product_model = ?
            ORDER BY modified_date DESC, created_date DESC
        ''', (alias_name,))
        
        for row in cursor.fetchall():
            spec_id = row[0]
            spec_data = row[1:]  # Include all fields including product_model
            
            # 检查是否已经在结果中（避免重复）
            if not any(r["spec_id"] == spec_id for r in results):
                # 使用智能置信度计算
                confidence = self.calculate_alias_confidence(alias_name, row[11])  # row[11] is product_model
                results.append({
                    "spec_id": spec_id,
                    "confidence": confidence,
                    "source": "产品型号匹配",
                    "remarks": f"产品型号: {row[11]}",  # product_model is at index 11
                    "usage_count": 0,
                    "spec_data": spec_data,
                    "match_type": "确定"
                })
        
        # 3. 如果没有精确匹配，尝试部分匹配（别名和产品型号）
        if not results:
            # 3a. 搜索别名部分匹配
            cursor.execute('''
                SELECT asm.spec_id, asm.alias_name, asm.confidence, asm.source, asm.remarks, asm.usage_count,
                       ps.category, ps.voltage_rating, ps.conductor, ps.insulation, ps.shield_type,
                       ps.armor, ps.outer_sheath, ps.is_fire_resistant, ps.special_performance,
                       ps.structure_string, ps.product_model
                FROM alias_spec_mapping asm
                JOIN product_specs ps ON asm.spec_id = ps.spec_id
                ORDER BY asm.last_used_date DESC, asm.created_date DESC, asm.confidence DESC, asm.usage_count DESC
            ''')
            
            for row in cursor.fetchall():
                spec_id, alias_name_db, confidence, source, remarks, usage_count = row[:6]
                spec_data = row[6:]
                
                # 使用独立型号匹配检查
                if self.is_independent_model_match(alias_name, alias_name_db):
                    # 检查是否已经在结果中（避免重复）
                    if not any(r["spec_id"] == spec_id for r in results):
                        # 使用智能置信度计算
                        calculated_confidence = self.calculate_alias_confidence(alias_name, alias_name_db)
                        results.append({
                            "spec_id": spec_id,
                            "confidence": calculated_confidence,
                            "source": f"别名部分匹配: {alias_name_db}",
                            "remarks": remarks or f"匹配别名: {alias_name_db}",
                            "usage_count": usage_count,
                            "spec_data": spec_data,
                            "match_type": "确定"
                        })
            
            # 3b. 搜索产品型号部分匹配
            cursor.execute('''
                SELECT spec_id, category, voltage_rating, conductor, insulation, shield_type,
                       armor, outer_sheath, is_fire_resistant, special_performance, structure_string,
                       product_model
                FROM product_specs
                WHERE product_model IS NOT NULL AND product_model != ''
                ORDER BY modified_date DESC, created_date DESC
            ''')
            
            for row in cursor.fetchall():
                spec_id = row[0]
                spec_data = row[1:]  # Include all fields including product_model
                
                # 使用独立型号匹配检查
                if self.is_independent_model_match(alias_name, row[11]):  # product_model is at index 11
                    # 检查是否已经在结果中（避免重复）
                    if not any(r["spec_id"] == spec_id for r in results):
                        # 使用智能置信度计算
                        confidence = self.calculate_alias_confidence(alias_name, row[11])
                        results.append({
                            "spec_id": spec_id,
                            "confidence": confidence,
                            "source": f"产品型号部分匹配: {row[11]}",
                            "remarks": f"产品型号: {row[11]}",
                            "usage_count": 0,
                            "spec_data": spec_data,
                            "match_type": "确定"
                        })
        
        # 4. 如果没有任何匹配，尝试智能解析
        if not results:
            candidates = self.parse_model_alias(alias_name)
            for candidate in candidates:
                results.append({
                    "spec_id": None,
                    "confidence": candidate.get("confidence", 0.5),
                    "source": candidate.get("source", "智能解析"),
                    "remarks": "需要用户确认",
                    "usage_count": 0,
                    "spec_data": None,
                    "candidate_params": candidate,
                    "match_type": "推测"
                })
        
        conn.close()
        
        # 按置信度降序排序，然后按使用次数降序排序
        results.sort(key=lambda x: (x.get("confidence", 0), x.get("usage_count", 0)), reverse=True)
        
        return results
    
    def search_by_structure(self, structure_query):
        """根据结构字符串搜索参数卡片"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 解析结构查询
        structure_parts = [part.strip() for part in structure_query.split('/') if part.strip()]
        
        # 构建SQL查询条件
        conditions = []
        params = []
        
        for part in structure_parts:
            part_upper = part.upper()
            # 根据常见材料代码匹配
            if part_upper in ['CU', 'AL', 'TAC', 'PABC', 'HDBC']:
                conditions.append("conductor = ?")
                params.append(part_upper)
            elif part_upper in ['XLPE', 'PVC', 'XLPO', 'LSZH']:
                conditions.append("(insulation = ? OR outer_sheath = ?)")
                params.extend([part_upper, part_upper])
            elif part_upper in ['CTS', 'CWS', 'CWB']:
                conditions.append("shield_type = ?")
                params.append(part_upper)
            elif part_upper in ['STA', 'SWA', 'SSTA', 'AWA', 'ATA']:
                conditions.append("armor = ?")
                params.append(part_upper)
        
        if conditions:
            where_clause = " AND ".join(conditions)
            cursor.execute(f'''
                SELECT spec_id, category, voltage_rating, conductor, insulation, shield_type,
                       armor, outer_sheath, is_fire_resistant, special_performance, structure_string,
                       product_model, usage_count
                FROM product_specs
                WHERE {where_clause}
                ORDER BY usage_count DESC
            ''', params)
        else:
            # 如果没有识别到结构组件，进行模糊搜索
            cursor.execute('''
                SELECT spec_id, category, voltage_rating, conductor, insulation, shield_type,
                       armor, outer_sheath, is_fire_resistant, special_performance, structure_string,
                       product_model, usage_count
                FROM product_specs
                WHERE structure_string LIKE ?
                ORDER BY usage_count DESC
            ''', (f'%{structure_query}%',))
        
        results = []
        for row in cursor.fetchall():
            spec_id = row[0]
            spec_data = row[1:]
            
            # 获取该规格的所有型号别名 - 优先显示最近更新的别名
            cursor.execute('''
                SELECT alias_name, confidence, source
                FROM alias_spec_mapping
                WHERE spec_id = ?
                ORDER BY last_used_date DESC, created_date DESC, confidence DESC, usage_count DESC
            ''', (spec_id,))
            
            aliases = cursor.fetchall()
            
            # 计算结构匹配的置信度
            structure_string = spec_data[9] if len(spec_data) > 9 else ""  # structure_string在索引9
            matched_components = [spec_data[2], spec_data[3], spec_data[4], spec_data[5], spec_data[6]]  # conductor, insulation, shield_type, armor, outer_sheath
            confidence = self.calculate_structure_confidence(structure_query, structure_string, matched_components)
            
            results.append({
                "spec_id": spec_id,
                "spec_data": spec_data,
                "aliases": aliases,
                "confidence": confidence,
                "source": "结构匹配",
                "match_type": "确定"
            })
        
        conn.close()
        
        # 按置信度降序排序，然后按使用次数降序排序
        def sort_key(x):
            confidence = x.get("confidence", 0)
            spec_data = x.get("spec_data", [])
            # usage_count 应该在 spec_data 的最后一个位置，但要确保它是数字
            usage_count = 0
            if spec_data and len(spec_data) > 0:
                last_item = spec_data[-1]
                if isinstance(last_item, (int, float)):
                    usage_count = last_item
            return (confidence, usage_count)
        
        results.sort(key=sort_key, reverse=True)
        
        return results
    
    def get_all_specs(self):
        """获取所有参数卡片"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT spec_id, category, voltage_rating, conductor, insulation, shield_type,
                   inner_sheath, armor, outer_sheath, is_fire_resistant, special_performance,
                   product_model, structure_string, quota_path, spec_path, created_date, modified_date, usage_count
            FROM product_specs ORDER BY modified_date DESC, created_date DESC, usage_count DESC
        ''')
        
        specs = cursor.fetchall()
        
        # 为每个规格获取关联的型号别名
        results = []
        for spec in specs:
            spec_id = spec[0]
            cursor.execute('''
                SELECT alias_name, confidence, source, usage_count
                FROM alias_spec_mapping
                WHERE spec_id = ?
                ORDER BY last_used_date DESC, created_date DESC, confidence DESC, usage_count DESC
            ''', (spec_id,))
            
            aliases = cursor.fetchall()
            results.append({
                "spec": spec,
                "aliases": aliases
            })
        
        conn.close()
        return results
    
    def get_spec_by_id(self, spec_id):
        """根据规格ID获取详细信息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT spec_id, category, voltage_rating, conductor, insulation, shield_type,
                   inner_sheath, armor, outer_sheath, is_fire_resistant, special_performance,
                   product_model, structure_string, quota_path, spec_path, created_date, modified_date, usage_count
            FROM product_specs WHERE spec_id = ?
        ''', (spec_id,))
        
        spec = cursor.fetchone()
        if not spec:
            conn.close()
            return None
        
        # 获取关联的型号别名 - 优先显示最近更新的别名
        cursor.execute('''
            SELECT alias_name, confidence, source, usage_count, remarks
            FROM alias_spec_mapping
            WHERE spec_id = ?
            ORDER BY last_used_date DESC, created_date DESC, confidence DESC, usage_count DESC
        ''', (spec_id,))
        
        aliases = cursor.fetchall()
        
        conn.close()
        return {
            "spec": spec,
            "aliases": aliases
        }
    
    def update_spec_paths(self, spec_id, quota_path=None, spec_path=None):
        """更新规格的路径信息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        updates = []
        values = []
        
        if quota_path is not None:
            updates.append("quota_path = ?")
            values.append(quota_path)
        
        if spec_path is not None:
            updates.append("spec_path = ?")
            values.append(spec_path)
        
        if updates:
            updates.append("modified_date = ?")
            values.append(datetime.now().isoformat())
            values.append(spec_id)
            
            query = f"UPDATE product_specs SET {', '.join(updates)} WHERE spec_id = ?"
            cursor.execute(query, values)
            conn.commit()
        
        conn.close()
    
    def record_alias_usage(self, alias_name, spec_id):
        """记录型号别名的使用"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        current_time = datetime.now().isoformat()
        
        # 更新别名使用次数
        cursor.execute('''
            UPDATE model_aliases 
            SET usage_count = usage_count + 1 
            WHERE alias_name = ?
        ''', (alias_name,))
        
        # 更新映射使用次数和最后使用时间
        cursor.execute('''
            UPDATE alias_spec_mapping 
            SET usage_count = usage_count + 1, last_used_date = ?
            WHERE alias_name = ? AND spec_id = ?
        ''', (current_time, alias_name, spec_id))
        
        # 更新规格使用次数
        cursor.execute('''
            UPDATE product_specs 
            SET usage_count = usage_count + 1, modified_date = ?
            WHERE spec_id = ?
        ''', (current_time, spec_id))
        
        conn.commit()
        conn.close()

    def is_independent_model_match(self, search_text, target):
        """检查搜索文本是否作为独立的型号部分出现在目标字符串中"""
        # 定义分隔符
        separators = ['.', '-', '_', ' ']
        
        # 将目标字符串按分隔符分割成部分
        parts = [target]
        for sep in separators:
            new_parts = []
            for part in parts:
                new_parts.extend(part.split(sep))
            parts = new_parts
        
        # 检查搜索文本是否作为独立部分存在
        for part in parts:
            if part.strip() == search_text:
                return True
        
        return False
    
    def calculate_alias_confidence(self, search_query, matched_model):
        """计算型号别名匹配的置信度"""
        search_query = search_query.upper().strip()
        matched_model = matched_model.upper().strip()
        
        # 1. 完全精确匹配 = 100%
        if search_query == matched_model:
            return 1.0
        
        # 2. 移除常见前缀进行比较
        prefixes = ['ZR.', 'ZC.', 'ZA.', 'ZB.', 'NH.', 'WDZ.', 'WDZR.', 'WDZC.', 'WDZN.', 'ZCN.']
        
        search_clean = search_query
        matched_clean = matched_model
        
        # 记录是否有前缀被移除
        search_had_prefix = False
        matched_had_prefix = False
        
        for prefix in prefixes:
            if search_clean.startswith(prefix):
                search_clean = search_clean[len(prefix):]
                search_had_prefix = True
            if matched_clean.startswith(prefix):
                matched_clean = matched_clean[len(prefix):]
                matched_had_prefix = True
        
        # 3. 去除前缀后的精确匹配 = 100%
        if search_clean == matched_clean:
            return 1.0
        
        # 4. 核心型号匹配（如 YJV22 匹配 YJV22.LV）
        if (matched_clean.startswith(search_clean + '.') or 
            matched_clean.startswith(search_clean + '_') or
            search_clean.startswith(matched_clean + '.') or
            search_clean.startswith(matched_clean + '_')):
            
            # 如果匹配的型号有前缀而搜索没有，稍微降低置信度
            if matched_had_prefix and not search_had_prefix:
                return 0.90  # 有前缀的核心匹配
            else:
                return 1.0   # 无前缀的核心匹配
        
        # 5. 部分匹配但有差异（如YJV22 vs YJV23）= 75%
        if len(search_clean) >= 3 and len(matched_clean) >= 3:
            # 计算公共前缀长度
            common_prefix = 0
            min_len = min(len(search_clean), len(matched_clean))
            for i in range(min_len):
                if search_clean[i] == matched_clean[i]:
                    common_prefix += 1
                else:
                    break
            
            # 如果公共前缀至少3个字符，且占搜索词的大部分
            if common_prefix >= 3:
                search_similarity = common_prefix / len(search_clean)  # 基于搜索词长度
                if search_similarity >= 0.6:  # 降低阈值
                    return 0.75
        
        # 6. 弱匹配 = 60%
        if search_clean in matched_clean or matched_clean in search_clean:
            return 0.6
        
        # 7. 默认低置信度 = 50%
        return 0.5
    
    def calculate_structure_confidence(self, search_query, matched_structure, matched_components):
        """计算结构搜索的置信度"""
        search_parts = [part.strip().upper() for part in search_query.split('/') if part.strip()]
        matched_parts = [part.strip().upper() for part in matched_structure.split('/') if part.strip()]
        
        if not search_parts:
            return 0.5
        
        # 1. 完全匹配 = 100%
        if set(search_parts) == set(matched_parts):
            return 1.0
        
        # 2. 搜索的所有组件都在匹配结果中 = 95%
        if all(part in matched_parts for part in search_parts):
            return 0.95
        
        # 3. 大部分组件匹配 = 85%
        matched_count = sum(1 for part in search_parts if part in matched_parts)
        match_ratio = matched_count / len(search_parts)
        
        if match_ratio >= 0.8:
            return 0.85
        elif match_ratio >= 0.6:
            return 0.75
        elif match_ratio >= 0.4:
            return 0.65
        elif match_ratio > 0:
            return 0.55
        else:
            return 0.5
    
    def get_spec_aliases(self, spec_id):
        """获取规格的所有别名"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                SELECT alias_name, confidence, source
                FROM alias_spec_mapping
                WHERE spec_id = ?
                ORDER BY last_used_date DESC, created_date DESC, confidence DESC, usage_count DESC
            ''', (spec_id,))
            
            aliases = cursor.fetchall()
            return aliases
        except Exception as e:
            print(f"获取别名失败: {str(e)}")
            return []
        finally:
            conn.close()

class CableDesignSystemV4:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("电缆设计结构库软件 V4.0 - 陈颖")
        self.root.geometry("1300x850") # 稍微调大默认窗口以适应更多信息
        self.root.minsize(1200, 700)  # 设置最小窗口尺寸，确保界面元素可见
        self.root.configure(bg="#f5f6fa") # 改为更现代的浅灰蓝色背景
        
        # 应用原生的现代样式配置
        self.apply_modern_styling()
        
        # 配置文件路径
        self.config_file = "cable_config.json"
        self.load_config()
        
        # 初始化编码管理器
        self.code_manager = CableCodeManagerV4()
        
        # 电缆参数数据
        self.init_cable_data()
        
        # 创建主界面
        self.create_main_interface()

    def apply_modern_styling(self):
        """应用现代化的原生 Tkinter 样式"""
        style = ttk.Style()
        
        # 尝试使用 clam 或 winnative 主题作为基础以获得更好的跨平台美观度
        if 'clam' in style.theme_names():
            style.theme_use('clam')
            
        # 1. 设置基础字体 (微软雅黑, 10号)
        default_font = ('Microsoft YaHei', 10)
        self.root.option_add('*Font', default_font)
        
        # 2. 配置各种 ttk 组件的默认样式
        style.configure('.', font=default_font)
        style.configure('TFrame', background='#f5f6fa')
        style.configure('TLabelframe', background='#f5f6fa')
        style.configure('TLabelframe.Label', font=('Microsoft YaHei', 10, 'bold'), background='#f5f6fa', foreground='#2f3640')
        
        style.configure('TLabel', background='#f5f6fa', foreground='#2f3640')
        style.configure('TButton', font=('Microsoft YaHei', 10), padding=(10, 5))
        style.configure('TRadiobutton', background='#f5f6fa')
        style.configure('TCheckbutton', background='#f5f6fa')
        
        # 3. 创建 Accent 高亮按钮样式
        style.configure('Accent.TButton', font=('Microsoft YaHei', 10, 'bold'), foreground='#ffffff', background='#00a8ff')
        style.map('Accent.TButton',
            background=[('active', '#0097e6'), ('disabled', '#dcdde1')],
            foreground=[('disabled', '#a4b0be')]
        )
        
        # 4. 优化 Treeview (表格) 样式
        style.configure('Treeview', 
            font=('Microsoft YaHei', 10), 
            rowheight=28, 
            background='#ffffff',
            fieldbackground='#ffffff',
            foreground='#2f3640'
        )
        style.configure('Treeview.Heading', font=('Microsoft YaHei', 10, 'bold'), background='#dcdde1', foreground='#2f3640', padding=5)
        style.map('Treeview', 
            background=[('selected', '#00a8ff')],
            foreground=[('selected', '#ffffff')]
        )
        
        # 5. Notebook 标签页优化
        style.configure('TNotebook', background='#f5f6fa', tabmargins=[2, 5, 2, 0])
        style.configure('TNotebook.Tab', font=('Microsoft YaHei', 10), padding=[15, 5], background='#dcdde1')
        style.map('TNotebook.Tab',
            expand=[('selected', [1, 1, 1, 0])],
            background=[('selected', '#ffffff')],
            foreground=[('selected', '#00a8ff')]
        )
        
    def load_config(self):
        """加载配置文件"""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            self.config = {
                "default_folder": os.getcwd(),
                "quota_folder": "",
                "spec_folder": ""
            }
            self.save_config()
    
    def save_config(self):
        """保存配置文件"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)
    
    def init_cable_data(self):
        """初始化电缆参数数据"""
        # 电压等级分类（修正版）
        self.voltage_levels = {
            "低压": ["0.6/1kV", "1.8/3kV"],
            "中压": ["3.6/6kV", "6/10kV", "6.35/11kV", "8.7/15kV", "12/20kV", 
                   "12.7/22kV", "18/30kV", "19/33kV", "26/35kV"],
            "布线": ["300/500V", "450/750V", "0.6/1kV"],
            "光伏缆": ["DC 1500V"],
            "控缆和仪表缆": ["450/750V"],
            "裸铜线": ["N/A"],
            "橡套电缆": ["300/500V", "450/750V", "0.6/1kV", "1.8/3kV", "3.6/6kV", "6/10kV", 
                      "6.35/11kV", "8.7/15kV", "12/20kV", "12.7/22kV", "18/30kV", "19/33kV", "26/35kV"]
        }
        
        # 电缆结构组件选项（精简版）
        self.conductors = {
            "CU": "铜",
            "AL": "铝", 
            "AAAC": "铝合金",
            "Cl1(CU)": "第1类绞合铜导体",
            "Cl5(CU)": "第5类绞合铜导体",
            "TAC": "退火铜合金导体",
            "PABC": "退火裸铜绞线",
            "HDBC": "硬拉裸铜线"
        }
        
        self.insulations = {
            "无": "无绝缘",
            "XLPE": "交联聚乙烯",
            "PVC": "聚氯乙烯",
            "XLPO": "交联聚烯烃",
            "LSZH": "低烟无卤",
            "EPR": "乙丙橡胶"  # 新增：橡套电缆绝缘材料，中压、低压电缆也可能用到
        }
        
        self.shields = {
            "无": "无屏蔽",
            "CTS": "铜带屏蔽",
            "CWS": "铜丝疏绕屏蔽",
            "CWB": "铜丝编织屏蔽",
            "AL-PET": "铝塑复合带屏蔽"
        }
        
        self.armors = {
            "无": "无铠装",
            "STA": "镀锌钢带铠装",
            "SWA": "镀锌钢丝铠装",
            "SSWA": "不锈钢丝铠装",
            "SSTA": "不锈钢带铠装",
            "AWA": "铝合金丝铠装",
            "ATA": "铝合金带铠装"
        }
        
        self.sheaths = {
            "PVC": "聚氯乙烯护套",
            "HDPE": "高密度聚乙烯护套",
            "LSZH": "低烟无卤护套",
            "SE4": "氯丁橡胶护套"  # 新增：仅橡套电缆使用的护套材料
        }
        
        # 特殊性能选项（修正版）
        self.fire_resistant_options = ["ZR", "ZC", "ZB", "ZA"]  # 阻燃等级，互斥
        self.other_properties = ["防鼠", "防白蚁", "耐油", "耐酸碱"]  # 其他性能，可多选
        
        # 电缆型号映射表
        self.cable_type_mapping = self.build_cable_type_mapping()

    def build_cable_type_mapping(self):
        """构建电缆型号映射表"""
        mapping = {
            # 低压电缆 (LV)
            "CU/XLPE/PVC": "YJV.LV",
            "CU/XLPE/PVC/STA/PVC": "YJV22.LV", 
            "CU/XLPE/PVC/SWA/PVC": "YJV32.LV",
            "CU/XLPE/PVC/SSTA/PVC": "YJV62.LV",
            "CU/XLPE/PVC/AWA/PVC": "YJV72.LV",
            "CU/XLPE/CTS/PVC/STA/PVC": "YJVP2.22.LV",
            "CU/XLPE/CTS/PVC/SWA/PVC": "YJVP2.32.LV",
            "CU/XLPE/CTS/PVC/SSTA/PVC": "YJVP2.62.LV",
            "CU/XLPE/CTS/PVC/AWA/PVC": "YJVP2.72.LV",
            "CU/XLPE/CWS/PVC": "YJVP.LV",
            "CU/XLPE/CTS/PVC": "YJVP2.LV",
            "Cl5(CU)/XLPE/PVC": "YJVR.LV",
            "CU/XLPE/HDPE": "YJY.LV",
            "CU/XLPE/HDPE/STA/HDPE": "YJY23.LV",
            "AL/XLPE/PVC": "YJLV.LV",
            "AL/XLPE/PVC/STA/PVC": "YJLV22.LV",
            "AL/XLPE/PVC/STA/HDPE": "YJLV23.LV",
            "AL/XLPE/PVC/SWA/PVC": "YJLV32.LV",
            "AL/XLPE/PVC/SSTA/PVC": "YJLV62.LV",
            "AL/XLPE/PVC/AWA/PVC": "YJLV72.LV",
            
            # 中压电缆 (MV)
            "CU/XLPE/CTS/PVC": "YJV.MV",
            "CU/XLPE/CTS/PVC/STA/PVC": "YJV22.MV",
            "CU/XLPE/CTS/PVC/SWA/PVC": "YJV32.MV",
            "CU/XLPE/CTS/PVC/SSTA/PVC": "YJV62.MV",
            "CU/XLPE/CTS/PVC/ATA/PVC": "YJV62.MV",
            "CU/XLPE/CTS/PVC/AWA/PVC": "YJV72.MV",
            "CU/XLPE/CTS/PVC/SSWA/PVC": "YJV72.MV",
            "AL/XLPE/CTS/PVC": "YJLV.MV",
            "AL/XLPE/CTS/PVC/STA/PVC": "YJLV22.MV",
            "AL/XLPE/CTS/PVC/SWA/PVC": "YJLV32.MV",
            "AL/XLPE/CTS/PVC/SSTA/PVC": "YJLV62.MV",
            "AL/XLPE/CTS/PVC/ATA/PVC": "YJLV62.MV",
            "AL/XLPE/CTS/PVC/AWA/PVC": "YJLV72.MV",
            "CU/XLPE/CTS/HDPE": "YJY.MV",
            "CU/XLPE/CTS/HDPE/STA/HDPE": "YJY23.MV",
            "CU/XLPE/CTS/HDPE/AWA/HDPE": "YJY73.MV",
            
            # 布线电缆
            "Cl5(CU)/PVC/PVC": "RVV",
            "CU/PVC": "BV",
            "CU/PVC": "BVR",
            
            # 控制电缆
            "CU/PVC/PVC": "KVV",
            "CU/PVC/CWB/PVC": "KVVP",
            "CU/PVC/CTS/PVC": "KVVP2",
            "Cl5(CU)/PVC/PVC": "KVVR",
            "Cl5(CU)/PVC/CWB/PVC": "KVVRP",
            
            # 光伏电缆
            "TAC/XLPO/XLPO": "H1Z2Z2.K",
            
            # 裸铜线
            "PABC": "PABC",
            "HDBC": "HDBC"
        }
        return mapping

    def create_main_interface(self):
        """创建主界面"""
        # 创建笔记本控件
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # 第一个标签页：项目生成
        self.project_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.project_frame, text="📁 项目管理")
        self.create_project_interface()
        
        # 第二个标签页：项目清单列表界面
        self.project_list_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.project_list_frame, text="📋 项目清单列表")
        self.create_project_list_interface()
        
        # 第三个标签页：参数卡片编辑
        self.card_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.card_frame, text="📋 参数卡片编辑")
        self.create_parameter_card_interface()
        
        # 第四个标签页：产品编码管理
        self.management_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.management_frame, text="📊 产品编码管理")
        self.create_product_management_interface()
        
        # 第五个标签页：智能清单解析
        self.intelligent_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.intelligent_frame, text="🤖 智能清单解析")
        self.create_intelligent_parser_interface()
        
        # 第六个标签页：设置
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="系统设置")
        self.create_settings_interface()

    def create_project_interface(self):
        """创建项目生成界面"""
        # 标题
        title_label = tk.Label(self.project_frame, text="项目文件夹自动生成", 
                              font=("Microsoft YaHei", 16, "bold"), bg="#f0f0f0")
        title_label.pack(pady=15)

        # 创建主容器，使用水平布局
        main_container = tk.Frame(self.project_frame, bg="#f0f0f0")
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=15)

        # 左侧：项目信息输入
        left_frame = ttk.LabelFrame(main_container, text="项目信息输入", padding=25)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        # 项目编号
        tk.Label(left_frame, text="项目编号:", font=("Microsoft YaHei", 11)).grid(row=0, column=0, sticky=tk.W, pady=8)
        self.project_code_var = tk.StringVar(value="2026888888-LL")
        tk.Entry(left_frame, textvariable=self.project_code_var, width=25, font=("Microsoft YaHei", 10)).grid(row=0, column=1, pady=8, padx=15)
        
        # 项目名称
        tk.Label(left_frame, text="项目名称:", font=("Microsoft YaHei", 11)).grid(row=1, column=0, sticky=tk.W, pady=8)
        self.project_name_var = tk.StringVar(value="项目名称")
        tk.Entry(left_frame, textvariable=self.project_name_var, width=40, font=("Microsoft YaHei", 10)).grid(row=1, column=1, pady=8, padx=15)
        
        # 业务员
        tk.Label(left_frame, text="业务员:", font=("Microsoft YaHei", 11)).grid(row=2, column=0, sticky=tk.W, pady=8)
        self.project_manager_var = tk.StringVar(value="业务员")
        tk.Entry(left_frame, textvariable=self.project_manager_var, width=25, font=("Microsoft YaHei", 10)).grid(row=2, column=1, pady=8, padx=15)
        
        # 保存路径
        tk.Label(left_frame, text="保存路径:", font=("Microsoft YaHei", 11)).grid(row=3, column=0, sticky=tk.W, pady=8)
        path_frame = tk.Frame(left_frame)
        path_frame.grid(row=3, column=1, pady=8, padx=15, sticky=tk.W)
        
        self.save_path_var = tk.StringVar(value=self.config.get("default_folder", os.getcwd()))
        tk.Entry(path_frame, textvariable=self.save_path_var, width=35, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        tk.Button(path_frame, text="浏览", command=self.browse_folder, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=8)
        
        # 生成按钮
        generate_btn = tk.Button(left_frame, text="生成项目文件夹", 
                               command=self.generate_project_folders,
                               bg="#4CAF50", fg="white", font=("Microsoft YaHei", 12, "bold"),
                               height=2, width=18)
        generate_btn.grid(row=4, column=0, columnspan=2, pady=25)

        # 右侧：近期项目列表
        right_frame = ttk.LabelFrame(main_container, text="近期项目列表", padding=15)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 近期项目列表
        self.create_recent_projects_list(right_frame)

    def create_recent_projects_list(self, parent):
        """创建近期项目列表"""
        # 筛选框架
        filter_frame = ttk.LabelFrame(parent, text="项目筛选", padding=10)
        filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 第一行筛选
        filter_row1 = tk.Frame(filter_frame)
        filter_row1.pack(fill=tk.X, pady=5)
        
        # 月份筛选（新增）
        tk.Label(filter_row1, text="月份:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_month_var = tk.StringVar(value="全部")
        self.month_filter_combo = ttk.Combobox(filter_row1, textvariable=self.filter_month_var,
                                              state="readonly", width=12)
        self.month_filter_combo.pack(side=tk.LEFT, padx=(5, 15))
        self.month_filter_combo.bind('<<ComboboxSelected>>', self.on_recent_month_filter_change)
        
        # 业务员筛选
        tk.Label(filter_row1, text="业务员:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_manager_var = tk.StringVar(value="全部")
        self.manager_filter_combo = ttk.Combobox(filter_row1, textvariable=self.filter_manager_var,
                                               state="readonly", width=12)
        self.manager_filter_combo.pack(side=tk.LEFT, padx=(5, 15))
        self.manager_filter_combo.bind('<<ComboboxSelected>>', self.apply_project_filters)
        
        # 时间筛选
        tk.Label(filter_row1, text="时间范围:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_time_var = tk.StringVar(value="全部")
        time_filter_combo = ttk.Combobox(filter_row1, textvariable=self.filter_time_var,
                                       values=["全部", "今天", "本周", "本月", "最近3个月", "最近6个月", "今年"],
                                       state="readonly", width=12)
        time_filter_combo.pack(side=tk.LEFT, padx=(5, 15))
        time_filter_combo.bind('<<ComboboxSelected>>', self.apply_project_filters)
        
        # 清空筛选按钮
        tk.Button(filter_row1, text="清空筛选", command=self.clear_project_filters,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 9), width=8).pack(side=tk.LEFT, padx=5)
        
        # 统计信息标签（新增）
        self.recent_projects_stats_label = tk.Label(filter_frame, text="", 
                                                    font=("Microsoft YaHei", 10), fg="#666")
        self.recent_projects_stats_label.pack(pady=5)
        
        # 列表框架
        list_frame = tk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 创建Treeview显示近期项目
        columns = ("项目编号", "项目名称", "业务员", "型号数量", "技术规范数量", "创建时间")
        self.recent_projects_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=12)
        
        # 设置列标题和宽度
        self.recent_projects_tree.heading("项目编号", text="项目编号")
        self.recent_projects_tree.heading("项目名称", text="项目名称")
        self.recent_projects_tree.heading("业务员", text="业务员")
        self.recent_projects_tree.heading("型号数量", text="型号数量")
        self.recent_projects_tree.heading("技术规范数量", text="技术规范数量")
        self.recent_projects_tree.heading("创建时间", text="创建时间")
        
        self.recent_projects_tree.column("项目编号", width=100)
        self.recent_projects_tree.column("项目名称", width=150)
        self.recent_projects_tree.column("业务员", width=70)
        self.recent_projects_tree.column("型号数量", width=80)
        self.recent_projects_tree.column("技术规范数量", width=100)
        self.recent_projects_tree.column("创建时间", width=90)

        # 滚动条
        scrollbar_recent = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.recent_projects_tree.yview)
        self.recent_projects_tree.configure(yscrollcommand=scrollbar_recent.set)
        
        self.recent_projects_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_recent.pack(side=tk.RIGHT, fill=tk.Y)

        # 双击事件绑定
        self.recent_projects_tree.bind("<Double-1>", self.on_recent_project_double_click)

        # 右键菜单
        self.recent_context_menu = tk.Menu(self.root, tearoff=0)
        self.recent_context_menu.add_command(label="📁 打开项目文件夹", command=self.open_recent_project_folder)
        self.recent_context_menu.add_command(label="📝 编辑项目信息", command=self.edit_recent_project)
        self.recent_context_menu.add_command(label="📊 编辑数据统计", command=self.edit_project_statistics)
        self.recent_context_menu.add_command(label="🗑️ 删除记录", command=self.delete_recent_project)
        
        self.recent_projects_tree.bind("<Button-3>", self.show_recent_context_menu)

        # 按钮框架
        button_frame = tk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        tk.Button(button_frame, text="🔄 刷新列表", command=self.refresh_recent_projects,
                 bg="#2196F3", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="📊 导出统计", command=self.export_project_statistics,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="🗑️ 清空列表", command=self.clear_recent_projects,
                 bg="#f44336", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="📂 项目导入", command=self.import_projects_from_folder,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)

        # 加载近期项目
        self.load_recent_projects()

    def load_recent_projects(self):
        """加载近期项目列表"""
        recent_projects = self.config.get("recent_projects", [])
        
        # 更新月份筛选选项（先更新月份）
        self.update_month_filter_options(recent_projects)
        
        # 根据当前选择的月份，筛选项目用于更新业务员选项
        selected_month = self.filter_month_var.get() if hasattr(self, 'filter_month_var') else "全部"
        
        if selected_month == "全部":
            # 显示所有月份的业务员
            projects_for_manager_filter = recent_projects
        else:
            # 只显示选中月份的项目的业务员
            projects_for_manager_filter = []
            for project in recent_projects:
                project_code = project.get("code", "")
                month_str = self.extract_month_from_project_code(project_code)
                if month_str == selected_month:
                    projects_for_manager_filter.append(project)
        
        # 更新业务员筛选选项（基于月份筛选后的项目）
        self.update_manager_filter_options(projects_for_manager_filter)
        
        # 应用筛选
        filtered_projects = self.apply_project_filter_logic(recent_projects)
        
        # 清空现有项目
        for item in self.recent_projects_tree.get_children():
            self.recent_projects_tree.delete(item)
        
        # 按创建时间倒序排列，显示筛选后的项目
        filtered_projects.sort(key=lambda x: x.get("created_time", ""), reverse=True)
        
        # 移除显示数量限制，显示所有筛选后的项目
        for project in filtered_projects:  # 显示所有项目，不再限制为50个
            self.recent_projects_tree.insert("", "end", values=(
                project.get("code", ""),
                project.get("name", ""),
                project.get("manager", ""),
                project.get("model_count", 0),  # 型号数量
                project.get("spec_count", 0),   # 技术规范数量
                project.get("created_time", "").split(" ")[0] if project.get("created_time") else ""  # 只显示日期部分
            ))
        
        # 更新统计信息（新增）
        total_count = len(recent_projects)
        filtered_count = len(filtered_projects)
        if hasattr(self, 'recent_projects_stats_label'):
            stats_text = f"显示 {filtered_count}/{total_count} 个项目"
            self.recent_projects_stats_label.config(text=stats_text)

    def update_manager_filter_options(self, projects):
        """更新业务员筛选选项（保留当前选择）"""
        # 保存当前选择的业务员
        current_selection = self.filter_manager_var.get() if hasattr(self, 'filter_manager_var') else "全部"
        
        managers = set()
        for project in projects:
            manager = project.get("manager", "")
            if manager:
                managers.add(manager)
        
        # 更新业务员筛选下拉框
        manager_options = ["全部"] + sorted(list(managers))
        self.manager_filter_combo['values'] = manager_options
        
        # 恢复之前的选择，如果该选项仍然存在的话
        if current_selection in manager_options:
            self.filter_manager_var.set(current_selection)
        else:
            # 如果之前选择的业务员不存在了，重置为"全部"
            self.filter_manager_var.set("全部")
    
    def update_month_filter_options(self, projects):
        """更新月份筛选选项（保留当前选择）"""
        # 保存当前选择的月份
        current_selection = self.filter_month_var.get() if hasattr(self, 'filter_month_var') else "全部"
        
        months = set()
        for project in projects:
            project_code = project.get("code", "")
            month_str = self.extract_month_from_project_code(project_code)
            if month_str:
                months.add(month_str)
        
        # 更新月份筛选下拉框，按时间倒序排列
        month_options = ["全部"] + sorted(list(months), reverse=True)
        self.month_filter_combo['values'] = month_options
        
        # 恢复之前的选择，如果该选项仍然存在的话
        if current_selection in month_options:
            self.filter_month_var.set(current_selection)
        else:
            # 如果之前选择的月份不存在了，重置为"全部"
            self.filter_month_var.set("全部")

    def apply_project_filter_logic(self, projects):
        """应用项目筛选逻辑（支持月份筛选）"""
        filtered_projects = []
        
        # 获取筛选条件
        month_filter = getattr(self, 'filter_month_var', None)  # 新增：月份筛选
        manager_filter = getattr(self, 'filter_manager_var', None)
        time_filter = getattr(self, 'filter_time_var', None)
        
        month_value = month_filter.get() if month_filter else "全部"
        manager_value = manager_filter.get() if manager_filter else "全部"
        time_value = time_filter.get() if time_filter else "全部"
        
        from datetime import datetime, timedelta
        now = datetime.now()
        
        for project in projects:
            # 月份筛选（新增）
            if month_value != "全部":
                project_code = project.get("code", "")
                project_month = self.extract_month_from_project_code(project_code)
                if project_month != month_value:
                    continue
            
            # 业务员筛选
            if manager_value != "全部":
                if project.get("manager", "") != manager_value:
                    continue
            
            # 时间筛选
            if time_value != "全部":
                project_time_str = project.get("created_time", "")
                if not project_time_str:
                    continue
                
                try:
                    project_time = datetime.strptime(project_time_str, "%Y-%m-%d %H:%M:%S")
                    
                    if time_value == "今天":
                        if project_time.date() != now.date():
                            continue
                    elif time_value == "本周":
                        week_start = now - timedelta(days=now.weekday())
                        if project_time < week_start:
                            continue
                    elif time_value == "本月":
                        month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                        if project_time < month_start:
                            continue
                    elif time_value == "最近3个月":
                        three_months_ago = now - timedelta(days=90)
                        if project_time < three_months_ago:
                            continue
                    elif time_value == "最近6个月":
                        six_months_ago = now - timedelta(days=180)
                        if project_time < six_months_ago:
                            continue
                    elif time_value == "今年":
                        year_start = now.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
                        if project_time < year_start:
                            continue
                except ValueError:
                    # 时间格式错误，跳过此项目
                    continue
            
            filtered_projects.append(project)
        
        return filtered_projects
    
    def on_recent_month_filter_change(self, event=None):
        """月份筛选改变时的级联处理"""
        # 获取所有项目
        all_projects = self.config.get("recent_projects", [])
        
        # 根据选择的月份筛选项目
        selected_month = self.filter_month_var.get()
        
        if selected_month == "全部":
            # 显示所有月份的业务员
            filtered_projects = all_projects
        else:
            # 只显示选中月份的项目
            filtered_projects = []
            for project in all_projects:
                project_code = project.get("code", "")
                month_str = self.extract_month_from_project_code(project_code)
                if month_str == selected_month:
                    filtered_projects.append(project)
        
        # 更新业务员筛选选项（只显示筛选后项目中的业务员）
        self.update_manager_filter_options(filtered_projects)
        
        # 应用筛选
        self.apply_project_filters()

    def apply_project_filters(self, event=None):
        """应用项目筛选"""
        self.load_recent_projects()

    def clear_project_filters(self):
        """清空项目筛选"""
        if hasattr(self, 'filter_month_var'):  # 新增：清空月份筛选
            self.filter_month_var.set("全部")
        if hasattr(self, 'filter_manager_var'):
            self.filter_manager_var.set("全部")
        if hasattr(self, 'filter_time_var'):
            self.filter_time_var.set("全部")
        self.load_recent_projects()

    def add_recent_project(self, code, name, manager, folder_path, model_count=0, spec_count=0):
        """添加项目到近期列表"""
        recent_projects = self.config.get("recent_projects", [])
        
        # 检查是否已存在相同项目（根据项目编号判断）
        existing_index = -1
        for i, project in enumerate(recent_projects):
            if project.get("code") == code:
                existing_index = i
                break
        
        # 创建新项目记录
        new_project = {
            "code": code,
            "name": name,
            "manager": manager,
            "folder_path": folder_path,
            "model_count": model_count,
            "spec_count": spec_count,
            "created_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # 如果已存在，更新记录；否则添加新记录
        if existing_index >= 0:
            # 保留原有的数量数据，除非明确传入了新值
            old_project = recent_projects[existing_index]
            if model_count == 0:
                new_project["model_count"] = old_project.get("model_count", 0)
            if spec_count == 0:
                new_project["spec_count"] = old_project.get("spec_count", 0)
            recent_projects[existing_index] = new_project
        else:
            recent_projects.insert(0, new_project)
        
        # 移除存储数量限制，保留所有项目记录
        # 注释掉原有的50个限制
        # if len(recent_projects) > 50:
        #     recent_projects = recent_projects[:50]
        
        self.config["recent_projects"] = recent_projects
        self.save_config()
        self.load_recent_projects()

    def on_recent_project_double_click(self, event):
        """双击近期项目时打开文件夹"""
        self.open_recent_project_folder()

    def show_recent_context_menu(self, event):
        """显示近期项目右键菜单"""
        item = self.recent_projects_tree.identify_row(event.y)
        if item:
            self.recent_projects_tree.selection_set(item)
            self.recent_context_menu.post(event.x_root, event.y_root)

    def open_recent_project_folder(self):
        """打开选中的近期项目文件夹"""
        selected = self.recent_projects_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个项目！")
            return
        
        item = self.recent_projects_tree.item(selected[0])
        project_code = item['values'][0]
        
        # 从配置中找到对应的项目路径
        recent_projects = self.config.get("recent_projects", [])
        for project in recent_projects:
            if project.get("code") == project_code:
                folder_path = project.get("folder_path", "")
                if folder_path and os.path.exists(folder_path):
                    os.startfile(folder_path)
                    return
                else:
                    messagebox.showerror("错误", f"项目文件夹不存在：\n{folder_path}")
                    return
        
        messagebox.showerror("错误", "未找到项目文件夹路径！")

    def edit_recent_project(self):
        """编辑选中的近期项目信息"""
        selected = self.recent_projects_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个项目！")
            return
        
        item = self.recent_projects_tree.item(selected[0])
        project_code = item['values'][0]
        project_name = item['values'][1]
        project_manager = item['values'][2]
        
        # 填充到输入框
        self.project_code_var.set(project_code)
        self.project_name_var.set(project_name)
        self.project_manager_var.set(project_manager)
        
        messagebox.showinfo("提示", "项目信息已加载到输入框，您可以修改后重新生成项目文件夹")

    def edit_project_statistics(self):
        """编辑项目统计数据"""
        selected = self.recent_projects_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个项目！")
            return
        
        item = self.recent_projects_tree.item(selected[0])
        project_code = item['values'][0]
        project_name = item['values'][1]
        current_model_count = item['values'][3]
        current_spec_count = item['values'][4]
        
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑项目统计数据 - {project_code}")
        dialog.geometry("450x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 设置对话框图标和属性
        try:
            dialog.iconbitmap(default=self.root.iconbitmap())
        except:
            pass
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (350 // 2)
        dialog.geometry(f"450x350+{x}+{y}")
        
        # 设置背景色
        dialog.configure(bg="#f0f0f0")
        
        # 标题
        title_label = tk.Label(dialog, text="编辑项目统计数据", 
                              font=("Microsoft YaHei", 14, "bold"), bg="#f0f0f0", fg="#333")
        title_label.pack(pady=20)
        
        # 主容器
        main_container = tk.Frame(dialog, bg="#f0f0f0")
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)
        
        # 项目信息框架
        info_frame = tk.LabelFrame(main_container, text="项目信息", 
                                  font=("Microsoft YaHei", 11, "bold"), bg="#f0f0f0", padx=15, pady=10)
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(info_frame, text=f"项目编号：{project_code}", 
                font=("Microsoft YaHei", 11), bg="#f0f0f0", fg="#333").pack(anchor=tk.W, pady=2)
        tk.Label(info_frame, text=f"项目名称：{project_name}", 
                font=("Microsoft YaHei", 11), bg="#f0f0f0", fg="#666").pack(anchor=tk.W, pady=2)
        
        # 数据编辑框架
        data_frame = tk.LabelFrame(main_container, text="统计数据", 
                                  font=("Microsoft YaHei", 11, "bold"), bg="#f0f0f0", padx=15, pady=15)
        data_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 型号数量
        model_row = tk.Frame(data_frame, bg="#f0f0f0")
        model_row.pack(fill=tk.X, pady=8)
        
        tk.Label(model_row, text="型号数量:", font=("Microsoft YaHei", 11), 
                bg="#f0f0f0", width=12, anchor=tk.W).pack(side=tk.LEFT)
        
        model_count_var = tk.StringVar(value=str(current_model_count))
        model_count_entry = tk.Entry(model_row, textvariable=model_count_var, 
                                   font=("Microsoft YaHei", 11), width=15, relief=tk.SOLID, bd=1)
        model_count_entry.pack(side=tk.LEFT, padx=(10, 0))
        
        # 技术规范数量
        spec_row = tk.Frame(data_frame, bg="#f0f0f0")
        spec_row.pack(fill=tk.X, pady=8)
        
        tk.Label(spec_row, text="技术规范数量:", font=("Microsoft YaHei", 11), 
                bg="#f0f0f0", width=12, anchor=tk.W).pack(side=tk.LEFT)
        
        spec_count_var = tk.StringVar(value=str(current_spec_count))
        spec_count_entry = tk.Entry(spec_row, textvariable=spec_count_var, 
                                  font=("Microsoft YaHei", 11), width=15, relief=tk.SOLID, bd=1)
        spec_count_entry.pack(side=tk.LEFT, padx=(10, 0))
        
        # 保存函数
        def save_statistics():
            try:
                model_count = int(model_count_var.get() or 0)
                spec_count = int(spec_count_var.get() or 0)
                
                if model_count < 0 or spec_count < 0:
                    messagebox.showerror("错误", "数量不能为负数！", parent=dialog)
                    return
                
                # 更新项目数据
                recent_projects = self.config.get("recent_projects", [])
                updated = False
                for project in recent_projects:
                    if project.get("code") == project_code:
                        project["model_count"] = model_count
                        project["spec_count"] = spec_count
                        updated = True
                        break
                
                if updated:
                    self.config["recent_projects"] = recent_projects
                    self.save_config()
                    self.load_recent_projects()
                    
                    dialog.destroy()
                    messagebox.showinfo("成功", f"项目统计数据已更新！\n\n型号数量：{model_count}\n技术规范数量：{spec_count}")
                else:
                    messagebox.showerror("错误", "未找到对应的项目记录！", parent=dialog)
                
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字！", parent=dialog)
            except Exception as e:
                messagebox.showerror("错误", f"保存失败：{str(e)}", parent=dialog)
        
        def cancel_edit():
            dialog.destroy()
        
        # 按钮框架
        button_frame = tk.Frame(main_container, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, pady=20)
        
        # 保存按钮
        save_btn = tk.Button(button_frame, text="💾 保存", command=save_statistics,
                           bg="#4CAF50", fg="white", font=("Microsoft YaHei", 11, "bold"), 
                           width=12, height=2, relief=tk.RAISED, bd=2)
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 取消按钮
        cancel_btn = tk.Button(button_frame, text="❌ 取消", command=cancel_edit,
                             bg="#f44336", fg="white", font=("Microsoft YaHei", 11, "bold"), 
                             width=12, height=2, relief=tk.RAISED, bd=2)
        cancel_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # 设置焦点到第一个输入框
        model_count_entry.focus_set()
        model_count_entry.select_range(0, tk.END)
        
        # 绑定回车键保存
        def on_enter(event):
            save_statistics()
        
        dialog.bind('<Return>', on_enter)
        model_count_entry.bind('<Return>', on_enter)
        spec_count_entry.bind('<Return>', on_enter)
        
        # 绑定ESC键取消
        def on_escape(event):
            cancel_edit()
        
        dialog.bind('<Escape>', on_escape)

    def export_project_statistics(self):
        """导出项目统计数据"""
        recent_projects = self.config.get("recent_projects", [])
        
        if not recent_projects:
            messagebox.showwarning("警告", "没有项目数据可导出！")
            return
        
        try:
            from tkinter import filedialog
            
            # 选择保存位置
            filename = filedialog.asksaveasfilename(
                title="导出项目统计数据",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel文件", "*.xlsx"),
                    ("CSV文件", "*.csv"),
                    ("所有文件", "*.*")
                ]
            )
            
            if not filename:
                return
            
            # 准备数据
            data = []
            total_models = 0
            total_specs = 0
            
            for project in recent_projects:
                model_count = project.get("model_count", 0)
                spec_count = project.get("spec_count", 0)
                total_models += model_count
                total_specs += spec_count
                
                data.append([
                    project.get("code", ""),
                    project.get("name", ""),
                    project.get("manager", ""),
                    model_count,
                    spec_count,
                    project.get("created_time", ""),
                    project.get("folder_path", "")
                ])
            
            # 添加统计汇总行
            data.append([
                "合计",
                f"共{len(recent_projects)}个项目",
                "",
                total_models,
                total_specs,
                "",
                ""
            ])
            
            # 创建DataFrame并导出
            columns = ["项目编号", "项目名称", "业务员", "型号数量", "技术规范数量", "创建时间", "文件夹路径"]
            
            if filename.endswith('.csv'):
                # 导出CSV
                import csv
                with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(columns)
                    writer.writerows(data)
            else:
                # 导出Excel
                df = pd.DataFrame(data, columns=columns)
                df.to_excel(filename, index=False, sheet_name='项目统计')
            
            messagebox.showinfo("成功", f"项目统计数据已导出到：\n{filename}\n\n统计汇总：\n项目总数：{len(recent_projects)}\n型号总数：{total_models}\n技术规范总数：{total_specs}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def delete_recent_project(self):
        """删除选中的近期项目记录"""
        selected = self.recent_projects_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个项目！")
            return
        
        item = self.recent_projects_tree.item(selected[0])
        project_code = item['values'][0]
        project_name = item['values'][1]
        
        if messagebox.askyesno("确认删除", f"确定要删除项目记录吗？\n\n项目编号：{project_code}\n项目名称：{project_name}\n\n注意：这只会删除记录，不会删除实际的项目文件夹。"):
            recent_projects = self.config.get("recent_projects", [])
            recent_projects = [p for p in recent_projects if p.get("code") != project_code]
            self.config["recent_projects"] = recent_projects
            self.save_config()
            self.load_recent_projects()

    def refresh_recent_projects(self):
        """刷新近期项目列表"""
        self.load_recent_projects()
        messagebox.showinfo("提示", "近期项目列表已刷新")

    def clear_recent_projects(self):
        """清空近期项目列表"""
        if messagebox.askyesno("确认清空", "确定要清空所有近期项目记录吗？\n\n注意：这只会清空记录，不会删除实际的项目文件夹。"):
            self.config["recent_projects"] = []
            self.save_config()
            self.load_recent_projects()
            messagebox.showinfo("提示", "近期项目列表已清空")

    def create_settings_interface(self):
        """创建设置界面"""
        # 标题
        title_label = tk.Label(self.settings_frame, text="系统设置", 
                              font=("Microsoft YaHei", 16, "bold"), bg="#f0f0f0")
        title_label.pack(pady=15)
        
        # 创建滚动区域容器
        scroll_container = tk.Frame(self.settings_frame, bg="#f0f0f0")
        scroll_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 路径设置框架
        path_frame = ttk.LabelFrame(scroll_container, text="路径设置", padding=25)
        path_frame.pack(fill=tk.X, padx=30, pady=15)
        
        # 默认项目文件夹
        tk.Label(path_frame, text="默认项目文件夹:", font=("Microsoft YaHei", 11)).grid(row=0, column=0, sticky=tk.W, pady=10)
        default_frame = tk.Frame(path_frame)
        default_frame.grid(row=0, column=1, pady=10, padx=15, sticky=tk.W)
        
        self.default_folder_var = tk.StringVar(value=self.config.get("default_folder", ""))
        tk.Entry(default_frame, textvariable=self.default_folder_var, width=50, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        tk.Button(default_frame, text="浏览", 
                 command=lambda: self.browse_setting_folder("default_folder"), font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=8)
        
        # 定额文件夹
        tk.Label(path_frame, text="定额文件夹:", font=("Microsoft YaHei", 11)).grid(row=1, column=0, sticky=tk.W, pady=10)
        quota_frame = tk.Frame(path_frame)
        quota_frame.grid(row=1, column=1, pady=10, padx=15, sticky=tk.W)
        
        self.quota_folder_var = tk.StringVar(value=self.config.get("quota_folder", ""))
        tk.Entry(quota_frame, textvariable=self.quota_folder_var, width=50, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        tk.Button(quota_frame, text="浏览", 
                 command=lambda: self.browse_setting_folder("quota_folder"), font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=8)
        
        # 技术规范文件夹
        tk.Label(path_frame, text="技术规范文件夹:", font=("Microsoft YaHei", 11)).grid(row=2, column=0, sticky=tk.W, pady=10)
        spec_frame = tk.Frame(path_frame)
        spec_frame.grid(row=2, column=1, pady=10, padx=15, sticky=tk.W)
        
        self.spec_folder_var = tk.StringVar(value=self.config.get("spec_folder", ""))
        tk.Entry(spec_frame, textvariable=self.spec_folder_var, width=50, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        tk.Button(spec_frame, text="浏览", 
                 command=lambda: self.browse_setting_folder("spec_folder"), font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=8)
        
        # 保存设置按钮 - 固定在窗口底部
        button_frame = tk.Frame(self.settings_frame, bg="#f0f0f0")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10, padx=20)
        
        button_container = tk.Frame(button_frame, bg="#f0f0f0")
        button_container.pack()
        
        tk.Button(button_container, text="💾 保存设置", command=self.save_settings,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 12, "bold"),
                 height=2, width=15).pack()

    def create_parameter_card_interface(self):
        """创建参数卡片编辑界面"""
        # 标题
        title_label = tk.Label(self.card_frame, text="电缆产品参数卡片编辑", 
                              font=("Microsoft YaHei", 16, "bold"), bg="#f0f0f0")
        title_label.pack(pady=15)
        
        # 创建滚动区域容器
        scroll_container = tk.Frame(self.card_frame, bg="#f0f0f0")
        scroll_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建滚动区域
        canvas = tk.Canvas(scroll_container, bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 型号别名输入功能（支持多个）
        model_input_frame = ttk.LabelFrame(scrollable_frame, text="🔍 智能搜索与型号管理", padding=15)
        model_input_frame.pack(fill=tk.X, padx=30, pady=15)
        
        # 统一搜索区域
        search_frame = tk.Frame(model_input_frame)
        search_frame.grid(row=0, column=0, columnspan=4, sticky=tk.W+tk.E, pady=5)
        
        tk.Label(search_frame, text="搜索内容:", font=("Microsoft YaHei", 11, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.unified_search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.unified_search_var, width=40, font=("Microsoft YaHei", 11))
        search_entry.grid(row=0, column=1, pady=5, padx=10, sticky=tk.W)
        
        # 搜索按钮组
        button_frame = tk.Frame(search_frame)
        button_frame.grid(row=0, column=2, pady=5, padx=10)
        
        tk.Button(button_frame, text="型号搜索", command=self.unified_search_by_model,
                 bg="#2196F3", fg="white", font=("Microsoft YaHei", 10), width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(button_frame, text="结构搜索", command=self.unified_search_by_structure,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 10), width=10).pack(side=tk.LEFT, padx=2)
        # 暂时移除智能搜索功能
        # tk.Button(button_frame, text="智能搜索", command=self.unified_smart_search,
        #          bg="#FF9800", fg="white", font=("Microsoft YaHei", 10), width=10).pack(side=tk.LEFT, padx=2)
        
        # 搜索提示
        hint_frame = tk.Frame(model_input_frame)
        hint_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        tk.Label(hint_frame, text="搜索提示：", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT)
        tk.Label(hint_frame, text="型号搜索: YJV22.LV, NH.YJV.MV  |  结构搜索: CU/XLPE/STA/PVC", 
                font=("Microsoft YaHei", 9), fg="gray").pack(side=tk.LEFT, padx=5)
        
        # 型号绑定区域 - 已移除，功能与参数卡片保存重复
        # binding_frame = ttk.LabelFrame(model_input_frame, text="型号绑定管理", padding=10)
        # binding_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W+tk.E, pady=10)
        
        # 搜索结果显示区域
        self.create_search_results_area(scrollable_frame)
        
        # 核心参数卡片
        self.create_core_parameters_card(scrollable_frame)
        
        # 预测编码显示
        self.create_predicted_code_display(scrollable_frame)
        
        # 操作按钮 - 固定在窗口底部（不在滚动区域内）
        self.create_card_buttons(self.card_frame)

    def create_search_results_area(self, parent):
        """创建搜索结果显示区域"""
        results_frame = ttk.LabelFrame(parent, text="🔍 搜索结果", padding=15)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=15)
        
        # 创建Treeview显示搜索结果
        columns = ("置信度", "产品型号", "结构描述", "数据来源", "参数概要")
        self.search_tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=8)
        
        # 设置列标题和宽度
        self.search_tree.heading("置信度", text="置信度")
        self.search_tree.heading("产品型号", text="产品型号")
        self.search_tree.heading("结构描述", text="结构描述")
        self.search_tree.heading("数据来源", text="数据来源")
        self.search_tree.heading("参数概要", text="参数概要")
        
        self.search_tree.column("置信度", width=80)
        self.search_tree.column("产品型号", width=120)
        self.search_tree.column("结构描述", width=150)
        self.search_tree.column("数据来源", width=120)
        self.search_tree.column("参数概要", width=200)
        
        # 添加滚动条
        search_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.search_tree.yview)
        self.search_tree.configure(yscrollcommand=search_scrollbar.set)
        
        # 绑定双击事件
        self.search_tree.bind("<Double-1>", self.load_search_result_to_card)
        
        self.search_tree.pack(side="left", fill="both", expand=True)
        search_scrollbar.pack(side="right", fill="y")
        
        # 操作按钮
        search_button_frame = tk.Frame(results_frame)
        search_button_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(search_button_frame, text="加载到编辑区", command=self.load_selected_result,
                 bg="#2196F3", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(search_button_frame, text="清空结果", command=self.clear_search_results,
                 bg="#9E9E9E", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(search_button_frame, text="查看详细信息", command=self.show_search_detail,
                 bg="#607D8B", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)

    def create_core_parameters_card(self, parent):
        """创建核心参数卡片"""
        card = ttk.LabelFrame(parent, text="🔧 核心参数设置（10个必填字段）", padding=20)
        card.pack(fill=tk.X, padx=30, pady=15)
        
        # 第一行：产品大类和额定电压
        row = 0
        tk.Label(card, text="产品大类:", font=("Microsoft YaHei", 11)).grid(row=row, column=0, sticky=tk.W, pady=8)
        self.category_var = tk.StringVar()
        category_combo = ttk.Combobox(card, textvariable=self.category_var,
                                    values=list(self.voltage_levels.keys()), state="readonly", width=15)
        category_combo.grid(row=row, column=1, pady=8, padx=8)
        category_combo.bind('<<ComboboxSelected>>', self.on_card_category_change)
        
        tk.Label(card, text="额定电压:", font=("Microsoft YaHei", 11)).grid(row=row, column=2, sticky=tk.W, pady=8)
        self.voltage_rating_var = tk.StringVar()
        self.voltage_rating_combo = ttk.Combobox(card, textvariable=self.voltage_rating_var,
                                               state="readonly", width=15)
        self.voltage_rating_combo.grid(row=row, column=3, pady=8, padx=8)
        self.voltage_rating_combo.bind('<<ComboboxSelected>>', self.on_card_voltage_change)
        
        # 3kV提示标签（参数卡片版）
        self.card_kv3_warning_label = tk.Label(card, text="3kV需要有金属层", 
                                              font=("Microsoft YaHei", 9), fg="red")
        self.card_kv3_warning_label.grid(row=row, column=4, pady=8, padx=8)
        self.card_kv3_warning_label.grid_remove()  # 初始隐藏
        
        # 第二行：导体和绝缘
        row += 1
        tk.Label(card, text="导体材质:", font=("Microsoft YaHei", 11)).grid(row=row, column=0, sticky=tk.W, pady=8)
        self.conductor_card_var = tk.StringVar()
        conductor_combo = ttk.Combobox(card, textvariable=self.conductor_card_var,
                                     values=list(self.conductors.keys()), state="readonly", width=15)
        conductor_combo.grid(row=row, column=1, pady=8, padx=8)
        conductor_combo.bind('<<ComboboxSelected>>', self.update_predicted_code)
        
        tk.Label(card, text="绝缘材料:", font=("Microsoft YaHei", 11)).grid(row=row, column=2, sticky=tk.W, pady=8)
        self.insulation_card_var = tk.StringVar()
        insulation_combo = ttk.Combobox(card, textvariable=self.insulation_card_var,
                                      values=list(self.insulations.keys()), state="readonly", width=15)
        insulation_combo.grid(row=row, column=3, pady=8, padx=8)
        insulation_combo.bind('<<ComboboxSelected>>', self.update_predicted_code)
        
        # 第三行：屏蔽和内护套
        row += 1
        tk.Label(card, text="屏蔽类型:", font=("Microsoft YaHei", 11)).grid(row=row, column=0, sticky=tk.W, pady=8)
        self.shield_type_var = tk.StringVar()
        shield_combo = ttk.Combobox(card, textvariable=self.shield_type_var,
                                  values=list(self.shields.keys()), state="readonly", width=15)
        shield_combo.grid(row=row, column=1, pady=8, padx=8)
        shield_combo.bind('<<ComboboxSelected>>', self.update_predicted_code)
        
        tk.Label(card, text="内护套:", font=("Microsoft YaHei", 11)).grid(row=row, column=2, sticky=tk.W, pady=8)
        self.inner_sheath_var = tk.StringVar()
        inner_sheath_combo = ttk.Combobox(card, textvariable=self.inner_sheath_var,
                                        values=["无"] + list(self.sheaths.keys()), state="readonly", width=15)
        inner_sheath_combo.grid(row=row, column=3, pady=8, padx=8)
        inner_sheath_combo.bind('<<ComboboxSelected>>', self.update_predicted_code)
        
        # 第四行：铠装和外护套
        row += 1
        tk.Label(card, text="铠装类型:", font=("Microsoft YaHei", 11)).grid(row=row, column=0, sticky=tk.W, pady=8)
        self.armor_card_var = tk.StringVar()
        armor_combo = ttk.Combobox(card, textvariable=self.armor_card_var,
                                 values=list(self.armors.keys()), state="readonly", width=15)
        armor_combo.grid(row=row, column=1, pady=8, padx=8)
        armor_combo.bind('<<ComboboxSelected>>', self.update_predicted_code)
        
        tk.Label(card, text="外护套:", font=("Microsoft YaHei", 11)).grid(row=row, column=2, sticky=tk.W, pady=8)
        self.outer_sheath_var = tk.StringVar()
        outer_sheath_combo = ttk.Combobox(card, textvariable=self.outer_sheath_var,
                                        values=["无"] + list(self.sheaths.keys()), state="readonly", width=15)
        outer_sheath_combo.grid(row=row, column=3, pady=8, padx=8)
        outer_sheath_combo.bind('<<ComboboxSelected>>', self.update_predicted_code)
        
        # 第五行：是否耐火
        row += 1
        tk.Label(card, text="是否耐火:", font=("Microsoft YaHei", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=8)
        self.is_fire_resistant_var = tk.BooleanVar()
        fire_check = tk.Checkbutton(card, text="耐火电缆", variable=self.is_fire_resistant_var,
                                  font=("Microsoft YaHei", 10), command=self.update_predicted_code)
        fire_check.grid(row=row, column=1, pady=8, padx=8, sticky=tk.W)
        
        # 特殊性能选择
        tk.Label(card, text="特殊性能:", font=("Microsoft YaHei", 11)).grid(row=row, column=2, sticky=tk.W, pady=8)
        special_frame = tk.Frame(card)
        special_frame.grid(row=row, column=3, columnspan=2, pady=8, padx=8, sticky=tk.W)
        
        # 阻燃性能（互斥选择）- 默认选择"无"
        tk.Label(special_frame, text="阻燃等级:", font=("Microsoft YaHei", 10)).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.fire_rating_var = tk.StringVar(value="无")  # 默认选择"无"
        fire_rating_frame = tk.Frame(special_frame)
        fire_rating_frame.grid(row=0, column=1, columnspan=4, sticky=tk.W, pady=2, padx=5)
        
        fire_options = ["无", "ZA", "ZB", "ZC", "ZR"]
        for i, option in enumerate(fire_options):
            btn = tk.Radiobutton(fire_rating_frame, text=option, variable=self.fire_rating_var, 
                               value=option, font=("Microsoft YaHei", 9), command=self.on_parameter_change)
            btn.pack(side=tk.LEFT, padx=3)
        
        # 其他性能（多选）
        tk.Label(special_frame, text="其他性能:", font=("Microsoft YaHei", 10)).grid(row=1, column=0, sticky=tk.W, pady=2)
        other_frame = tk.Frame(special_frame)
        other_frame.grid(row=1, column=1, columnspan=4, sticky=tk.W, pady=2, padx=5)
        
        self.special_performance_vars = {}
        other_options = ["防鼠", "防白蚁", "耐油", "耐酸碱"]
        for i, option in enumerate(other_options):
            var = tk.BooleanVar()
            btn = tk.Checkbutton(other_frame, text=option, variable=var, font=("Microsoft YaHei", 9),
                               command=self.on_parameter_change)
            btn.pack(side=tk.LEFT, padx=3)
            self.special_performance_vars[option] = var

    def create_predicted_code_display(self, parent):
        """创建预测编码显示和确认区域"""
        display_frame = ttk.LabelFrame(parent, text="📊 产品信息预览与确认", padding=15)
        display_frame.pack(fill=tk.X, padx=30, pady=15)
        
        # 第一行：预测编码和确认按钮
        code_frame = tk.Frame(display_frame)
        code_frame.grid(row=0, column=0, columnspan=4, sticky=tk.W+tk.E, pady=5)
        
        tk.Label(code_frame, text="预测编码:", font=("Microsoft YaHei", 11, "bold")).pack(side=tk.LEFT)
        self.predicted_code_var = tk.StringVar(value="CBL-SPEC-XXXXXX")
        code_label = tk.Label(code_frame, textvariable=self.predicted_code_var, 
                            font=("Microsoft YaHei", 14, "bold"), fg="blue")
        code_label.pack(side=tk.LEFT, padx=15)
        
        # 确认按钮
        self.confirm_button = tk.Button(code_frame, text="🔍 确认编码", command=self.confirm_specification,
                                       bg="#4CAF50", fg="white", font=("Microsoft YaHei", 11, "bold"), width=12)
        self.confirm_button.pack(side=tk.RIGHT, padx=10)
        
        # 编码状态指示
        self.code_status_var = tk.StringVar(value="待确认")
        self.status_label = tk.Label(code_frame, textvariable=self.code_status_var, 
                               font=("Microsoft YaHei", 10), fg="orange")
        self.status_label.pack(side=tk.RIGHT, padx=10)
        
        # 第二行：结构字符串
        tk.Label(display_frame, text="结构字符串:", font=("Microsoft YaHei", 11, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.structure_string_var = tk.StringVar()
        structure_entry = tk.Entry(display_frame, textvariable=self.structure_string_var, 
                                 width=50, font=("Microsoft YaHei", 11))
        structure_entry.grid(row=1, column=1, columnspan=2, pady=5, padx=15, sticky=tk.W)
        
        tk.Label(display_frame, text="(可编辑)", font=("Microsoft YaHei", 9), fg="gray").grid(row=1, column=3, sticky=tk.W, pady=5)
        
        # 第三行：型号名称
        tk.Label(display_frame, text="型号名称:", font=("Microsoft YaHei", 11, "bold")).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.auto_model_name_var = tk.StringVar()
        model_entry = tk.Entry(display_frame, textvariable=self.auto_model_name_var, 
                             width=30, font=("Microsoft YaHei", 11))
        model_entry.grid(row=2, column=1, pady=5, padx=15, sticky=tk.W)
        
        tk.Button(display_frame, text="自动生成", command=self.auto_generate_model_name,
                 bg="#607D8B", fg="white", font=("Microsoft YaHei", 9), width=10).grid(row=2, column=2, pady=5, padx=5)
        
        tk.Label(display_frame, text="(自动生成+可编辑)", font=("Microsoft YaHei", 9), fg="gray").grid(row=2, column=3, sticky=tk.W, pady=5)
        
        # 第四行：产品描述
        tk.Label(display_frame, text="产品描述:", font=("Microsoft YaHei", 11)).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.description_var = tk.StringVar()
        desc_entry = tk.Entry(display_frame, textvariable=self.description_var, 
                            width=50, font=("Microsoft YaHei", 11))
        desc_entry.grid(row=3, column=1, columnspan=2, pady=5, padx=15, sticky=tk.W)
        
        # 第五行：确认后的操作区域
        confirm_frame = tk.Frame(display_frame)
        confirm_frame.grid(row=4, column=0, columnspan=4, sticky=tk.W+tk.E, pady=10)
        
        # 结构保存按钮（新编码时显示）
        self.structure_save_frame = tk.Frame(confirm_frame)
        
        tk.Label(self.structure_save_frame, text="新编码操作:", font=("Microsoft YaHei", 11, "bold"), fg="blue").pack(side=tk.LEFT)
        tk.Button(self.structure_save_frame, text="🔗 结构保存", command=self.save_structure_binding,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 10, "bold"), width=12).pack(side=tk.LEFT, padx=10)
        tk.Label(self.structure_save_frame, text="确认保存新编码的型号和结构绑定", 
                font=("Microsoft YaHei", 9), fg="gray").pack(side=tk.LEFT, padx=5)
        
        # 型号别名管理（确认后显示）
        self.alias_management_frame = tk.Frame(confirm_frame)
        
        tk.Label(self.alias_management_frame, text="关联型号别名:", font=("Microsoft YaHei", 11, "bold")).pack(side=tk.LEFT)
        self.confirmed_aliases_var = tk.StringVar()
        aliases_entry = tk.Entry(self.alias_management_frame, textvariable=self.confirmed_aliases_var, 
                               width=40, font=("Microsoft YaHei", 10))
        aliases_entry.pack(side=tk.LEFT, padx=10)
        
        tk.Button(self.alias_management_frame, text="添加别名", command=self.add_model_alias,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 9), width=10).pack(side=tk.LEFT, padx=5)
        
        # 初始隐藏所有管理区域
        # self.structure_save_frame.pack() 和 self.alias_management_frame.pack() 将根据状态显示

    def create_card_buttons(self, parent):
        """创建参数卡片操作按钮 - 固定在窗口底部"""
        button_frame = tk.Frame(parent, bg="#f0f0f0")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10, padx=20)
        
        # 创建按钮容器，居中显示
        button_container = tk.Frame(button_frame, bg="#f0f0f0")
        button_container.pack()
        
        tk.Button(button_container, text="� 保存参数卡片", command=self.save_parameter_card,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 12, "bold"), width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(button_container, text="🔄 重置参数", command=self.reset_parameter_card,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 12, "bold"), width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(button_container, text="📁 创建文件夹", command=self.create_product_folders,
                 bg="#2196F3", fg="white", font=("Microsoft YaHei", 12, "bold"), width=15, height=2).pack(side=tk.LEFT, padx=10)

    def create_product_management_interface(self):
        """创建产品编码管理界面"""
        # 创建滚动区域容器
        scroll_container = tk.Frame(self.management_frame, bg="#f0f0f0")
        scroll_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建主滚动框架
        main_canvas = tk.Canvas(scroll_container, bg="#f0f0f0")
        main_scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=main_canvas.yview)
        scrollable_frame = tk.Frame(main_canvas, bg="#f0f0f0")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=main_scrollbar.set)
        
        # 标题
        title_label = tk.Label(scrollable_frame, text="产品编码管理", 
                              font=("Microsoft YaHei", 16, "bold"), bg="#f0f0f0")
        title_label.pack(pady=15)
        
        # 筛选和搜索区域
        filter_frame = ttk.LabelFrame(scrollable_frame, text="🔍 筛选和搜索", padding=15)
        filter_frame.pack(fill=tk.X, padx=30, pady=(0, 15))
        
        # 第一行：产品大类筛选和搜索框
        filter_row1 = tk.Frame(filter_frame)
        filter_row1.pack(fill=tk.X, pady=5)
        
        tk.Label(filter_row1, text="产品大类:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_category_var = tk.StringVar(value="全部")
        category_filter = ttk.Combobox(filter_row1, textvariable=self.filter_category_var,
                                     values=["全部"] + list(self.voltage_levels.keys()), 
                                     state="readonly", width=12)
        category_filter.pack(side=tk.LEFT, padx=(5, 20))
        category_filter.bind('<<ComboboxSelected>>', self.apply_filters)
        
        tk.Label(filter_row1, text="搜索:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.search_filter_var = tk.StringVar()
        search_entry = tk.Entry(filter_row1, textvariable=self.search_filter_var, width=25, font=("Microsoft YaHei", 11))
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<KeyRelease>', self.apply_filters)
        
        tk.Button(filter_row1, text="清空筛选", command=self.clear_filters,
                 bg="#9E9E9E", fg="white", font=("Microsoft YaHei", 10), width=10).pack(side=tk.LEFT, padx=10)
        
        # 第二行：其他筛选选项
        filter_row2 = tk.Frame(filter_frame)
        filter_row2.pack(fill=tk.X, pady=5)
        
        tk.Label(filter_row2, text="导体:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_conductor_var = tk.StringVar(value="全部")
        conductor_filter = ttk.Combobox(filter_row2, textvariable=self.filter_conductor_var,
                                      values=["全部"] + list(self.conductors.keys()), 
                                      state="readonly", width=10)
        conductor_filter.pack(side=tk.LEFT, padx=(5, 20))
        conductor_filter.bind('<<ComboboxSelected>>', self.apply_filters)
        
        tk.Label(filter_row2, text="绝缘:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_insulation_var = tk.StringVar(value="全部")
        insulation_filter = ttk.Combobox(filter_row2, textvariable=self.filter_insulation_var,
                                       values=["全部"] + list(self.insulations.keys()), 
                                       state="readonly", width=10)
        insulation_filter.pack(side=tk.LEFT, padx=(5, 20))
        insulation_filter.bind('<<ComboboxSelected>>', self.apply_filters)
        
        tk.Label(filter_row2, text="耐火:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_fire_resistant_var = tk.StringVar(value="全部")
        fire_filter = ttk.Combobox(filter_row2, textvariable=self.filter_fire_resistant_var,
                                 values=["全部", "是", "否"], 
                                 state="readonly", width=8)
        fire_filter.pack(side=tk.LEFT, padx=(5, 20))
        fire_filter.bind('<<ComboboxSelected>>', self.apply_filters)
        
        # 第三行：新增筛选选项
        filter_row3 = tk.Frame(filter_frame)
        filter_row3.pack(fill=tk.X, pady=5)
        
        tk.Label(filter_row3, text="阻燃等级:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_flame_retardant_var = tk.StringVar(value="全部")
        flame_retardant_filter = ttk.Combobox(filter_row3, textvariable=self.filter_flame_retardant_var,
                                            values=["全部", "无"] + self.fire_resistant_options, 
                                            state="readonly", width=10)
        flame_retardant_filter.pack(side=tk.LEFT, padx=(5, 20))
        flame_retardant_filter.bind('<<ComboboxSelected>>', self.apply_filters)
        
        tk.Label(filter_row3, text="屏蔽类型:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_shield_var = tk.StringVar(value="全部")
        shield_filter = ttk.Combobox(filter_row3, textvariable=self.filter_shield_var,
                                   values=["全部"] + list(self.shields.keys()), 
                                   state="readonly", width=10)
        shield_filter.pack(side=tk.LEFT, padx=(5, 20))
        shield_filter.bind('<<ComboboxSelected>>', self.apply_filters)
        
        tk.Label(filter_row3, text="铠装类型:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.filter_armor_var = tk.StringVar(value="全部")
        armor_filter = ttk.Combobox(filter_row3, textvariable=self.filter_armor_var,
                                  values=["全部"] + list(self.armors.keys()), 
                                  state="readonly", width=10)
        armor_filter.pack(side=tk.LEFT, padx=5)
        armor_filter.bind('<<ComboboxSelected>>', self.apply_filters)
        
        # 产品列表
        list_frame = ttk.LabelFrame(scrollable_frame, text="📋 已保存的产品列表", padding=15)
        list_frame.pack(fill=tk.X, padx=30, pady=15)
        
        # 列可见性控制区域
        column_control_frame = tk.Frame(list_frame)
        column_control_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(column_control_frame, text="列显示控制:", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT)
        
        # 初始化列可见性状态
        self.column_visibility = {
            "规格编号": tk.BooleanVar(value=True),
            "产品型号": tk.BooleanVar(value=True),
            "结构描述": tk.BooleanVar(value=True),  # 新增结构描述列
            "大类": tk.BooleanVar(value=True),
            "电压": tk.BooleanVar(value=True),
            "导体": tk.BooleanVar(value=False),  # 默认隐藏
            "绝缘": tk.BooleanVar(value=False),  # 默认隐藏
            "屏蔽": tk.BooleanVar(value=False),  # 默认隐藏
            "铠装": tk.BooleanVar(value=False),  # 默认隐藏
            "护套": tk.BooleanVar(value=False),  # 默认隐藏
            "耐火": tk.BooleanVar(value=True),
            "特殊性能": tk.BooleanVar(value=True),
            "关联型号": tk.BooleanVar(value=True),
            "使用次数": tk.BooleanVar(value=True),
            "创建时间": tk.BooleanVar(value=False)  # 默认隐藏
        }
        
        # 创建列可见性复选框（分两行显示）
        checkbox_frame1 = tk.Frame(column_control_frame)
        checkbox_frame1.pack(side=tk.LEFT, padx=10)
        
        checkbox_frame2 = tk.Frame(column_control_frame)
        checkbox_frame2.pack(side=tk.LEFT, padx=10)
        
        # 第一行复选框
        first_row_columns = ["规格编号", "产品型号", "结构描述", "大类", "电压", "耐火", "特殊性能"]
        for i, col in enumerate(first_row_columns):
            cb = tk.Checkbutton(checkbox_frame1, text=col, variable=self.column_visibility[col],
                               command=self.update_column_visibility, font=("Microsoft YaHei", 9))
            cb.pack(side=tk.LEFT, padx=3)
        
        # 第二行复选框
        second_row_columns = ["导体", "绝缘", "屏蔽", "铠装", "护套", "关联型号", "使用次数", "创建时间"]
        for i, col in enumerate(second_row_columns):
            cb = tk.Checkbutton(checkbox_frame2, text=col, variable=self.column_visibility[col],
                               command=self.update_column_visibility, font=("Microsoft YaHei", 9))
            cb.pack(side=tk.LEFT, padx=3)
        
        # 创建Treeview - 添加结构描述列
        self.all_columns = ("规格编号", "产品型号", "结构描述", "大类", "电压", "导体", "绝缘", "屏蔽", "铠装", "护套", "耐火", "特殊性能", "关联型号", "使用次数", "创建时间")
        self.product_tree = ttk.Treeview(list_frame, columns=self.all_columns, show="headings", height=10)
        
        # 设置列标题和宽度
        self.column_widths = {
            "规格编号": 120,
            "产品型号": 120,
            "结构描述": 150,  # 新增结构描述列
            "大类": 80,
            "电压": 100,
            "导体": 60,
            "绝缘": 60,
            "屏蔽": 60,
            "铠装": 60,
            "护套": 60,
            "耐火": 50,
            "特殊性能": 100,
            "关联型号": 150,
            "使用次数": 80,
            "创建时间": 120
        }
        
        # 初始化列显示
        self.update_column_visibility()
        
        # 初始化排序状态
        self.sort_column = None
        self.sort_reverse = False
        
        for col in self.all_columns:
            self.product_tree.heading(col, text=col, 
                                     command=lambda c=col: self.sort_product_list(c))
            self.product_tree.column(col, width=self.column_widths.get(col, 80))
        
        # 添加滚动条
        tree_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.product_tree.yview)
        self.product_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        # 绑定事件
        self.product_tree.bind("<Double-1>", self.edit_product_card)
        self.product_tree.bind("<Button-3>", self.show_context_menu)  # 右键菜单
        
        self.product_tree.pack(side="left", fill="both", expand=True)
        tree_scrollbar.pack(side="right", fill="y")
        
        # 创建右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="📝 编辑", command=self.edit_selected_product)
        self.context_menu.add_command(label="📋 复制编号", command=self.copy_spec_id)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🗑️ 删除", command=self.delete_selected_product)
        
        # 状态栏
        self.status_frame = tk.Frame(scrollable_frame, bg="#f0f0f0")
        self.status_frame.pack(fill=tk.X, padx=30, pady=5)
        
        self.status_label = tk.Label(self.status_frame, text="就绪", font=("Microsoft YaHei", 10), fg="gray")
        self.status_label.pack(side=tk.LEFT)
        
        self.count_label = tk.Label(self.status_frame, text="", font=("Microsoft YaHei", 10), fg="blue")
        self.count_label.pack(side=tk.RIGHT)
        
        # 打包主滚动组件
        main_canvas.pack(side="left", fill="both", expand=True)
        main_scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 管理按钮 - 固定在窗口底部（不在滚动区域内）
        mgmt_button_frame = tk.Frame(self.management_frame, bg="#f0f0f0")
        mgmt_button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10, padx=20)
        
        # 创建按钮容器，居中显示
        button_container = tk.Frame(mgmt_button_frame, bg="#f0f0f0")
        button_container.pack()
        
        # 第一行按钮
        button_row1 = tk.Frame(button_container, bg="#f0f0f0")
        button_row1.pack(pady=5)
        
        tk.Button(button_row1, text="🔄 刷新列表", command=self.refresh_product_list,
                 bg="#607D8B", fg="white", font=("Microsoft YaHei", 11), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_row1, text="📝 编辑选中", command=self.edit_selected_product,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 11), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_row1, text="🗑️ 删除选中", command=self.delete_selected_product,
                 bg="#F44336", fg="white", font=("Microsoft YaHei", 11), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_row1, text="📊 导出列表", command=self.export_product_list,
                 bg="#9C27B0", fg="white", font=("Microsoft YaHei", 11), width=12).pack(side=tk.LEFT, padx=5)
        
        # 第二行按钮
        button_row2 = tk.Frame(button_container, bg="#f0f0f0")
        button_row2.pack(pady=5)
        
        tk.Button(button_row2, text="📁 打开定额文件夹", command=self.open_product_quota_folder,
                 bg="#2196F3", fg="white", font=("Microsoft YaHei", 11), width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(button_row2, text="📄 打开技术规范文件夹", command=self.open_product_spec_folder,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 11), width=18).pack(side=tk.LEFT, padx=5)
        tk.Button(button_row2, text="🔢 刷新使用次数", command=self.refresh_usage_count,
                 bg="#00BCD4", fg="white", font=("Microsoft YaHei", 11), width=15).pack(side=tk.LEFT, padx=5)
        
        # 初始加载产品列表
        self.refresh_product_list()

    def browse_folder(self):
        """浏览文件夹"""
        folder = filedialog.askdirectory(initialdir=self.save_path_var.get())
        if folder:
            self.save_path_var.set(folder)

    def browse_setting_folder(self, setting_key):
        """浏览设置文件夹"""
        folder = filedialog.askdirectory()
        if folder:
            if setting_key == "default_folder":
                self.default_folder_var.set(folder)
            elif setting_key == "quota_folder":
                self.quota_folder_var.set(folder)
            elif setting_key == "spec_folder":
                self.spec_folder_var.set(folder)

    def generate_project_folders(self):
        """生成项目文件夹（修复版）"""
        # 获取输入并清理换行符和制表符
        code = self.project_code_var.get().strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        name = self.project_name_var.get().strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        manager = self.project_manager_var.get().strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        save_path = self.save_path_var.get().strip()
        
        if not all([code, name, manager, save_path]):
            messagebox.showerror("错误", "请填写完整的项目信息！")
            return
        
        if not os.path.exists(save_path):
            messagebox.showerror("错误", "保存路径不存在！")
            return
        
        try:
            # 创建主项目文件夹（修复路径问题）
            folder_name = f"（陈颖）{code} {name} {manager}"
            
            # 清理文件夹名称，移除Windows不支持的字符
            clean_folder_name = self.sanitize_folder_name(folder_name)
            
            # 规范化路径
            project_path = os.path.normpath(os.path.join(save_path, clean_folder_name))
            
            if os.path.exists(project_path):
                if not messagebox.askyesno("确认", f"文件夹 '{clean_folder_name}' 已存在，是否覆盖？"):
                    return
            
            os.makedirs(project_path, exist_ok=True)
            
            # 创建子文件夹
            os.makedirs(os.path.join(project_path, "项目资料"), exist_ok=True)
            quota_path = os.path.join(project_path, "清单定额")
            os.makedirs(quota_path, exist_ok=True)
            
            # 创建Excel文件
            excel_name = f"{code}-清单-00.xlsx"
            excel_path = os.path.join(quota_path, excel_name)
            
            # 创建基础Excel模板
            self.create_excel_template(excel_path, code, name)
            
            # 更新配置
            self.config["default_folder"] = save_path
            self.save_config()
            
            # 添加到近期项目列表（使用原始名称）
            self.add_recent_project(code, name, manager, project_path)
            
            messagebox.showinfo("成功", f"项目文件夹创建成功！\n路径：{project_path}")
            
            # 询问是否打开文件夹
            if messagebox.askyesno("打开文件夹", "是否打开创建的项目文件夹？"):
                os.startfile(project_path)
                
        except Exception as e:
            messagebox.showerror("错误", f"创建文件夹失败：{str(e)}")

    def sanitize_folder_name(self, folder_name):
        """清理文件夹名称，移除Windows不支持的字符"""
        import re
        
        # 首先移除换行符和制表符
        clean_name = re.sub(r'[\r\n\t]', ' ', folder_name)
        
        # Windows文件系统不支持的字符
        invalid_chars = r'[<>:"/\\|?*]'
        
        # 替换无效字符为下划线
        clean_name = re.sub(invalid_chars, '_', clean_name)
        
        # 移除连续的空格和下划线
        clean_name = re.sub(r'\s+', ' ', clean_name)  # 多个空格变成一个
        clean_name = re.sub(r'_+', '_', clean_name)   # 多个下划线变成一个
        
        # 移除开头和结尾的下划线和空格
        clean_name = clean_name.strip('_ ')
        
        # 确保名称不为空
        if not clean_name:
            clean_name = "未命名项目"
        
        return clean_name

    def create_excel_template(self, file_path, project_code, project_name):
        """创建Excel模板"""
        try:
            # 创建空的电缆清单表
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 创建完全空的DataFrame
                df = pd.DataFrame()
                df.to_excel(writer, sheet_name='电缆清单', index=False)
                        
        except Exception as e:
            # 如果pandas不可用，创建简单的文本文件
            with open(file_path.replace('.xlsx', '.txt'), 'w', encoding='utf-8') as f:
                f.write(f"项目编号: {project_code}\n")
                f.write(f"项目名称: {project_name}\n")
                f.write(f"创建时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 50 + "\n")
                f.write("空白电缆清单模板\n")

    def save_settings(self):
        """保存设置"""
        self.config["default_folder"] = self.default_folder_var.get()
        self.config["quota_folder"] = self.quota_folder_var.get()
        self.config["spec_folder"] = self.spec_folder_var.get()
        
        try:
            self.save_config()
            messagebox.showinfo("成功", "设置已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败：{str(e)}")

    # ==================== 参数卡片相关方法 ====================
    
    def confirm_specification(self):
        """确认规格编码 - 核心方法"""
        try:
            # 1. 收集当前参数
            special_performance = []
            
            # 添加阻燃等级
            fire_rating = self.fire_rating_var.get()
            if fire_rating and fire_rating != "无":
                special_performance.append(fire_rating)
            
            # 添加其他性能
            for option, var in self.special_performance_vars.items():
                if var.get():
                    special_performance.append(option)
            
            # 2. 标准化内护套值
            inner_sheath_value = self.inner_sheath_var.get()
            if inner_sheath_value == "None":
                inner_sheath_value = "无"
            
            # 3. 创建产品对象
            product = CableProduct(
                category=self.category_var.get(),
                voltage_rating=self.voltage_rating_var.get(),
                conductor=self.conductor_card_var.get(),
                insulation=self.insulation_card_var.get(),
                shield_type=self.shield_type_var.get(),
                inner_sheath=inner_sheath_value,
                armor=self.armor_card_var.get(),
                outer_sheath=self.outer_sheath_var.get(),
                is_fire_resistant=self.is_fire_resistant_var.get(),
                special_performance=special_performance,
                model_name="",
                description=self.description_var.get()
            )
            
            # 4. 验证必填字段
            is_valid, missing_fields = product.validate()
            if not is_valid:
                missing_display = []
                field_names = {
                    'category': '产品大类',
                    'voltage_rating': '额定电压',
                    'conductor': '导体材质',
                    'insulation': '绝缘材料',
                    'shield_type': '屏蔽类型',
                    'armor': '铠装类型',
                    'outer_sheath': '外护套'
                }
                for field in missing_fields:
                    missing_display.append(field_names.get(field, field))
                
                messagebox.showerror("验证失败", f"以下字段为必填项：\n{chr(10).join(missing_display)}")
                return
            
            # 5. 计算参数哈希和生成编码
            param_hash = self.code_manager.calculate_param_hash(product)
            
            # 6. 检查是否已存在
            conn = sqlite3.connect(self.code_manager.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT spec_id FROM product_specs WHERE param_hash = ?", (param_hash,))
            existing = cursor.fetchone()
            
            if existing:
                # 复用已有编码
                spec_id = existing[0]
                self.predicted_code_var.set(spec_id)
                self.code_status_var.set("已存在")
                self.status_label.config(fg="green")
                
                # 加载已有规格的信息
                spec_info = self.code_manager.get_spec_by_id(spec_id)
                if spec_info:
                    spec = spec_info["spec"]
                    aliases = spec_info["aliases"]
                    
                    # 设置产品型号为保存的版本（从product_model字段）
                    saved_product_model = ""  # 初始化变量
                    if spec[11]:  # product_model字段
                        saved_product_model = spec[11]
                        self.auto_model_name_var.set(saved_product_model)
                        print(f"🔄 加载保存的产品型号: {saved_product_model}")
                    elif aliases:
                        # 如果没有产品型号，使用第一个别名作为备选
                        fallback_model = aliases[0][0]
                        self.auto_model_name_var.set(fallback_model)
                        print(f"🔄 使用别名作为产品型号: {fallback_model}")
                    
                    # 设置结构字符串为保存的版本
                    if spec[12]:  # structure_string字段
                        saved_structure = spec[12]
                        self.structure_string_var.set(saved_structure)
                        print(f"🔄 加载保存的结构字符串: {saved_structure}")
                    
                    # 设置所有关联型号（排除产品型号本身）
                    alias_names = [alias[0] for alias in aliases if alias[0] != saved_product_model]
                    self.confirmed_aliases_var.set(", ".join(alias_names))
                
                messagebox.showinfo("编码确认", f"找到相同参数的规格：{spec_id}\n将复用此编码")
                
            else:
                # 创建新编码 - 使用哈希前12位
                hash_suffix = param_hash[:12].upper()
                spec_id = f"CBL-SPEC-{hash_suffix}"
                
                self.predicted_code_var.set(spec_id)
                self.code_status_var.set("新建")
                self.status_label.config(fg="blue")
                
                # 对于新建编码，自动生成型号和结构
                self.auto_generate_model_name()
                self.update_structure_string()
                
                messagebox.showinfo("编码确认", f"创建新规格编码：{spec_id}")
            
            conn.close()
            
            # 7. 根据编码状态显示相应的操作区域（移到这里，避免重复调用生成方法）
            # 8. 根据编码状态显示相应的操作区域
            if existing:
                # 已存在编码：显示别名管理区域
                self.alias_management_frame.pack(fill=tk.X, pady=10)
                self.structure_save_frame.pack_forget()
            else:
                # 新建编码：显示结构保存按钮
                self.structure_save_frame.pack(fill=tk.X, pady=5)
                self.alias_management_frame.pack_forget()
            
            # 9. 更新按钮状态
            self.confirm_button.config(text="✓ 已确认", state="disabled")
            
            # 10. 存储当前确认的产品对象供后续使用
            self.confirmed_product = product
            self.confirmed_spec_id = spec_id
            
        except Exception as e:
            messagebox.showerror("错误", f"确认编码失败：{str(e)}")

    def save_structure_binding(self):
        """保存新编码的结构绑定 - 实现用户指定的操作顺序"""
        try:
            # 检查型号名称和结构描述是否已填写
            model_name = self.auto_model_name_var.get().strip()
            structure_desc = self.structure_string_var.get().strip()
            
            if not model_name:
                messagebox.showwarning("警告", "请输入型号名称！")
                return
            
            if not structure_desc:
                messagebox.showwarning("警告", "请确认结构描述！")
                return
            
            # 检查是否已确认编码
            if not hasattr(self, 'confirmed_product') or not hasattr(self, 'confirmed_spec_id'):
                messagebox.showwarning("警告", "请先确认编码！")
                return
            
            # 获取确认的产品和编码
            product = self.confirmed_product
            spec_id = self.confirmed_spec_id
            
            # 更新产品信息
            product.model_name = model_name
            product.description = self.description_var.get()
            
            # 保存到数据库
            conn = sqlite3.connect(self.code_manager.db_path)
            cursor = conn.cursor()
            
            # 创建新规格记录
            param_hash = self.code_manager.calculate_param_hash(product)
            product_dict = product.to_dict()
            special_performance_json = json.dumps(product_dict['special_performance'], ensure_ascii=False)
            
            cursor.execute('''
                INSERT OR REPLACE INTO product_specs (
                    spec_id, param_hash, category, voltage_rating, conductor, insulation,
                    shield_type, inner_sheath, armor, outer_sheath, is_fire_resistant,
                    special_performance, product_model, structure_string, created_date, modified_date, usage_count
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                spec_id, param_hash, product.category, product.voltage_rating, product.conductor,
                product.insulation, product.shield_type, product.inner_sheath, product.armor,
                product.outer_sheath, product.is_fire_resistant, special_performance_json,
                model_name, structure_desc, datetime.now().isoformat(), datetime.now().isoformat(), 1
            ))
            
            # 添加型号别名映射
            self.code_manager.add_alias_mapping(
                model_name, spec_id, "结构保存", 1.0, 
                f"用户通过结构保存功能创建：{structure_desc}", cursor
            )
            
            conn.commit()
            conn.close()
            
            # 更新界面状态
            self.code_status_var.set("已保存")
            self.status_label.config(fg="green")
            self.confirmed_aliases_var.set("")  # 不自动填充产品型号，保持空白
            
            # 隐藏结构保存按钮，显示别名管理区域
            self.structure_save_frame.pack_forget()
            self.alias_management_frame.pack(fill=tk.X, pady=10)
            
            # 刷新产品列表
            self.refresh_product_list()
            
            messagebox.showinfo("保存成功", f"新编码 {spec_id} 已保存！\n型号：{model_name}\n结构：{structure_desc}")
            
        except Exception as e:
            messagebox.showerror("错误", f"结构保存失败：{str(e)}")
    
    def auto_generate_model_name(self):
        """自动生成型号名称"""
        try:
            # 根据参数生成标准型号
            model_parts = []
            
            # 耐火前缀
            if self.is_fire_resistant_var.get():
                model_parts.append("NH")
            
            # 阻燃前缀
            fire_rating = self.fire_rating_var.get()
            if fire_rating and fire_rating != "无":
                model_parts.append(fire_rating)
            
            # 基础型号
            base_model = ""
            
            # 导体和绝缘
            conductor = self.conductor_card_var.get()
            insulation = self.insulation_card_var.get()
            
            if insulation == "XLPE":
                if conductor == "CU":
                    base_model = "YJV"
                elif conductor == "AL":
                    base_model = "YJLV"
            elif insulation == "PVC":
                if conductor == "CU":
                    base_model = "VV"
                elif conductor == "AL":
                    base_model = "VLV"
            elif insulation == "XLPO":
                base_model = "H1Z2Z2-K"
            
            # 屏蔽
            shield = self.shield_type_var.get()
            if shield == "CWS":
                base_model += "P"
            elif shield == "CTS":
                base_model += "P2"
            
            # 铠装
            armor = self.armor_card_var.get()
            if armor == "STA":
                base_model += "22"
            elif armor == "SWA":
                base_model += "32"
            elif armor == "SSTA":
                base_model += "62"
            elif armor == "AWA":
                base_model += "72"
            
            model_parts.append(base_model)
            
            # 电压等级后缀
            category = self.category_var.get()
            if category == "低压":
                model_parts.append("LV")
            elif category == "中压":
                model_parts.append("MV")
            elif category == "布线":
                voltage = self.voltage_rating_var.get()
                if "750V" in voltage:
                    model_parts.append("750V")
                elif "500V" in voltage:
                    model_parts.append("500V")
            
            # 组合型号
            if len(model_parts) > 1:
                # 使用点号连接主要部分
                if model_parts[0] in ["NH", "ZR", "ZA", "ZB", "ZC"]:
                    model_name = f"{model_parts[0]}.{'.'.join(model_parts[1:])}"
                else:
                    model_name = ".".join(model_parts)
            else:
                model_name = model_parts[0] if model_parts else "CUSTOM"
            
            self.auto_model_name_var.set(model_name)
            
        except Exception as e:
            self.auto_model_name_var.set("AUTO-GEN")
    
    def update_structure_string(self):
        """更新结构字符串"""
        try:
            # 根据当前参数生成结构字符串
            if hasattr(self, 'confirmed_product'):
                structure = self.confirmed_product.get_structure_string()
                self.structure_string_var.set(structure)
            else:
                # 临时生成
                parts = []
                
                conductor = self.conductor_card_var.get()
                if conductor:
                    parts.append(conductor)
                
                insulation = self.insulation_card_var.get()
                if insulation:
                    if self.is_fire_resistant_var.get():
                        parts.append(f"MT/{insulation}")
                    else:
                        parts.append(insulation)
                
                shield = self.shield_type_var.get()
                if shield and shield != "无":
                    parts.append(shield)
                
                inner_sheath = self.inner_sheath_var.get()
                if inner_sheath and inner_sheath != "无":
                    parts.append(inner_sheath)
                
                armor = self.armor_card_var.get()
                if armor and armor != "无":
                    parts.append(armor)
                
                outer_sheath = self.outer_sheath_var.get()
                if outer_sheath and outer_sheath != "无":
                    parts.append(outer_sheath)
                
                structure = "/".join(parts)
                self.structure_string_var.set(structure)
                
        except Exception as e:
            self.structure_string_var.set("结构生成失败")
    
    def add_model_alias(self):
        """添加型号别名"""
        if not hasattr(self, 'confirmed_spec_id'):
            messagebox.showwarning("警告", "请先确认编码！")
            return
        
        # 获取当前别名输入
        alias_input = self.confirmed_aliases_var.get().strip()
        if not alias_input:
            messagebox.showwarning("警告", "请输入型号别名！")
            return
        
        try:
            # 解析多个别名（逗号分隔）
            aliases = [alias.strip() for alias in alias_input.split(',') if alias.strip()]
            
            for alias in aliases:
                # 添加到映射表
                self.code_manager.add_alias_mapping(
                    alias, 
                    self.confirmed_spec_id, 
                    "手动添加", 
                    1.0, 
                    "用户手动添加的型号别名"
                )
            
            messagebox.showinfo("成功", f"已添加型号别名：{', '.join(aliases)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"添加别名失败：{str(e)}")

    def create_reverse_query_area(self, parent):
        """创建反向查询区域"""
        query_frame = ttk.LabelFrame(parent, text="🔍 反向查询功能", padding=15)
        query_frame.pack(fill=tk.X, padx=30, pady=15)
        
        # 方式A：通过型号查询
        tk.Label(query_frame, text="方式A - 型号查询:", font=("Microsoft YaHei", 11, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.model_query_var = tk.StringVar()
        model_query_entry = tk.Entry(query_frame, textvariable=self.model_query_var, width=25, font=("Microsoft YaHei", 11))
        model_query_entry.grid(row=0, column=1, pady=5, padx=10)
        
        tk.Button(query_frame, text="查询型号", command=self.query_by_model,
                 bg="#2196F3", fg="white", font=("Microsoft YaHei", 10), width=12).grid(row=0, column=2, pady=5, padx=5)
        
        tk.Label(query_frame, text="输入型号如：YJV22.LV", font=("Microsoft YaHei", 9), fg="gray").grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)
        
        # 方式B：通过结构查询
        tk.Label(query_frame, text="方式B - 结构查询:", font=("Microsoft YaHei", 11, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.structure_query_var = tk.StringVar()
        structure_query_entry = tk.Entry(query_frame, textvariable=self.structure_query_var, width=25, font=("Microsoft YaHei", 11))
        structure_query_entry.grid(row=1, column=1, pady=5, padx=10)
        
        tk.Button(query_frame, text="查询结构", command=self.query_by_structure,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 10), width=12).grid(row=1, column=2, pady=5, padx=5)
        
        tk.Label(query_frame, text="输入结构如：CU/XLPE/STA/PVC", font=("Microsoft YaHei", 9), fg="gray").grid(row=1, column=3, sticky=tk.W, pady=5, padx=5)
        
        # 查询结果显示区域
        result_frame = ttk.LabelFrame(query_frame, text="查询结果", padding=10)
        result_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W+tk.E, pady=10)
        
        # 创建查询结果树
        query_columns = ("规格编号", "匹配方式", "参数概要", "关联型号", "使用次数")
        self.query_tree = ttk.Treeview(result_frame, columns=query_columns, show="headings", height=6)
        
        for col in query_columns:
            self.query_tree.heading(col, text=col)
            if col == "规格编号":
                self.query_tree.column(col, width=120)
            elif col == "关联型号":
                self.query_tree.column(col, width=200)
            else:
                self.query_tree.column(col, width=100)
        
        # 添加滚动条
        query_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.query_tree.yview)
        self.query_tree.configure(yscrollcommand=query_scrollbar.set)
        
        # 绑定双击事件
        self.query_tree.bind("<Double-1>", self.load_query_result_to_card)
        
        self.query_tree.pack(side="left", fill="both", expand=True)
        query_scrollbar.pack(side="right", fill="y")
        
        # 查询操作按钮
        query_button_frame = tk.Frame(result_frame)
        query_button_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(query_button_frame, text="加载到编辑区", command=self.load_selected_query_result,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(query_button_frame, text="查看详细信息", command=self.show_query_detail,
                 bg="#607D8B", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(query_button_frame, text="清空结果", command=self.clear_query_results,
                 bg="#9E9E9E", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
    
    def query_by_model(self):
        """方式A：通过型号查询"""
        model_name = self.model_query_var.get().strip()
        if not model_name:
            messagebox.showwarning("警告", "请输入型号名称！")
            return
        
        try:
            # 清空现有结果
            for item in self.query_tree.get_children():
                self.query_tree.delete(item)
            
            # 搜索型号
            results = self.code_manager.search_by_alias(model_name)
            
            if not results:
                messagebox.showinfo("查询结果", f"未找到型号 '{model_name}' 的相关信息")
                return
            
            # 显示结果
            for result in results:
                if result["match_type"] == "确定" and result.get("spec_data"):
                    spec_data = result["spec_data"]
                    spec_id = result["spec_id"]
                    
                    # 构建参数概要
                    category = spec_data[0] if len(spec_data) > 0 else ""
                    voltage = spec_data[1] if len(spec_data) > 1 else ""
                    conductor = spec_data[2] if len(spec_data) > 2 else ""
                    insulation = spec_data[3] if len(spec_data) > 3 else ""
                    param_summary = f"{category} {voltage} {conductor}/{insulation}"
                    
                    # 获取所有关联型号
                    spec_info = self.code_manager.get_spec_by_id(spec_id)
                    if spec_info:
                        aliases = [alias[0] for alias in spec_info["aliases"]]
                        alias_text = ", ".join(aliases[:3])  # 显示前3个
                        if len(aliases) > 3:
                            alias_text += f" (+{len(aliases)-3})"
                        
                        usage_count = result.get("usage_count", 0)
                        
                        self.query_tree.insert("", "end", values=(
                            spec_id, "型号匹配", param_summary, alias_text, usage_count
                        ))
            
            messagebox.showinfo("查询完成", f"找到 {len(results)} 个匹配的规格")
            
        except Exception as e:
            messagebox.showerror("错误", f"查询失败：{str(e)}")
    
    def query_by_structure(self):
        """方式B：通过结构查询"""
        structure_query = self.structure_query_var.get().strip()
        if not structure_query:
            messagebox.showwarning("警告", "请输入结构描述！")
            return
        
        try:
            # 清空现有结果
            for item in self.query_tree.get_children():
                self.query_tree.delete(item)
            
            # 搜索结构
            results = self.code_manager.search_by_structure(structure_query)
            
            if not results:
                messagebox.showinfo("查询结果", f"未找到匹配结构 '{structure_query}' 的规格")
                return
            
            # 显示结果
            for result in results:
                spec_id = result["spec_id"]
                spec_data = result["spec_data"]
                
                # 构建参数概要
                category = spec_data[0] if len(spec_data) > 0 else ""
                voltage = spec_data[1] if len(spec_data) > 1 else ""
                conductor = spec_data[2] if len(spec_data) > 2 else ""
                insulation = spec_data[3] if len(spec_data) > 3 else ""
                param_summary = f"{category} {voltage} {conductor}/{insulation}"
                
                # 获取关联型号
                aliases = result.get("aliases", [])
                alias_names = [alias[0] for alias in aliases[:3]]
                alias_text = ", ".join(alias_names)
                if len(aliases) > 3:
                    alias_text += f" (+{len(aliases)-3})"
                
                usage_count = spec_data[11] if len(spec_data) > 11 else 0
                
                self.query_tree.insert("", "end", values=(
                    spec_id, "结构匹配", param_summary, alias_text, usage_count
                ))
            
            messagebox.showinfo("查询完成", f"找到 {len(results)} 个匹配的规格")
            
        except Exception as e:
            messagebox.showerror("错误", f"查询失败：{str(e)}")
    
    def load_query_result_to_card(self, event):
        """双击加载查询结果到参数卡片"""
        self.load_selected_query_result()
    
    def load_selected_query_result(self):
        """加载选中的查询结果"""
        selected = self.query_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个查询结果！")
            return
        
        item = self.query_tree.item(selected[0])
        spec_id = item['values'][0]
        
        try:
            # 获取规格详细信息
            spec_info = self.code_manager.get_spec_by_id(spec_id)
            if spec_info:
                spec = spec_info["spec"]
                aliases = spec_info["aliases"]
                
                # 加载到参数卡片
                self.load_spec_to_card(spec)
                
                # 设置型号别名 - 使用确认别名变量
                alias_names = [alias[0] for alias in aliases]
                self.confirmed_aliases_var.set(", ".join(alias_names))
                
                # 记录使用
                if alias_names:
                    self.code_manager.record_alias_usage(alias_names[0], spec_id)
                
                messagebox.showinfo("成功", f"已加载规格 {spec_id} 到编辑区")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载失败：{str(e)}")
    
    def show_query_detail(self):
        """显示查询结果的详细信息"""
        selected = self.query_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个查询结果！")
            return
        
        item = self.query_tree.item(selected[0])
        spec_id = item['values'][0]
        
        try:
            # 获取规格详细信息
            spec_info = self.code_manager.get_spec_by_id(spec_id)
            if spec_info:
                spec = spec_info["spec"]
                aliases = spec_info["aliases"]
                
                # 构建详细信息
                detail_info = f"规格编号：{spec[0]}\n"
                detail_info += f"产品大类：{spec[1]}\n"
                detail_info += f"额定电压：{spec[2]}\n"
                detail_info += f"导体材质：{spec[3]}\n"
                detail_info += f"绝缘材料：{spec[4]}\n"
                detail_info += f"屏蔽类型：{spec[5]}\n"
                detail_info += f"内护套：{spec[6] or '无'}\n"
                detail_info += f"铠装类型：{spec[7]}\n"
                detail_info += f"外护套：{spec[8]}\n"
                detail_info += f"是否耐火：{'是' if spec[9] else '否'}\n"
                
                # 特殊性能
                try:
                    special_perf = json.loads(spec[10]) if spec[10] else []
                    detail_info += f"特殊性能：{', '.join(special_perf) if special_perf else '无'}\n"
                except:
                    detail_info += f"特殊性能：{spec[10] or '无'}\n"
                
                detail_info += f"结构字符串：{spec[11] or '未生成'}\n"
                detail_info += f"使用次数：{spec[16] if len(spec) > 16 else 0}\n"
                detail_info += f"创建时间：{spec[14][:16] if spec[14] else '未知'}\n\n"
                
                # 关联型号
                detail_info += "关联型号别名：\n"
                for alias in aliases:
                    alias_name, confidence, source, usage_count = alias[:4]
                    detail_info += f"  • {alias_name} (置信度: {confidence}, 来源: {source}, 使用: {usage_count}次)\n"
                
                # 显示详细信息对话框
                detail_window = tk.Toplevel(self.root)
                detail_window.title(f"规格详细信息 - {spec_id}")
                detail_window.geometry("600x500")
                detail_window.configure(bg="#f0f0f0")
                
                # 创建滚动文本框
                text_frame = tk.Frame(detail_window)
                text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
                
                text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Microsoft YaHei", 10))
                scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
                text_widget.configure(yscrollcommand=scrollbar.set)
                
                text_widget.insert(tk.END, detail_info)
                text_widget.config(state=tk.DISABLED)
                
                text_widget.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")
                
                # 关闭按钮
                tk.Button(detail_window, text="关闭", command=detail_window.destroy,
                         bg="#9E9E9E", fg="white", font=("Microsoft YaHei", 11), width=10).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("错误", f"获取详细信息失败：{str(e)}")
    
    def clear_query_results(self):
        """清空查询结果"""
        for item in self.query_tree.get_children():
            self.query_tree.delete(item)

    def unified_search_by_model(self):
        """统一搜索 - 按型号搜索"""
        query = self.unified_search_var.get().strip()
        if not query:
            messagebox.showwarning("警告", "请输入型号名称！")
            return
        
        try:
            # 清空现有结果
            for item in self.search_tree.get_children():
                self.search_tree.delete(item)
            
            # 搜索型号
            raw_results = self.code_manager.search_by_alias(query)
            
            if not raw_results:
                messagebox.showinfo("搜索结果", f"未找到型号 '{query}' 的相关信息")
                return
            
            # 转换结果格式，添加display_name和优化置信度
            unified_results = []
            for result in raw_results:
                # 获取所有相关别名
                spec_id = result.get("spec_id")
                spec_data = result.get("spec_data")
                if spec_id:
                    # 获取该规格的所有别名
                    all_aliases = self.code_manager.get_spec_aliases(spec_id)
                    display_name = self.get_best_display_name(spec_id, all_aliases, spec_data)
                else:
                    all_aliases = []
                    display_name = query  # 对于候选结果，使用查询作为显示名称
                
                # 调整置信度（基于匹配类型）
                confidence = result.get("confidence", 0.5)
                source = result.get("source", "未知")
                
                # 精确匹配的置信度更高
                if query.upper() == display_name.upper():
                    confidence = min(confidence * 1.1, 1.0)
                elif source.startswith("产品型号匹配"):
                    confidence = min(confidence * 1.05, 1.0)
                
                unified_results.append({
                    "spec_id": spec_id,
                    "display_name": display_name,
                    "confidence": confidence,
                    "source": source,
                    "remarks": result.get("remarks", ""),
                    "usage_count": result.get("usage_count", 0),
                    "spec_data": result.get("spec_data"),
                    "aliases": all_aliases,
                    "candidate_params": result.get("candidate_params"),
                    "match_type": result.get("match_type", "型号")
                })
            
            # 按置信度排序
            unified_results.sort(key=lambda x: x['confidence'], reverse=True)
            
            # 显示结果
            self.display_unified_search_results(unified_results, "型号搜索")
            messagebox.showinfo("搜索完成", f"找到 {len(unified_results)} 个匹配的规格")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"型号搜索错误详情: {error_details}")
            messagebox.showerror("错误", f"型号搜索失败：{str(e)}")
    
    def unified_search_by_structure(self):
        """统一搜索 - 按结构搜索"""
        query = self.unified_search_var.get().strip()
        if not query:
            messagebox.showwarning("警告", "请输入结构描述！")
            return
        
        try:
            # 清空现有结果
            for item in self.search_tree.get_children():
                self.search_tree.delete(item)
            
            # 搜索结构
            results = self.code_manager.search_by_structure(query)
            
            if not results:
                messagebox.showinfo("搜索结果", f"未找到匹配结构 '{query}' 的规格")
                return
            
            # 转换结果格式以适应统一显示
            unified_results = []
            query_parts = set(part.upper().strip() for part in query.split('/') if part.strip())
            
            for result in results:
                spec_data = result["spec_data"]
                aliases = result["aliases"]
                
                # 安全地获取结构字符串和使用次数
                structure_string = spec_data[9] if len(spec_data) > 9 else "N/A"
                # usage_count 应该在最后一个位置，但要确保它是数字
                usage_count = 0
                if len(spec_data) > 11:  # 如果有12个字段，最后一个是usage_count
                    last_item = spec_data[11]
                    if isinstance(last_item, (int, float)):
                        usage_count = last_item
                elif len(spec_data) > 10:  # 如果只有11个字段，检查最后一个是否是数字
                    last_item = spec_data[10]
                    if isinstance(last_item, (int, float)):
                        usage_count = last_item
                
                # 计算置信度（基于结构匹配程度）
                confidence = self.calculate_structure_confidence(query, structure_string, query_parts, spec_data)
                
                # 获取最佳显示名称（优先使用数据库中的product_model）
                display_name = self.get_best_display_name(result["spec_id"], aliases, spec_data)
                
                unified_results.append({
                    "spec_id": result["spec_id"],
                    "display_name": display_name,
                    "confidence": confidence,
                    "source": "结构匹配",
                    "remarks": f"匹配结构: {structure_string}",
                    "usage_count": usage_count,
                    "spec_data": spec_data,
                    "aliases": aliases,
                    "match_type": "结构"
                })
            
            # 按置信度和使用次数排序
            unified_results.sort(key=lambda x: (x['confidence'], x['usage_count']), reverse=True)
            
            self.display_unified_search_results(unified_results, "结构搜索")
            messagebox.showinfo("搜索完成", f"找到 {len(unified_results)} 个匹配的规格")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"结构搜索错误详情: {error_details}")
            messagebox.showerror("错误", f"结构搜索失败：{str(e)}")
    
    def calculate_structure_confidence(self, query, structure_string, query_parts, spec_data):
        """计算结构匹配置信度"""
        if not structure_string or structure_string == "N/A":
            return 0.3
        
        # 基础置信度
        base_confidence = 0.5
        
        # 完全匹配检查
        if query.upper() == structure_string.upper():
            return 1.0
        
        # 部分匹配检查
        structure_parts = set(part.upper().strip() for part in structure_string.split('/') if part.strip())
        
        if not query_parts or not structure_parts:
            return base_confidence
        
        # 计算匹配度
        matched_parts = query_parts.intersection(structure_parts)
        match_ratio = len(matched_parts) / len(query_parts) if query_parts else 0
        
        # 根据匹配度计算置信度
        if match_ratio >= 1.0:
            confidence = 0.95  # 查询的所有部分都匹配
        elif match_ratio >= 0.8:
            confidence = 0.85  # 80%以上匹配
        elif match_ratio >= 0.6:
            confidence = 0.75  # 60%以上匹配
        elif match_ratio >= 0.4:
            confidence = 0.65  # 40%以上匹配
        elif match_ratio >= 0.2:
            confidence = 0.55  # 20%以上匹配
        else:
            confidence = base_confidence
        
        # 考虑结构复杂度（更复杂的结构匹配度稍微降低）
        structure_complexity = len(structure_parts)
        if structure_complexity > 6:
            confidence *= 0.95
        elif structure_complexity > 4:
            confidence *= 0.98
        
        # 考虑使用频率（使用次数高的规格置信度稍微提高）
        # 安全地获取usage_count，它应该在最后一个位置且是数字
        usage_count = 0
        if len(spec_data) > 11:  # 如果有12个字段，最后一个是usage_count
            last_item = spec_data[11]
            if isinstance(last_item, (int, float)):
                usage_count = last_item
        
        if usage_count > 10:
            confidence *= 1.05
        elif usage_count > 5:
            confidence *= 1.02
        
        return min(confidence, 1.0)
    
    def get_best_display_name(self, spec_id, aliases, spec_data=None):
        """获取最佳显示名称 - 优先显示数据库中的product_model"""
        
        # 第一优先级：使用数据库中的product_model字段
        if spec_data and len(spec_data) >= 11:
            # 尝试不同的索引位置来找到product_model
            for idx in [10, 11]:  # product_model可能在索引10或11
                if idx < len(spec_data):
                    candidate = spec_data[idx]
                    if isinstance(candidate, str) and candidate not in ['[]', 'None', '', 'null'] and not candidate.startswith('CBL-SPEC-'):
                        return candidate.strip()
        
        # 如果没有spec_data，尝试从数据库直接查询product_model
        if spec_id and not spec_id.startswith('CBL-SPEC-'):
            try:
                import sqlite3
                conn = sqlite3.connect("cable_products_v4.db")  # 使用固定路径
                cursor = conn.cursor()
                cursor.execute("SELECT product_model FROM product_specs WHERE spec_id = ?", (spec_id,))
                result = cursor.fetchone()
                if result and result[0] and result[0].strip():
                    conn.close()
                    return result[0].strip()
                conn.close()
            except:
                pass
        
        # 第二优先级：从别名中选择看起来像标准型号的别名
        if aliases:
            standard_models = []
            for alias_info in aliases:
                alias_name = alias_info[0].strip()
                
                # 跳过编码和过长描述
                if len(alias_name) > 50 or alias_name.startswith('CBL-SPEC-'):
                    continue
                
                # 优先选择看起来像型号的别名（包含字母和数字，不全是数字）
                if (any(c.isalpha() for c in alias_name) and 
                    any(c.isdigit() for c in alias_name) and
                    not alias_name.isdigit()):
                    standard_models.append(alias_name)
            
            # 从标准型号中选择最短的（通常是最简洁的型号）
            if standard_models:
                return min(standard_models, key=len)
            
            # 第三优先级：选择较短的非描述性别名
            short_aliases = []
            for alias_info in aliases:
                alias_name = alias_info[0].strip()
                
                # 跳过编码和过长描述
                if (len(alias_name) <= 30 and 
                    not alias_name.startswith('CBL-SPEC-') and
                    not alias_name.startswith('电缆') and
                    not alias_name.startswith('铜芯') and
                    not alias_name.startswith('铝芯')):
                    short_aliases.append(alias_name)
            
            if short_aliases:
                return short_aliases[0]
            
            # 第四优先级：如果只有描述性别名，选择最短的描述
            descriptive_aliases = []
            for alias_info in aliases:
                alias_name = alias_info[0].strip()
                if not alias_name.startswith('CBL-SPEC-') and len(alias_name) <= 50:
                    descriptive_aliases.append(alias_name)
            
            if descriptive_aliases:
                return min(descriptive_aliases, key=len)
        
        # 最后：如果实在没有合适的别名，返回通用名称
        return "未命名型号"
    
    def unified_smart_search(self):
        """统一搜索 - 智能搜索"""
        query = self.unified_search_var.get().strip()
        if not query:
            messagebox.showwarning("警告", "请输入搜索内容！")
            return
        
        try:
            # 清空现有结果
            for item in self.search_tree.get_children():
                self.search_tree.delete(item)
            
            results = []
            
            # 1. 按型号别名搜索
            alias_results = self.code_manager.search_by_alias(query)
            results.extend(alias_results)
            
            # 2. 按结构搜索（如果包含结构特征）
            if '/' in query or any(keyword in query.upper() for keyword in ['CU', 'AL', 'XLPE', 'PVC', 'STA', 'CTS']):
                structure_results = self.code_manager.search_by_structure(query)
                for sr in structure_results:
                    # 避免重复添加
                    if not any(r.get("spec_id") == sr["spec_id"] for r in results):
                        results.append({
                            "spec_id": sr["spec_id"],
                            "confidence": 0.8,
                            "source": "结构匹配",
                            "remarks": f"匹配结构: {sr['spec_data'][10]}",
                            "usage_count": sr["spec_data"][11],
                            "spec_data": sr["spec_data"],
                            "aliases": sr["aliases"],
                            "match_type": "结构"
                        })
            
            if results:
                self.display_unified_search_results(results, "智能搜索")
                messagebox.showinfo("搜索完成", f"找到 {len(results)} 个相关结果")
            else:
                messagebox.showinfo("搜索结果", "未找到相关结果")
                
        except Exception as e:
            messagebox.showerror("错误", f"智能搜索失败：{str(e)}")
    
    def display_unified_search_results(self, results, search_type):
        """显示统一搜索结果"""
        # 清空现有结果
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        for result in results:
            match_type = result.get("match_type", search_type)
            
            # 优先使用display_name，如果没有则使用spec_id
            display_id = result.get("display_name", result.get("spec_id", "新建"))
            
            # 确保不显示编码
            if display_id.startswith('CBL-SPEC-'):
                display_id = "未命名型号"
            
            confidence = f"{result['confidence']:.1%}" if result.get('confidence') else "N/A"
            source = result.get("source", "未知")
            
            if result.get("spec_data"):
                # 已有规格
                spec_data = result["spec_data"]
                category = spec_data[0] if len(spec_data) > 0 else ""
                voltage = spec_data[1] if len(spec_data) > 1 else ""
                conductor = spec_data[2] if len(spec_data) > 2 else ""
                insulation = spec_data[3] if len(spec_data) > 3 else ""
                shield_type = spec_data[4] if len(spec_data) > 4 else ""
                armor = spec_data[5] if len(spec_data) > 5 else ""
                outer_sheath = spec_data[6] if len(spec_data) > 6 else ""
                
                # 构建结构描述
                structure_parts = []
                if conductor: structure_parts.append(conductor)
                if insulation: structure_parts.append(insulation)
                if shield_type and shield_type != "无": structure_parts.append(shield_type)
                if armor and armor != "无": structure_parts.append(armor)
                if outer_sheath: structure_parts.append(outer_sheath)
                structure_desc = "/".join(structure_parts)
                
                param_summary = f"{category} {voltage}"
                
                # 获取关联型号（显示其他别名，但不显示编码）
                aliases = result.get("aliases", [])
                display_name = result.get("display_name", "")
                
                # 过滤掉已经作为主显示名称的别名，以及编码
                other_aliases = []
                for alias_info in aliases:
                    alias_name = alias_info[0].strip()
                    # 跳过编码、主显示名称和过长的别名
                    if (alias_name != display_name and 
                        len(alias_name) <= 40 and 
                        not alias_name.startswith('CBL-SPEC-')):
                        other_aliases.append(alias_name)
                
                # 最多显示2个其他别名
                alias_names = other_aliases[:2]
                alias_text = ", ".join(alias_names)
                if len(other_aliases) > 2:
                    alias_text += f" (+{len(other_aliases)-2}个)"
                
                # 如果没有其他别名，显示空字符串
                if not alias_text:
                    alias_text = ""
                    
            else:
                # 候选参数
                candidate = result.get("candidate_params", {})
                structure_desc = f"{candidate.get('conductor', '')}/{candidate.get('insulation', '')}"
                param_summary = f"{candidate.get('category', '')} {candidate.get('voltage_rating', '')}"
                alias_text = "待确认"
            
            # 根据置信度设置标签
            if result.get('confidence', 0) >= 0.9:
                tags = ("high_confidence",)
            elif result.get('confidence', 0) >= 0.7:
                tags = ("medium_confidence",)
            else:
                tags = ("low_confidence",)
            
            # 按新的列顺序插入数据：置信度、产品型号、结构描述、数据来源、参数概要
            # 同时在 tags 中存储 spec_id 用于后续加载
            item_id = self.search_tree.insert("", "end", values=(
                confidence, display_id, structure_desc, source, param_summary
            ), tags=tags)
            
            # 使用字典存储完整的结果数据，以item_id为键
            if not hasattr(self, 'search_result_data'):
                self.search_result_data = {}
            self.search_result_data[item_id] = result
        
        # 设置不同置信度的颜色
        self.search_tree.tag_configure("high_confidence", background="#E8F5E8")  # 绿色 - 高置信度
        self.search_tree.tag_configure("medium_confidence", background="#FFF3E0")  # 橙色 - 中置信度
        self.search_tree.tag_configure("low_confidence", background="#FFEBEE")  # 红色 - 低置信度
        
        # 保持原有的匹配类型颜色（作为备用）
        self.search_tree.tag_configure("确定", background="#E8F5E8")
        self.search_tree.tag_configure("推测", background="#FFF3E0")
        self.search_tree.tag_configure("结构", background="#E3F2FD")
        self.search_tree.tag_configure("型号搜索", background="#E8F5E8")
        self.search_tree.tag_configure("结构搜索", background="#E3F2FD")
        self.search_tree.tag_configure("智能搜索", background="#F3E5F5")
    
    def bind_model_to_current_params(self):
        """绑定型号到当前参数 - 已移除，功能与参数卡片保存重复"""
        # 功能已移除，请使用参数卡片的保存功能
        messagebox.showinfo("提示", "此功能已移除，请使用参数卡片的保存功能来保存型号和参数")
        if not model_name:
            messagebox.showwarning("警告", "请输入型号名称！")
            return
        
        try:
            # 检查是否已确认编码
            if not hasattr(self, 'confirmed_product') or not hasattr(self, 'confirmed_spec_id'):
                # 如果未确认编码，先收集当前参数
                special_performance = []
                fire_rating = self.fire_rating_var.get()
                if fire_rating and fire_rating != "无":
                    special_performance.append(fire_rating)
                
                for option, var in self.special_performance_vars.items():
                    if var.get():
                        special_performance.append(option)
                
                # 标准化内护套值
                inner_sheath_value = self.inner_sheath_var.get()
                if inner_sheath_value == "None":
                    inner_sheath_value = "无"
                
                # 创建产品对象
                product = CableProduct(
                    category=self.category_var.get(),
                    voltage_rating=self.voltage_rating_var.get(),
                    conductor=self.conductor_card_var.get(),
                    insulation=self.insulation_card_var.get(),
                    shield_type=self.shield_type_var.get(),
                    inner_sheath=inner_sheath_value,
                    armor=self.armor_card_var.get(),
                    outer_sheath=self.outer_sheath_var.get(),
                    is_fire_resistant=self.is_fire_resistant_var.get(),
                    special_performance=special_performance,
                    model_name=model_name,
                    description=self.description_var.get()
                )
                
                # 验证必填字段
                is_valid, missing_fields = product.validate()
                if not is_valid:
                    missing_display = []
                    field_names = {
                        'category': '产品大类',
                        'voltage_rating': '额定电压',
                        'conductor': '导体材质',
                        'insulation': '绝缘材料',
                        'shield_type': '屏蔽类型',
                        'armor': '铠装类型',
                        'outer_sheath': '外护套'
                    }
                    for field in missing_fields:
                        missing_display.append(field_names.get(field, field))
                    
                    messagebox.showerror("参数不完整", f"请先完善以下参数：\n{chr(10).join(missing_display)}")
                    return
                
                # 查找或创建规格
                spec_id = self.code_manager.find_or_create_spec(product, [model_name])
                
                # 更新界面状态
                self.predicted_code_var.set(spec_id)
                self.code_status_var.set("已绑定")
                self.status_label.config(fg="green")
                
                messagebox.showinfo("绑定成功", f"型号 '{model_name}' 已绑定到规格 {spec_id}")
                
            else:
                # 已确认编码，直接添加别名映射
                self.code_manager.add_alias_mapping(
                    model_name, 
                    self.confirmed_spec_id, 
                    "用户绑定", 
                    1.0, 
                    "用户手动绑定的型号"
                )
                
                # 更新显示
                current_aliases = self.confirmed_aliases_var.get()
                if current_aliases:
                    new_aliases = f"{current_aliases}, {model_name}"
                else:
                    new_aliases = model_name
                
                self.confirmed_aliases_var.set(new_aliases)
                messagebox.showinfo("绑定成功", f"型号 '{model_name}' 已绑定到规格 {self.confirmed_spec_id}")
            
            # 清空输入框
            self.model_binding_var.set("")
            
            # 刷新产品列表
            self.refresh_product_list()
            
        except Exception as e:
            messagebox.showerror("错误", f"绑定失败：{str(e)}")
    
    def show_search_detail(self):
        """显示搜索结果的详细信息"""
        selected = self.search_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个搜索结果！")
            return
        
        item_id = selected[0]
        item = self.search_tree.item(item_id)
        values = item['values']
        
        # 新的列顺序：置信度、产品型号、结构描述、数据来源、参数概要
        confidence = values[0]
        product_model = values[1]
        structure_desc = values[2]
        source = values[3]
        param_summary = values[4]
        
        try:
            # 从存储的结果数据中获取完整信息
            if hasattr(self, 'search_result_data') and item_id in self.search_result_data:
                result = self.search_result_data[item_id]
                spec_id = result.get("spec_id")
                
                if not spec_id or spec_id == "新建":
                    messagebox.showinfo("提示", "这是一个推测结果，没有详细信息")
                    return
                
                # 获取规格详细信息
                spec_info = self.code_manager.get_spec_by_id(spec_id)
                if spec_info:
                    spec = spec_info["spec"]
                    aliases = spec_info["aliases"]
                    
                    # 构建详细信息
                    detail_info = f"规格编号：{spec[0]}\n"
                    detail_info += f"产品大类：{spec[1]}\n"
                    detail_info += f"额定电压：{spec[2]}\n"
                    detail_info += f"导体材质：{spec[3]}\n"
                    detail_info += f"绝缘材料：{spec[4]}\n"
                    detail_info += f"屏蔽类型：{spec[5]}\n"
                    detail_info += f"内护套：{spec[6] or '无'}\n"
                    detail_info += f"铠装类型：{spec[7]}\n"
                    detail_info += f"外护套：{spec[8]}\n"
                    detail_info += f"是否耐火：{'是' if spec[9] else '否'}\n"
                    
                    # 特殊性能
                    try:
                        special_perf = json.loads(spec[10]) if spec[10] else []
                        detail_info += f"特殊性能：{', '.join(special_perf) if special_perf else '无'}\n"
                    except:
                        detail_info += f"特殊性能：{spec[10] or '无'}\n"
                    
                    detail_info += f"产品型号：{spec[11] or '未设置'}\n"
                    detail_info += f"结构字符串：{spec[12] or '未生成'}\n"
                    detail_info += f"使用次数：{spec[17] if len(spec) > 17 else 0}\n"
                    detail_info += f"创建时间：{spec[15][:16] if spec[15] else '未知'}\n\n"
                    
                    # 关联型号
                    detail_info += "关联型号别名：\n"
                    for alias in aliases:
                        alias_name, confidence_val, source_val, usage_count = alias[:4]
                        detail_info += f"  • {alias_name} (置信度: {confidence_val}, 来源: {source_val}, 使用: {usage_count}次)\n"
                    
                    # 显示详细信息对话框
                    detail_window = tk.Toplevel(self.root)
                    detail_window.title(f"规格详细信息 - {spec_id}")
                    detail_window.geometry("600x500")
                    detail_window.configure(bg="#f0f0f0")
                    
                    # 创建滚动文本框
                    text_frame = tk.Frame(detail_window)
                    text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
                    
                    text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Microsoft YaHei", 10))
                    scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
                    text_widget.configure(yscrollcommand=scrollbar.set)
                    
                    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                    
                    text_widget.insert(tk.END, detail_info)
                    text_widget.config(state=tk.DISABLED)
                    
                    # 关闭按钮
                    close_button = tk.Button(detail_window, text="关闭", command=detail_window.destroy,
                                           bg="#9E9E9E", fg="white", font=("Microsoft YaHei", 10))
                    close_button.pack(pady=10)
                    
                else:
                    messagebox.showerror("错误", f"无法找到规格 {spec_id}")
            else:
                messagebox.showwarning("警告", "无法获取搜索结果数据，请重新搜索")
                
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"显示详细信息错误详情: {error_details}")
            messagebox.showerror("错误", f"显示详细信息失败：{str(e)}")
    
    def parse_model_alias_to_card(self):
        """解析型号别名并填充参数卡片"""
        # 这个方法已被新的搜索功能替代，保留为兼容性
        messagebox.showinfo("提示", "请使用智能搜索功能来查找和加载型号参数")
    
    def smart_search(self):
        """智能搜索功能"""
        query = self.search_query_var.get().strip()
        if not query:
            messagebox.showwarning("警告", "请输入搜索内容！")
            return
        
        try:
            results = []
            
            # 1. 按型号别名搜索
            alias_results = self.code_manager.search_by_alias(query)
            results.extend(alias_results)
            
            # 2. 按结构搜索
            if '/' in query or any(keyword in query.upper() for keyword in ['CU', 'AL', 'XLPE', 'PVC', 'STA', 'CTS']):
                structure_results = self.code_manager.search_by_structure(query)
                for sr in structure_results:
                    results.append({
                        "spec_id": sr["spec_id"],
                        "confidence": 0.8,
                        "source": "结构匹配",
                        "remarks": f"匹配结构: {sr['spec_data'][10]}",  # structure_string
                        "usage_count": sr["spec_data"][11],  # usage_count
                        "spec_data": sr["spec_data"],
                        "aliases": sr["aliases"],
                        "match_type": "结构"
                    })
            
            if results:
                self.display_search_results(results)
                messagebox.showinfo("搜索完成", f"找到 {len(results)} 个相关结果")
            else:
                messagebox.showinfo("搜索结果", "未找到相关结果")
                
        except Exception as e:
            messagebox.showerror("错误", f"搜索失败：{str(e)}")
    
    def display_search_results(self, results):
        """显示搜索结果"""
        # 清空现有结果
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        for result in results:
            match_type = result["match_type"]
            spec_id = result.get("spec_id", "新建")
            confidence = f"{result['confidence']:.2f}" if result['confidence'] else "N/A"
            source = result["source"]
            
            if result.get("spec_data"):
                # 已有规格
                spec_data = result["spec_data"]
                category = spec_data[0] if len(spec_data) > 0 else ""
                voltage = spec_data[1] if len(spec_data) > 1 else ""
                conductor = spec_data[2] if len(spec_data) > 2 else ""
                insulation = spec_data[3] if len(spec_data) > 3 else ""
                param_summary = f"{category} {voltage} {conductor}/{insulation}"
                
                # 获取关联型号
                aliases = result.get("aliases", [])
                alias_names = [alias[0] for alias in aliases[:3]]  # 最多显示3个
                alias_text = ", ".join(alias_names)
                if len(aliases) > 3:
                    alias_text += f" (+{len(aliases)-3}个)"
            else:
                # 候选参数
                candidate = result.get("candidate_params", {})
                param_summary = f"{candidate.get('category', '')} {candidate.get('voltage_rating', '')} {candidate.get('conductor', '')}/{candidate.get('insulation', '')}"
                alias_text = "待确认"
            
            self.search_tree.insert("", "end", values=(
                match_type, spec_id, confidence, source, param_summary, alias_text
            ), tags=(match_type.lower(),))
        
        # 设置不同匹配类型的颜色
        self.search_tree.tag_configure("确定", background="#E8F5E8")
        self.search_tree.tag_configure("推测", background="#FFF3E0")
        self.search_tree.tag_configure("结构", background="#E3F2FD")
    
    def load_search_result_to_card(self, event):
        """双击加载搜索结果到参数卡片"""
        self.load_selected_result()
    
    def load_selected_result(self):
        """加载选中的搜索结果"""
        selected = self.search_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个搜索结果！")
            return
        
        item_id = selected[0]
        item = self.search_tree.item(item_id)
        values = item['values']
        
        # 新的列顺序：置信度、产品型号、结构描述、数据来源、参数概要
        confidence = values[0]
        product_model = values[1]
        structure_desc = values[2]
        source = values[3]
        param_summary = values[4]
        
        try:
            # 从存储的结果数据中获取完整信息
            if hasattr(self, 'search_result_data') and item_id in self.search_result_data:
                result = self.search_result_data[item_id]
                spec_id = result.get("spec_id")
                
                if spec_id and spec_id != "新建" and not spec_id.startswith("推测"):
                    # 加载已有规格
                    spec_info = self.code_manager.get_spec_by_id(spec_id)
                    if spec_info:
                        self.load_spec_to_card(spec_info["spec"])
                        # 设置型号别名
                        aliases = [alias[0] for alias in spec_info["aliases"]]
                        self.confirmed_aliases_var.set(", ".join(aliases))
                        messagebox.showinfo("成功", f"已加载产品型号 '{product_model}' 到编辑区")
                    else:
                        messagebox.showerror("错误", f"无法找到规格 {spec_id}")
                elif result.get("candidate_params"):
                    # 加载候选参数
                    self.load_candidate_to_card(result["candidate_params"])
                    messagebox.showinfo("提示", "已加载推测参数到编辑区，请确认后保存")
                else:
                    messagebox.showinfo("提示", "这是一个推测结果，请手动确认参数后保存")
            else:
                messagebox.showwarning("警告", "无法获取搜索结果数据，请重新搜索")
                
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"加载搜索结果错误详情: {error_details}")
            messagebox.showerror("错误", f"加载失败：{str(e)}")
    
    def load_result_to_card(self, result):
        """加载搜索结果到参数卡片"""
        if result.get("spec_data"):
            self.load_spec_to_card(result["spec_data"])
    
    def load_candidate_to_card(self, candidate_params):
        """加载候选参数到参数卡片"""
        self.category_var.set(candidate_params.get("category", ""))
        self.voltage_rating_var.set(candidate_params.get("voltage_rating", ""))
        self.conductor_card_var.set(candidate_params.get("conductor", ""))
        self.insulation_card_var.set(candidate_params.get("insulation", ""))
        self.shield_type_var.set(candidate_params.get("shield_type", ""))
        self.inner_sheath_var.set(candidate_params.get("inner_sheath", ""))
        self.armor_card_var.set(candidate_params.get("armor", ""))
        self.outer_sheath_var.set(candidate_params.get("outer_sheath", ""))
        self.is_fire_resistant_var.set(candidate_params.get("is_fire_resistant", False))
        
        # 处理特殊性能
        special_performance = candidate_params.get("special_performance", [])
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)
        
        for perf in special_performance:
            if perf in ["ZA", "ZB", "ZC", "ZR"]:
                self.fire_rating_var.set(perf)
            elif perf in self.special_performance_vars:
                self.special_performance_vars[perf].set(True)
        
        self.update_predicted_code()
    
    def load_spec_to_card(self, spec_data):
        """加载规格数据到参数卡片"""
        if len(spec_data) >= 11:
            self.category_var.set(spec_data[1] or "")
            # 仅更新电压选项，不设置默认值
            self.update_voltage_options_only()
            
            self.voltage_rating_var.set(spec_data[2] or "")
            self.conductor_card_var.set(spec_data[3] or "")
            self.insulation_card_var.set(spec_data[4] or "")
            self.shield_type_var.set(spec_data[5] or "")
            # 处理内护套的"无"值
            inner_sheath_value = spec_data[6] or "无"
            if inner_sheath_value == "None":
                inner_sheath_value = "无"
            self.inner_sheath_var.set(inner_sheath_value)
            self.armor_card_var.set(spec_data[7] or "")
            self.outer_sheath_var.set(spec_data[8] or "")
            self.is_fire_resistant_var.set(bool(spec_data[9]))
            
            # 处理特殊性能
            try:
                special_performance = json.loads(spec_data[10]) if spec_data[10] else []
                self.fire_rating_var.set("无")
                for var in self.special_performance_vars.values():
                    var.set(False)
                
                for perf in special_performance:
                    if perf in ["ZA", "ZB", "ZC", "ZR"]:
                        self.fire_rating_var.set(perf)
                    elif perf in self.special_performance_vars:
                        self.special_performance_vars[perf].set(True)
            except:
                pass
            
            self.predicted_code_var.set(f"{spec_data[0]} (已存在)")
            self.update_predicted_code()
    
    def clear_search_results(self):
        """清空搜索结果"""
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
    
    def add_manual_mapping(self):
        """手动添加型号映射"""
        selected = self.search_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个搜索结果！")
            return
        
        # 这里可以打开一个对话框让用户手动添加映射
        messagebox.showinfo("功能提示", "手动映射功能开发中...")

    def parse_model_to_card(self):
        """解析输入的型号并填充参数卡片"""
        input_model = self.input_model_card_var.get().strip().upper()
        if not input_model:
            messagebox.showwarning("警告", "请输入电缆型号！")
            return
        
        try:
            # 根据型号映射表反向查找参数
            found_match = False
            
            for structure, model in self.cable_type_mapping.items():
                if model == input_model:
                    # 解析结构字符串
                    parts = structure.split('/')
                    
                    # 设置产品大类
                    if ".LV" in input_model:
                        self.category_var.set("低压")
                        self.voltage_rating_var.set("0.6/1kV")
                    elif ".MV" in input_model:
                        self.category_var.set("中压")
                        self.voltage_rating_var.set("6/10kV")
                    elif input_model in ["H1Z2Z2.K"]:
                        self.category_var.set("光伏缆")
                        self.voltage_rating_var.set("DC 1500V")
                    elif input_model in ["PABC", "HDBC"]:
                        self.category_var.set("裸铜线")
                        self.voltage_rating_var.set("N/A")
                    elif any(x in input_model for x in ["KVV", "RVV"]):
                        self.category_var.set("控缆和仪表缆")
                        self.voltage_rating_var.set("450/750V")
                    else:
                        self.category_var.set("布线")
                        self.voltage_rating_var.set("450/750V")
                    
                    # 解析结构组件
                    if len(parts) >= 1:
                        self.conductor_card_var.set(parts[0])
                    if len(parts) >= 2:
                        # 处理耐火情况
                        if parts[1].startswith("MT/"):
                            self.is_fire_resistant_var.set(True)
                            self.insulation_card_var.set(parts[1][3:])  # 去掉MT/前缀
                        else:
                            self.insulation_card_var.set(parts[1])
                    if len(parts) >= 3:
                        if parts[2] in self.shields:
                            self.shield_type_var.set(parts[2])
                        else:
                            self.outer_sheath_var.set(parts[2])
                    if len(parts) >= 4:
                        self.outer_sheath_var.set(parts[3])
                    if len(parts) >= 5:
                        self.armor_card_var.set(parts[4])
                    if len(parts) >= 6:
                        self.outer_sheath_var.set(parts[5])
                    
                    found_match = True
                    break
            
            if not found_match:
                # 简化的型号解析
                if ".LV" in input_model:
                    self.category_var.set("低压")
                    self.voltage_rating_var.set("0.6/1kV")
                    self.set_card_low_voltage_defaults()
                elif ".MV" in input_model:
                    self.category_var.set("中压")
                    self.voltage_rating_var.set("6/10kV")
                    self.set_card_medium_voltage_defaults()
                elif input_model in ["H1Z2Z2.K"]:
                    self.category_var.set("光伏缆")
                    self.set_card_pv_defaults()
                elif input_model in ["PABC", "HDBC"]:
                    self.category_var.set("裸铜线")
                    self.set_card_bare_wire_defaults()
                elif any(x in input_model for x in ["KVV", "RVV"]):
                    self.category_var.set("控缆和仪表缆")
                    self.set_card_control_defaults()
                else:
                    self.category_var.set("布线")
                    self.set_card_wire_defaults()
            
            # 设置型号名称
            self.model_name_var.set(input_model)
            
            # 更新预测编码
            self.update_predicted_code()
            
            messagebox.showinfo("成功", f"型号 {input_model} 解析完成！")
            
        except Exception as e:
            messagebox.showerror("错误", f"型号解析失败：{str(e)}")

    def on_card_voltage_change(self, event=None):
        """参数卡片中电压等级变化时的处理"""
        voltage_level = self.voltage_rating_var.get()
        
        # 显示或隐藏3kV警告
        if voltage_level == "1.8/3kV":
            self.card_kv3_warning_label.grid()
        else:
            self.card_kv3_warning_label.grid_remove()
        
        self.update_predicted_code()

    def update_voltage_options_only(self):
        """仅更新电压选项，不设置默认值（用于加载规格时）"""
        category = self.category_var.get()
        if category in self.voltage_levels:
            self.voltage_rating_combo['values'] = self.voltage_levels[category]

    def on_card_category_change(self, event=None):
        """参数卡片中产品大类改变时的处理 - 遵循筛选规则"""
        category = self.category_var.get()
        if category in self.voltage_levels:
            self.voltage_rating_combo['values'] = self.voltage_levels[category]
            self.voltage_rating_var.set('')
            
            # 根据分类设置默认值和限制
            if category == "低压":
                self.voltage_rating_var.set("0.6/1kV")
                self.set_card_low_voltage_defaults()
            elif category == "中压":
                self.set_card_medium_voltage_defaults()
            elif category == "布线":
                self.voltage_rating_var.set("450/750V")
                self.set_card_wire_defaults()
            elif category == "光伏缆":
                self.voltage_rating_var.set("DC 1500V")
                self.set_card_pv_defaults()
            elif category == "控缆和仪表缆":
                self.voltage_rating_var.set("450/750V")
                self.set_card_control_defaults()
            elif category == "裸铜线":
                self.voltage_rating_var.set("N/A")
                self.set_card_bare_wire_defaults()
            elif category == "橡套电缆":
                self.voltage_rating_var.set("0.6/1kV")  # 橡套电缆默认电压等级
                self.set_card_rubber_cable_defaults()
            
            self.update_predicted_code()

    def set_card_low_voltage_defaults(self):
        """设置参数卡片低压电缆默认值"""
        self.conductor_card_var.set("CU")
        self.insulation_card_var.set("XLPE")
        self.shield_type_var.set("无")
        self.inner_sheath_var.set("无")
        self.armor_card_var.set("无")
        self.outer_sheath_var.set("PVC")
        self.is_fire_resistant_var.set(False)
        
        # 重置阻燃等级和特殊性能
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)

    def set_card_medium_voltage_defaults(self):
        """设置参数卡片中压电缆默认值"""
        self.voltage_rating_var.set("6/10kV")  # 设置中压默认电压
        self.conductor_card_var.set("CU")
        self.insulation_card_var.set("XLPE")
        self.shield_type_var.set("CTS")
        self.inner_sheath_var.set("无")
        self.armor_card_var.set("无")
        self.outer_sheath_var.set("PVC")
        self.is_fire_resistant_var.set(False)
        
        # 重置阻燃等级和特殊性能
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)

    def set_card_wire_defaults(self):
        """设置参数卡片布线电缆默认值"""
        self.conductor_card_var.set("CU")
        self.insulation_card_var.set("PVC")
        self.shield_type_var.set("无")
        self.inner_sheath_var.set("无")
        self.armor_card_var.set("无")
        self.outer_sheath_var.set("无")
        self.is_fire_resistant_var.set(False)
        
        # 重置阻燃等级和特殊性能
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)

    def set_card_pv_defaults(self):
        """设置参数卡片光伏电缆默认值和限制"""
        self.conductor_card_var.set("TAC")
        self.insulation_card_var.set("XLPO")
        self.shield_type_var.set("无")
        self.inner_sheath_var.set("无")
        self.armor_card_var.set("无")
        self.outer_sheath_var.set("XLPO")
        self.is_fire_resistant_var.set(False)
        
        # 重置阻燃等级和特殊性能
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)

    def set_card_control_defaults(self):
        """设置参数卡片控制/仪表电缆默认值"""
        self.conductor_card_var.set("CU")
        self.insulation_card_var.set("PVC")
        self.shield_type_var.set("无")
        self.inner_sheath_var.set("无")
        self.armor_card_var.set("无")
        self.outer_sheath_var.set("PVC")
        self.is_fire_resistant_var.set(False)
        
        # 重置阻燃等级和特殊性能
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)

    def set_card_bare_wire_defaults(self):
        """设置参数卡片裸铜线默认值和限制"""
        self.conductor_card_var.set("PABC")
        self.insulation_card_var.set("无")  # 修改：裸铜线绝缘默认为"无"
        self.shield_type_var.set("无")      # 修改：屏蔽默认为"无"
        self.inner_sheath_var.set("无")     # 修改：内护套默认为"无"
        self.armor_card_var.set("无")       # 修改：铠装默认为"无"
        self.outer_sheath_var.set("无")     # 修改：外护套默认为"无"
        self.is_fire_resistant_var.set(False)
        
        # 重置阻燃等级和特殊性能
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)

    def set_card_rubber_cable_defaults(self):
        """设置参数卡片橡套电缆默认值"""
        self.conductor_card_var.set("CU")
        self.insulation_card_var.set("EPR")
        self.shield_type_var.set("无")      # 橡套电缆默认无金属屏蔽层
        self.inner_sheath_var.set("无")
        self.armor_card_var.set("无")       # 橡套电缆默认无铠装
        self.outer_sheath_var.set("SE4")    # 橡套电缆专用护套材料
        self.is_fire_resistant_var.set(False)
        
        # 重置阻燃等级和特殊性能
        self.fire_rating_var.set("无")
        for var in self.special_performance_vars.values():
            var.set(False)

    def on_parameter_change(self, event=None):
        """参数变化时触发编码预测 - 实现用户指定的操作顺序"""
        # 操作顺序：核心参数设置 → 编码预测 → 处理已存在/新编码 → 结构绑定
        self.update_predicted_code()
        self.check_code_existence()

    def check_code_existence(self):
        """检查编码是否已存在，并相应处理"""
        try:
            # 收集当前参数
            special_performance = []
            fire_rating = self.fire_rating_var.get()
            if fire_rating and fire_rating != "无":
                special_performance.append(fire_rating)
            
            for option, var in self.special_performance_vars.items():
                if var.get():
                    special_performance.append(option)
            
            # 标准化内护套值
            inner_sheath_value = self.inner_sheath_var.get()
            if inner_sheath_value == "None":
                inner_sheath_value = "无"
            
            # 创建临时产品对象用于检查
            temp_product = CableProduct(
                category=self.category_var.get(),
                voltage_rating=self.voltage_rating_var.get(),
                conductor=self.conductor_card_var.get(),
                insulation=self.insulation_card_var.get(),
                shield_type=self.shield_type_var.get(),
                inner_sheath=inner_sheath_value,
                armor=self.armor_card_var.get(),
                outer_sheath=self.outer_sheath_var.get(),
                is_fire_resistant=self.is_fire_resistant_var.get(),
                special_performance=special_performance
            )
            
            # 检查参数完整性
            is_valid, missing_fields = temp_product.validate()
            if not is_valid:
                # 参数不完整，显示待完善状态
                self.predicted_code_var.set("CBL-SPEC-XXXXXX")
                self.code_status_var.set("待完善参数")
                self.status_label.config(fg="gray")
                return
            
            # 计算参数哈希
            param_hash = self.code_manager.calculate_param_hash(temp_product)
            
            # 检查数据库中是否已存在
            conn = sqlite3.connect(self.code_manager.db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT spec_id FROM product_specs WHERE param_hash = ?", (param_hash,))
            existing = cursor.fetchone()
            
            if existing:
                # 情况A：编码已存在
                spec_id = existing[0]
                self.predicted_code_var.set(spec_id)
                self.code_status_var.set("已存在")
                self.status_label.config(fg="blue")
                
                # 获取已有的型号和结构信息
                spec_info = self.code_manager.get_spec_by_id(spec_id)
                if spec_info:
                    spec = spec_info["spec"]
                    aliases = spec_info["aliases"]
                    
                    # 自动填充型号名称和产品描述
                    # 优先使用保存的产品型号，而不是别名
                    saved_product_model = spec[11] if spec[11] else ""  # 初始化变量
                    if spec[11]:  # product_model字段
                        self.auto_model_name_var.set(saved_product_model)
                        print(f"🔄 加载保存的产品型号: {saved_product_model}")
                    elif aliases:
                        # 如果没有产品型号，使用第一个别名作为备选
                        fallback_model = aliases[0][0]
                        self.auto_model_name_var.set(fallback_model)
                        print(f"🔄 使用别名作为备选产品型号: {fallback_model}")
                    
                    # 填充结构描述
                    if spec[12]:  # structure_string字段（注意：添加product_model后索引变化）
                        saved_structure = spec[12]
                        self.structure_string_var.set(saved_structure)
                        print(f"🔄 加载保存的结构字符串: {saved_structure}")
                    
                    # 显示所有关联型号（排除产品型号本身）
                    alias_names = [alias[0] for alias in aliases if alias[0] != saved_product_model]
                    self.confirmed_aliases_var.set(", ".join(alias_names))
                
                # 显示别名管理区域
                self.alias_management_frame.pack(fill=tk.X, pady=10)
                # 隐藏结构保存按钮
                self.structure_save_frame.pack_forget()
                
            else:
                # 情况B：编码不存在（新的参数组合）
                hash_suffix = param_hash[:12].upper()
                spec_id = f"CBL-SPEC-{hash_suffix}"
                
                self.predicted_code_var.set(spec_id)
                self.code_status_var.set("新建")
                self.status_label.config(fg="green")
                
                # 对于新建编码，自动生成型号和结构
                self.auto_generate_model_name()
                self.update_structure_string()
                
                # 清空确认别名
                self.confirmed_aliases_var.set("")
                
                # 显示结构保存按钮，隐藏别名管理区域
                self.structure_save_frame.pack(fill=tk.X, pady=5)
                self.alias_management_frame.pack_forget()
            
            conn.close()
            
            # 注意：不在这里调用update_structure_string()
            # 因为对于已存在的编码，我们已经加载了保存的结构
            # 对于新建编码，会在下面的条件中处理
            
            # 如果是新建，自动生成型号名称和结构
            if not existing:
                self.auto_generate_model_name()
                self.update_structure_string()
                
        except Exception as e:
            self.predicted_code_var.set("CBL-SPEC-ERROR")
            self.code_status_var.set("检查失败")
            self.status_label.config(fg="red")

    def update_predicted_code(self, event=None):
        """更新预测编码和结构字符串"""
        try:
            # 重置确认状态
            if hasattr(self, 'confirmed_product'):
                delattr(self, 'confirmed_product')
            if hasattr(self, 'confirmed_spec_id'):
                delattr(self, 'confirmed_spec_id')
            
            # 重置界面状态
            self.predicted_code_var.set("CBL-SPEC-XXXXXX")
            self.code_status_var.set("待确认")
            self.status_label.config(fg="orange")
            self.confirm_button.config(text="🔍 确认编码", state="normal")
            self.alias_management_frame.pack_forget()
            self.structure_save_frame.pack_forget()
            
            # 更新结构字符串
            self.update_structure_string()
            
            # 注意：不在这里调用auto_generate_model_name()
            # 因为这会覆盖用户搜索到的型号或已保存的型号
            # 只有在确认编码时，对于新建编码才调用自动生成
                
        except Exception as e:
            self.predicted_code_var.set("CBL-SPEC-XXXXXX")
            self.structure_string_var.set("结构生成失败")

    def save_parameter_card(self):
        """保存参数卡片 - 使用确认的编码"""
        try:
            # 检查是否已确认编码
            if not hasattr(self, 'confirmed_product') or not hasattr(self, 'confirmed_spec_id'):
                messagebox.showwarning("警告", "请先点击'确认编码'按钮！")
                return
            
            # 使用确认的产品对象
            product = self.confirmed_product
            spec_id = self.confirmed_spec_id
            
            # 更新描述
            product.description = self.description_var.get()
            
            # 收集型号别名（不包括产品型号本身）
            model_aliases = []
            
            # 获取用户填写的产品型号
            user_product_model = self.auto_model_name_var.get().strip()
            
            # 添加确认区域的别名（排除产品型号本身）
            confirmed_aliases = self.confirmed_aliases_var.get().strip()
            if confirmed_aliases:
                conf_aliases = [alias.strip() for alias in confirmed_aliases.split(',') if alias.strip()]
                for alias in conf_aliases:
                    # 只添加与产品型号不同的别名
                    if alias != user_product_model and alias not in model_aliases:
                        model_aliases.append(alias)
            
            # 保存到数据库
            conn = sqlite3.connect(self.code_manager.db_path)
            cursor = conn.cursor()
            
            # 检查规格是否已存在
            cursor.execute("SELECT spec_id FROM product_specs WHERE spec_id = ?", (spec_id,))
            existing = cursor.fetchone()
            
            if not existing:
                # 创建新规格
                param_hash = self.code_manager.calculate_param_hash(product)
                product_dict = product.to_dict()
                special_performance_json = json.dumps(product_dict['special_performance'], ensure_ascii=False)
                
                # 获取用户填写的产品型号和结构
                user_product_model = self.auto_model_name_var.get().strip()
                user_structure = self.structure_string_var.get().strip()
                
                cursor.execute('''
                    INSERT INTO product_specs (
                        spec_id, param_hash, category, voltage_rating, conductor, insulation,
                        shield_type, inner_sheath, armor, outer_sheath, is_fire_resistant,
                        special_performance, product_model, structure_string, created_date, modified_date, usage_count
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    spec_id, param_hash, product.category, product.voltage_rating, product.conductor,
                    product.insulation, product.shield_type, product.inner_sheath, product.armor,
                    product.outer_sheath, product.is_fire_resistant, special_performance_json,
                    user_product_model, user_structure, datetime.now().isoformat(), 
                    datetime.now().isoformat(), 1
                ))
            else:
                # 更新现有规格的产品型号和结构
                user_product_model = self.auto_model_name_var.get().strip()
                user_structure = self.structure_string_var.get().strip()
                
                cursor.execute('''
                    UPDATE product_specs 
                    SET product_model = ?, structure_string = ?, modified_date = ?, usage_count = usage_count + 1
                    WHERE spec_id = ?
                ''', (user_product_model, user_structure, datetime.now().isoformat(), spec_id))
            
            # 先删除该规格的所有现有别名映射
            cursor.execute("DELETE FROM alias_spec_mapping WHERE spec_id = ?", (spec_id,))
            
            # 添加新的型号别名映射（不包括产品型号本身）
            for alias in model_aliases:
                if alias.strip():
                    self.code_manager.add_alias_mapping(alias.strip(), spec_id, "用户保存", 1.0, cursor=cursor)
            
            conn.commit()
            conn.close()
            
            # 更新显示
            self.predicted_code_var.set(f"{spec_id} (已保存)")
            self.code_status_var.set("已保存")
            self.status_label.config(fg="green")
            self.refresh_product_list()
            
            # 构建成功消息
            success_msg = f"参数卡片已保存！\n规格编号：{spec_id}"
            if user_product_model:
                success_msg += f"\n产品型号：{user_product_model}"
            if user_structure:
                success_msg += f"\n结构字符串：{user_structure}"
            if model_aliases:
                success_msg += f"\n关联型号别名：{', '.join(model_aliases)}"
            
            messagebox.showinfo("成功", success_msg)
            
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")

    def reset_parameter_card(self):
        """重置参数卡片"""
        # 重置所有变量
        self.category_var.set('')
        self.voltage_rating_var.set('')
        self.conductor_card_var.set('')
        self.insulation_card_var.set('')
        self.shield_type_var.set('')
        self.inner_sheath_var.set('无')  # 使用"无"而不是空字符串
        self.armor_card_var.set('')
        self.outer_sheath_var.set('')
        self.is_fire_resistant_var.set(False)
        self.unified_search_var.set('')  # 更新为新的搜索变量
        self.model_binding_var.set('')   # 更新为新的绑定变量
        self.description_var.set('')
        
        # 重置预测编码区域
        self.predicted_code_var.set("CBL-SPEC-XXXXXX")
        self.code_status_var.set("待确认")
        self.status_label.config(fg="orange")
        self.structure_string_var.set("")
        self.auto_model_name_var.set("")
        self.confirmed_aliases_var.set("")
        
        # 重置阻燃等级
        self.fire_rating_var.set("无")
        
        # 重置特殊性能选择
        for var in self.special_performance_vars.values():
            var.set(False)
        
        # 隐藏3kV警告
        self.card_kv3_warning_label.grid_remove()
        
        # 隐藏别名管理区域和结构保存按钮
        self.alias_management_frame.pack_forget()
        self.structure_save_frame.pack_forget()
        
        # 重置确认按钮
        self.confirm_button.config(text="🔍 确认编码", state="normal")
        
        # 清除确认状态
        if hasattr(self, 'confirmed_product'):
            delattr(self, 'confirmed_product')
        if hasattr(self, 'confirmed_spec_id'):
            delattr(self, 'confirmed_spec_id')
        
        # 清空搜索结果
        self.clear_search_results()

    def create_product_folders(self):
        """为当前规格创建定额和技术规范文件夹"""
        code = self.predicted_code_var.get()
        if "XXXXXX" in code or "预测" not in code:
            messagebox.showwarning("警告", "请先保存参数卡片！")
            return
        
        # 提取实际编码
        actual_spec_id = code.split(" ")[0]
        
        try:
            # 创建定额文件夹
            quota_base = self.config.get("quota_folder", "")
            if quota_base and os.path.exists(quota_base):
                category = self.category_var.get()
                quota_category_path = os.path.join(quota_base, category)
                os.makedirs(quota_category_path, exist_ok=True)
                
                quota_spec_path = os.path.join(quota_category_path, actual_spec_id)
                os.makedirs(quota_spec_path, exist_ok=True)
                
                # 更新数据库中的路径
                self.code_manager.update_spec_paths(actual_spec_id, quota_path=quota_spec_path)
            
            # 创建技术规范文件夹
            spec_base = self.config.get("spec_folder", "")
            if spec_base and os.path.exists(spec_base):
                category = self.category_var.get()
                spec_category_path = os.path.join(spec_base, category)
                os.makedirs(spec_category_path, exist_ok=True)
                
                # 使用第一个型号别名作为文件夹名，如果没有则使用规格ID
                aliases_input = self.confirmed_aliases_var.get().strip()
                if aliases_input:
                    first_alias = aliases_input.split(',')[0].strip()
                    folder_name = first_alias
                else:
                    folder_name = actual_spec_id
                
                spec_folder_path = os.path.join(spec_category_path, folder_name)
                os.makedirs(spec_folder_path, exist_ok=True)
                
                # 更新数据库中的路径
                self.code_manager.update_spec_paths(actual_spec_id, spec_path=spec_folder_path)
            
            messagebox.showinfo("成功", f"已为规格 {actual_spec_id} 创建文件夹！")
            self.refresh_product_list()
            
        except Exception as e:
            messagebox.showerror("错误", f"创建文件夹失败：{str(e)}")

    def refresh_product_list(self):
        """刷新产品列表 - 适应新的三层数据模型和筛选功能"""
        try:
            # 清空现有数据
            for item in self.product_tree.get_children():
                self.product_tree.delete(item)
            
            # 获取所有规格
            specs_data = self.code_manager.get_all_specs()
            
            # 应用筛选
            filtered_specs = self.apply_filters_to_data(specs_data)
            
            for spec_info in filtered_specs:
                spec = spec_info["spec"]
                aliases = spec_info["aliases"]
                
                spec_id, category, voltage_rating, conductor, insulation, shield_type, \
                inner_sheath, armor, outer_sheath, is_fire_resistant, special_performance, \
                product_model, structure_string, quota_path, spec_path, created_date, modified_date, usage_count = spec
                
                # 格式化显示数据
                fire_text = "是" if is_fire_resistant else "否"
                created_short = created_date[:16] if created_date else ""
                
                # 使用保存的产品型号，如果没有则使用第一个别名作为备选
                all_alias_names = [alias[0] for alias in aliases]
                display_model = product_model if product_model else (all_alias_names[0] if all_alias_names else spec_id)
                
                # 获取关联型号（排除显示的产品型号，最多显示3个）
                alias_names = [alias[0] for alias in aliases if alias[0] != display_model][:3]
                alias_text = ", ".join(alias_names)
                if len([alias[0] for alias in aliases if alias[0] != display_model]) > 3:
                    alias_text += f" (+{len([alias[0] for alias in aliases if alias[0] != display_model])-3})"
                
                # 处理特殊性能显示
                special_performance_text = ""
                if special_performance:
                    try:
                        import json
                        special_list = json.loads(special_performance) if isinstance(special_performance, str) else special_performance
                        if special_list:
                            special_performance_text = ", ".join(special_list)
                    except:
                        special_performance_text = str(special_performance) if special_performance else ""
                
                self.product_tree.insert("", "end", values=(
                    spec_id, display_model, structure_string or "", category, voltage_rating, conductor, insulation, 
                    shield_type, armor, outer_sheath, fire_text, special_performance_text,
                    alias_text, usage_count, created_short
                ))
            
            # 更新状态栏
            total_count = len(specs_data)
            filtered_count = len(filtered_specs)
            if total_count == filtered_count:
                self.count_label.config(text=f"共 {total_count} 条记录")
                self.status_label.config(text="就绪")
            else:
                self.count_label.config(text=f"显示 {filtered_count} / {total_count} 条记录")
                self.status_label.config(text="已应用筛选")
                
        except Exception as e:
            messagebox.showerror("错误", f"刷新列表失败：{str(e)}")
            self.status_label.config(text="刷新失败")
    
    def sort_product_list(self, col):
        """按指定列排序产品列表"""
        try:
            # 如果点击同一列，则反转排序顺序
            if self.sort_column == col:
                self.sort_reverse = not self.sort_reverse
            else:
                self.sort_column = col
                self.sort_reverse = False
            
            # 获取所有数据
            data = []
            for item in self.product_tree.get_children():
                values = self.product_tree.item(item)['values']
                data.append((item, values))
            
            # 获取列索引
            col_index = self.all_columns.index(col)
            
            # 定义排序键函数
            def sort_key(item):
                value = item[1][col_index]
                # 处理数字列
                if col in ["使用次数"]:
                    try:
                        return int(value) if value else 0
                    except:
                        return 0
                # 处理文本列（忽略大小写）
                return str(value).lower() if value else ""
            
            # 排序
            data.sort(key=sort_key, reverse=self.sort_reverse)
            
            # 重新插入排序后的数据
            for index, (item, values) in enumerate(data):
                self.product_tree.move(item, '', index)
            
            # 更新列标题显示排序指示器
            for c in self.all_columns:
                if c == col:
                    indicator = " ▼" if self.sort_reverse else " ▲"
                    self.product_tree.heading(c, text=c + indicator)
                else:
                    self.product_tree.heading(c, text=c, 
                                            command=lambda c=c: self.sort_product_list(c))
            
        except Exception as e:
            messagebox.showerror("错误", f"排序失败：{str(e)}")

    def update_column_visibility(self):
        """更新列的可见性"""
        try:
            # 获取当前显示的列
            visible_columns = []
            for col in self.all_columns:
                if self.column_visibility[col].get():
                    visible_columns.append(col)
            
            # 更新Treeview显示的列
            self.product_tree["displaycolumns"] = visible_columns
            
            # 刷新数据以应用列变化
            if hasattr(self, 'product_tree') and self.product_tree.get_children():
                self.refresh_product_list()
                
        except Exception as e:
            print(f"更新列可见性失败: {str(e)}")

    def apply_filters_to_data(self, specs_data):
        """对数据应用筛选条件"""
        if not hasattr(self, 'filter_category_var'):
            return specs_data  # 如果筛选控件还未初始化，返回原数据
        
        filtered_data = []
        
        # 获取筛选条件
        category_filter = self.filter_category_var.get()
        conductor_filter = self.filter_conductor_var.get()
        insulation_filter = self.filter_insulation_var.get()
        fire_filter = self.filter_fire_resistant_var.get()
        flame_retardant_filter = getattr(self, 'filter_flame_retardant_var', None)
        shield_filter = getattr(self, 'filter_shield_var', None)
        armor_filter = getattr(self, 'filter_armor_var', None)
        search_text = self.search_filter_var.get().strip().upper()
        
        for spec_info in specs_data:
            spec = spec_info["spec"]
            aliases = spec_info["aliases"]
            
            spec_id, category, voltage_rating, conductor, insulation, shield_type, \
            inner_sheath, armor, outer_sheath, is_fire_resistant, special_performance, \
            product_model, structure_string, quota_path, spec_path, created_date, modified_date, usage_count = spec
            
            # 应用筛选条件
            if category_filter != "全部" and category != category_filter:
                continue
            
            if conductor_filter != "全部" and conductor != conductor_filter:
                continue
            
            if insulation_filter != "全部" and insulation != insulation_filter:
                continue
            
            if fire_filter != "全部":
                fire_text = "是" if is_fire_resistant else "否"
                if fire_text != fire_filter:
                    continue
            
            # 新增：阻燃等级筛选
            if flame_retardant_filter and flame_retardant_filter.get() != "全部":
                flame_grade = flame_retardant_filter.get()
                if special_performance:
                    try:
                        import json
                        special_list = json.loads(special_performance) if isinstance(special_performance, str) else special_performance
                        if flame_grade == "无":
                            # 检查是否没有阻燃等级
                            has_flame_grade = any(grade in special_list for grade in self.fire_resistant_options)
                            if has_flame_grade:
                                continue
                        else:
                            # 检查是否包含指定的阻燃等级
                            if flame_grade not in special_list:
                                continue
                    except:
                        if flame_grade != "无":
                            continue
                else:
                    if flame_grade != "无":
                        continue
            
            # 新增：屏蔽类型筛选
            if shield_filter and shield_filter.get() != "全部":
                if shield_type != shield_filter.get():
                    continue
            
            # 新增：铠装类型筛选
            if armor_filter and armor_filter.get() != "全部":
                if armor != armor_filter.get():
                    continue
            
            # 搜索文本筛选（在规格编号、产品型号、型号别名、结构中搜索）
            if search_text:
                search_targets = [
                    spec_id.upper(),
                    product_model.upper() if product_model else "",  # 产品型号搜索
                    voltage_rating.upper() if voltage_rating else "",
                    structure_string.upper() if structure_string else ""
                ]
                
                # 添加所有别名到搜索目标
                for alias in aliases:
                    search_targets.append(alias[0].upper())
                
                # 改进搜索逻辑：更精确的匹配
                found_match = False
                for target in search_targets:
                    if target:
                        # 精确匹配
                        if search_text == target:
                            found_match = True
                            break
                        # 检查是否为独立的型号部分（用分隔符分隔）
                        elif self.is_independent_model_match(search_text, target):
                            found_match = True
                            break
                
                if not found_match:
                    continue
            
            filtered_data.append(spec_info)
        
        return filtered_data

    def is_independent_model_match(self, search_text, target):
        """检查搜索文本是否作为独立的型号部分出现在目标字符串中"""
        # 定义分隔符
        separators = ['.', '-', '_', ' ']
        
        # 将目标字符串按分隔符分割成部分
        parts = [target]
        for sep in separators:
            new_parts = []
            for part in parts:
                new_parts.extend(part.split(sep))
            parts = new_parts
        
        # 检查搜索文本是否作为独立部分存在
        for part in parts:
            if part.strip() == search_text:
                return True
        
        return False

    def apply_filters(self, event=None):
        """应用筛选条件"""
        self.refresh_product_list()

    def clear_filters(self):
        """清空所有筛选条件"""
        if hasattr(self, 'filter_category_var'):
            self.filter_category_var.set("全部")
            self.filter_conductor_var.set("全部")
            self.filter_insulation_var.set("全部")
            self.filter_fire_resistant_var.set("全部")
            
        # 清空新增的筛选条件
        if hasattr(self, 'filter_flame_retardant_var'):
            self.filter_flame_retardant_var.set("全部")
        if hasattr(self, 'filter_shield_var'):
            self.filter_shield_var.set("全部")
        if hasattr(self, 'filter_armor_var'):
            self.filter_armor_var.set("全部")
            
        # 清空搜索框
        if hasattr(self, 'search_filter_var'):
            self.search_filter_var.set("")
            
        # 刷新列表
        self.refresh_product_list()

    def show_context_menu(self, event):
        """显示右键菜单"""
        # 选中右键点击的项目
        item = self.product_tree.identify_row(event.y)
        if item:
            self.product_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def edit_selected_product(self):
        """编辑选中的产品"""
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个产品！")
            return
        
        # 调用原有的编辑方法
        self.edit_product_card(None)

    def copy_spec_id(self):
        """复制规格编号到剪贴板"""
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个产品！")
            return
        
        item = self.product_tree.item(selected[0])
        spec_id = item['values'][0]
        
        # 复制到剪贴板
        self.root.clipboard_clear()
        self.root.clipboard_append(spec_id)
        self.status_label.config(text=f"已复制规格编号: {spec_id}")

    def delete_selected_product(self):
        """删除选中的产品"""
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个产品！")
            return
        
        item = self.product_tree.item(selected[0])
        spec_id = item['values'][0]
        main_model = item['values'][1]
        
        # 确认删除
        result = messagebox.askyesno(
            "确认删除", 
            f"确定要删除以下产品吗？\n\n规格编号: {spec_id}\n产品型号: {main_model}\n\n此操作不可撤销！",
            icon="warning"
        )
        
        if result:
            try:
                self.delete_product_from_database(spec_id)
                self.refresh_product_list()
                self.status_label.config(text=f"已删除产品: {spec_id}")
                messagebox.showinfo("删除成功", f"产品 {spec_id} 已成功删除")
            except Exception as e:
                messagebox.showerror("删除失败", f"删除产品失败：{str(e)}")
                self.status_label.config(text="删除失败")

    def delete_product_from_database(self, spec_id):
        """从数据库中删除产品及其相关数据"""
        conn = sqlite3.connect(self.code_manager.db_path)
        cursor = conn.cursor()
        
        try:
            # 删除型号别名映射
            cursor.execute("DELETE FROM alias_spec_mapping WHERE spec_id = ?", (spec_id,))
            
            # 删除产品规格
            cursor.execute("DELETE FROM product_specs WHERE spec_id = ?", (spec_id,))
            
            # 检查是否有孤立的别名（没有映射的别名）
            cursor.execute('''
                DELETE FROM model_aliases 
                WHERE alias_name NOT IN (
                    SELECT DISTINCT alias_name FROM alias_spec_mapping
                )
            ''')
            
            conn.commit()
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()

    def export_product_list(self):
        """导出产品列表到Excel文件"""
        try:
            import pandas as pd
            from tkinter import filedialog
            
            # 获取当前显示的数据
            data = []
            for item in self.product_tree.get_children():
                values = self.product_tree.item(item)['values']
                data.append(values)
            
            if not data:
                messagebox.showwarning("警告", "没有数据可导出！")
                return
            
            # 创建DataFrame
            columns = ["规格编号", "产品型号", "大类", "电压", "导体", "绝缘", "屏蔽", "铠装", "护套", "耐火", "关联型号", "使用次数", "创建时间"]
            df = pd.DataFrame(data, columns=columns)
            
            # 选择保存位置
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("所有文件", "*.*")],
                title="导出产品列表"
            )
            
            if filename:
                if filename.endswith('.csv'):
                    df.to_csv(filename, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(filename, index=False)
                
                self.status_label.config(text=f"已导出到: {filename}")
                messagebox.showinfo("导出成功", f"产品列表已导出到:\n{filename}")
                
        except ImportError:
            messagebox.showerror("错误", "导出功能需要安装pandas库\n请运行: pip install pandas openpyxl")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出失败：{str(e)}")
            self.status_label.config(text="导出失败")

    def edit_product_card(self, event):
        """双击编辑产品参数卡片 - 适应新的数据模型"""
        selected = self.product_tree.selection()
        if not selected:
            return
        
        item = self.product_tree.item(selected[0])
        spec_id = item['values'][0]
        
        try:
            # 查询规格详细信息
            spec_info = self.code_manager.get_spec_by_id(spec_id)
            
            if spec_info:
                spec = spec_info["spec"]
                aliases = spec_info["aliases"]
                
                # 填充参数卡片
                self.load_spec_to_card(spec)
                
                # 设置型号别名 - 使用确认别名变量
                alias_names = [alias[0] for alias in aliases]
                self.confirmed_aliases_var.set(", ".join(alias_names))
                
                messagebox.showinfo("编辑模式", f"已加载规格 {spec_id} 的参数到编辑界面\n请切换到'参数卡片编辑'标签页进行修改")
                
        except Exception as e:
            messagebox.showerror("错误", f"加载规格信息失败：{str(e)}")

    def open_product_quota_folder(self):
        """打开选中规格的定额文件夹 - 自动创建文件夹版"""
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个规格！")
            return
        
        item = self.product_tree.item(selected[0])
        spec_id = item['values'][0]
        product_model = item['values'][1]  # 产品型号
        category = item['values'][3]  # 产品大类（修复：索引3，结构描述是索引2）
        
        try:
            quota_base = self.config.get("quota_folder", "")
            if not quota_base or not os.path.exists(quota_base):
                messagebox.showwarning("警告", "请先在设置中配置定额文件夹路径！")
                return
            
            # 1. 创建产品大类文件夹（如果不存在）
            category_folder = os.path.join(quota_base, category)
            os.makedirs(category_folder, exist_ok=True)
            
            # 2. 查询规格的定额路径
            spec_info = self.code_manager.get_spec_by_id(spec_id)
            target_folder = category_folder  # 默认目标文件夹
            
            if spec_info and spec_info["spec"][13] and os.path.exists(spec_info["spec"][13]):
                # 情况A：已有专属定额文件夹
                target_folder = spec_info["spec"][13]
                messagebox.showinfo("找到定额", f"打开型号 {product_model} 的专属定额文件夹")
            else:
                # 情况B：没有专属文件夹，创建新的产品型号文件夹
                # 使用产品型号命名，如果没有产品型号则使用规格ID作为备选
                folder_name = product_model if product_model else spec_id
                spec_folder = os.path.join(category_folder, folder_name)
                os.makedirs(spec_folder, exist_ok=True)
                target_folder = spec_folder
                
                # 更新数据库中的路径记录
                self.code_manager.update_spec_paths(spec_id, quota_path=spec_folder)
                
                messagebox.showinfo("创建定额文件夹", 
                                   f"已为型号 {product_model} 创建专属定额文件夹\n"
                                   f"规格：{spec_id}\n"
                                   f"路径：{category}/{folder_name}")
            
            # 3. 打开目标文件夹
            os.startfile(target_folder)
                
        except Exception as e:
            messagebox.showerror("错误", f"打开文件夹失败：{str(e)}")

    def open_product_spec_folder(self):
        """打开选中规格的技术规范文件夹 - 自动创建文件夹版"""
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个规格！")
            return
        
        item = self.product_tree.item(selected[0])
        spec_id = item['values'][0]
        category = item['values'][3]  # 产品大类（修复：索引3，结构描述是索引2）
        product_model = item['values'][1]  # 产品型号（使用产品型号，不是关联型号）
        
        try:
            spec_base = self.config.get("spec_folder", "")
            if not spec_base or not os.path.exists(spec_base):
                messagebox.showwarning("警告", "请先在设置中配置技术规范文件夹路径！")
                return
            
            # 1. 创建产品大类文件夹（如果不存在）
            category_folder = os.path.join(spec_base, category)
            os.makedirs(category_folder, exist_ok=True)
            
            # 2. 查询规格的技术规范路径
            spec_info = self.code_manager.get_spec_by_id(spec_id)
            target_folder = category_folder  # 默认目标文件夹
            
            if spec_info and spec_info["spec"][14] and os.path.exists(spec_info["spec"][14]):
                # 情况A：已有专属技术规范文件夹
                target_folder = spec_info["spec"][14]
                messagebox.showinfo("找到规范", f"打开规格 {spec_id} 的专属技术规范文件夹")
            else:
                # 情况B：没有专属文件夹，使用产品型号创建文件夹
                folder_name = product_model if product_model and product_model.strip() else spec_id
                
                # 检查是否已存在以产品型号命名的文件夹
                if folder_name != spec_id:
                    model_folder = os.path.join(category_folder, folder_name)
                    if os.path.exists(model_folder):
                        target_folder = model_folder
                        messagebox.showinfo("找到规范", f"找到型号 {folder_name} 的技术规范文件夹")
                    else:
                        # 创建以产品型号命名的文件夹
                        os.makedirs(model_folder, exist_ok=True)
                        target_folder = model_folder
                        
                        # 更新数据库中的路径记录
                        self.code_manager.update_spec_paths(spec_id, spec_path=model_folder)
                        
                        messagebox.showinfo("创建技术规范文件夹", 
                                           f"已为型号 {folder_name} 创建技术规范文件夹\n"
                                           f"规格：{spec_id}\n"
                                           f"路径：{category}/{folder_name}")
                else:
                    # 使用规格ID创建文件夹
                    spec_folder = os.path.join(category_folder, spec_id)
                    os.makedirs(spec_folder, exist_ok=True)
                    target_folder = spec_folder
                    
                    # 更新数据库中的路径记录
                    self.code_manager.update_spec_paths(spec_id, spec_path=spec_folder)
                    
                    messagebox.showinfo("创建技术规范文件夹", 
                                       f"已为规格 {spec_id} 创建专属技术规范文件夹\n"
                                       f"路径：{category}/{spec_id}")
            
            # 3. 打开目标文件夹
            os.startfile(target_folder)
                
        except Exception as e:
            messagebox.showerror("错误", f"打开文件夹失败：{str(e)}")
    
    def refresh_usage_count(self):
        """刷新产品使用次数统计"""
        try:
            # 确认操作
            response = messagebox.askyesno(
                "确认刷新", 
                "此操作将：\n"
                "1. 清空所有产品的现有使用次数\n"
                "2. 根据项目清单重新统计使用次数\n"
                "3. 统计规则：每个型号在不同项目中出现计数+1\n\n"
                "确认继续？",
                icon='question'
            )
            
            if not response:
                return
            
            # 显示进度提示
            self.status_label.config(text="正在统计使用次数...")
            self.root.update()
            
            from collections import defaultdict
            
            # 1. 加载项目清单数据
            project_lists = self.config.get("project_lists", {})
            
            if not project_lists:
                messagebox.showinfo("提示", "没有找到项目清单数据")
                self.status_label.config(text="就绪")
                return
            
            # 2. 统计每个型号在不同项目中的使用次数
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
            
            # 计算每个型号的使用次数
            model_usage_count = {}
            for model, projects in model_projects.items():
                model_usage_count[model] = len(projects)
            
            # 3. 清空现有的使用次数
            conn = sqlite3.connect(self.code_manager.db_path)
            cursor = conn.cursor()
            
            cursor.execute("UPDATE product_specs SET usage_count = 0")
            conn.commit()
            
            # 4. 更新使用次数
            updated_specs = set()
            unmatched_count = 0
            
            for model, count in model_usage_count.items():
                # 查找匹配的产品规格
                matching_specs = []
                
                # 匹配 product_model
                cursor.execute("""
                    SELECT spec_id FROM product_specs 
                    WHERE product_model = ?
                """, (model,))
                
                result = cursor.fetchone()
                if result:
                    matching_specs.append(result[0])
                
                # 匹配 alias_name
                cursor.execute("""
                    SELECT DISTINCT spec_id FROM alias_spec_mapping 
                    WHERE alias_name = ?
                """, (model,))
                
                for row in cursor.fetchall():
                    matching_specs.append(row[0])
                
                # 更新使用次数
                if matching_specs:
                    for spec_id in matching_specs:
                        cursor.execute("""
                            UPDATE product_specs 
                            SET usage_count = usage_count + ?
                            WHERE spec_id = ?
                        """, (count, spec_id))
                        updated_specs.add(spec_id)
                else:
                    unmatched_count += 1
            
            conn.commit()
            conn.close()
            
            # 5. 刷新显示
            self.refresh_product_list()
            
            # 6. 显示结果
            messagebox.showinfo(
                "刷新完成",
                f"使用次数统计完成！\n\n"
                f"项目总数：{len(project_lists)}\n"
                f"不同型号数：{len(model_usage_count)}\n"
                f"已匹配产品：{len(updated_specs)}\n"
                f"未匹配型号：{unmatched_count}\n\n"
                f"提示：未匹配的型号可能需要创建新的产品规格"
            )
            
            self.status_label.config(text="使用次数已更新")
            
        except Exception as e:
            messagebox.showerror("错误", f"刷新使用次数失败：{str(e)}")
            self.status_label.config(text="刷新失败")
            import traceback
            traceback.print_exc()

    def create_intelligent_parser_interface(self):
        """创建智能清单解析界面"""
        # 标题
        title_label = tk.Label(self.intelligent_frame, text="智能清单解析与匹配", 
                              font=("Microsoft YaHei", 16, "bold"), bg="#f0f0f0")
        title_label.pack(pady=10)
        
        # 创建主容器 - 使用垂直布局
        main_container = tk.Frame(self.intelligent_frame, bg="#f0f0f0")
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 上半部分：输入区域
        input_section = ttk.LabelFrame(main_container, text="📝 文本输入", padding=10)
        input_section.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(input_section, text="输入电缆描述文本 (每行一个):", 
                font=("Microsoft YaHei", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        # 文本输入框容器
        text_container = tk.Frame(input_section)
        text_container.pack(fill=tk.BOTH, expand=True)
        
        # 行号显示区域
        line_number_frame = tk.Frame(text_container, width=40, bg="#f5f5f5")
        line_number_frame.pack(side="left", fill="y")
        line_number_frame.pack_propagate(False)
        
        self.line_numbers = tk.Text(line_number_frame, width=4, padx=3, takefocus=0,
                                   border=0, state='disabled', wrap='none',
                                   font=("Microsoft YaHei", 10), bg="#f5f5f5", fg="#666")
        self.line_numbers.pack(fill="both", expand=True)
        
        # 主文本输入框
        text_input_frame = tk.Frame(text_container)
        text_input_frame.pack(side="left", fill="both", expand=True)
        
        self.input_text = tk.Text(text_input_frame, height=8, font=("Microsoft YaHei", 10), wrap=tk.WORD)
        input_scrollbar = ttk.Scrollbar(text_input_frame, orient="vertical", command=self.sync_scroll)
        self.input_text.configure(yscrollcommand=input_scrollbar.set)
        
        self.input_text.pack(side="left", fill="both", expand=True)
        input_scrollbar.pack(side="right", fill="y")
        
        # 绑定文本变化事件
        self.input_text.bind('<KeyRelease>', self.update_line_numbers)
        self.input_text.bind('<Button-1>', self.update_line_numbers)
        self.input_text.bind('<MouseWheel>', self.on_text_scroll)
        
        # 示例文本
        example_text = """YJV22 8.7/15kV 3x240+1x120 阻燃ZA级 防鼠防白蚁
YJLV 0.6/1kV 4x95 低烟无卤
铜芯交联聚乙烯绝缘钢带铠装聚氯乙烯护套电力电缆 10千伏 3x185
NH-YJV 6/10kV 3x120 耐火电缆
WDZ-YJV 0.6/1kV 5x16 无卤低烟阻燃"""
        
        self.input_text.insert("1.0", example_text)
        
        # 操作按钮
        button_frame = tk.Frame(input_section)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Button(button_frame, text="🚀 开始解析", command=self.start_intelligent_parsing,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 11, "bold"), 
                 width=12, height=2).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="🔄 清空输入", command=self.clear_input_text,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 11, "bold"), 
                 width=12, height=2).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="📋 复制结果", command=self.copy_parsing_results,
                 bg="#2196F3", fg="white", font=("Microsoft YaHei", 11, "bold"), 
                 width=12, height=2).pack(side=tk.LEFT, padx=5)
        
        # 下半部分：结果显示区域
        results_section = ttk.LabelFrame(main_container, text="🎯 解析结果", padding=10)
        results_section.pack(fill=tk.BOTH, expand=True)
        
        # 创建结果区域的水平布局
        results_container = tk.Frame(results_section)
        results_container.pack(fill=tk.BOTH, expand=True)
        
        # 左侧：结果表格
        table_frame = tk.Frame(results_container)
        table_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 结果表格
        columns = ("行号", "电压等级", "报价型号", "规格", "结构描述", "备注")
        self.results_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=12)
        
        # 设置列标题和宽度
        column_widths = {
            "行号": 50,
            "电压等级": 80,
            "报价型号": 100,
            "规格": 80,
            "结构描述": 150,
            "备注": 80
        }
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=column_widths.get(col, 100))
        
        # 滚动条
        results_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=results_scrollbar.set)
        
        self.results_tree.pack(side="left", fill="both", expand=True)
        results_scrollbar.pack(side="right", fill="y")
        
        # 绑定选择事件
        self.results_tree.bind('<<TreeviewSelect>>', self.on_result_select)
        self.results_tree.bind('<Double-1>', self.on_result_double_click)
        
        # 配置颜色标签
        self.results_tree.tag_configure("high_confidence", background="#E8F5E8")
        self.results_tree.tag_configure("medium_confidence", background="#FFF3E0")
        self.results_tree.tag_configure("low_confidence", background="#FFEBEE")
        self.results_tree.tag_configure("selected_line", background="#E3F2FD")
        
        # 右侧：编辑和统计区域
        right_panel = tk.Frame(results_container, width=300)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y)
        right_panel.pack_propagate(False)
        
        # 编辑区域
        edit_frame = ttk.LabelFrame(right_panel, text="✏️ 人工校验", padding=10)
        edit_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 编辑控件 - 使用网格布局
        # 第一行：电压等级
        tk.Label(edit_frame, text="电压等级:", font=("Microsoft YaHei", 10)).grid(row=0, column=0, sticky=tk.W, pady=2)
        self.edit_voltage_var = tk.StringVar()
        self.edit_voltage_entry = tk.Entry(edit_frame, textvariable=self.edit_voltage_var, 
                                          font=("Microsoft YaHei", 10), width=20)
        self.edit_voltage_entry.grid(row=0, column=1, pady=2, padx=(5, 0), sticky=tk.W+tk.E)
        
        # 第二行：报价型号
        tk.Label(edit_frame, text="报价型号:", font=("Microsoft YaHei", 10)).grid(row=1, column=0, sticky=tk.W, pady=2)
        self.edit_model_var = tk.StringVar()
        self.edit_model_entry = tk.Entry(edit_frame, textvariable=self.edit_model_var, 
                                        font=("Microsoft YaHei", 10), width=20)
        self.edit_model_entry.grid(row=1, column=1, pady=2, padx=(5, 0), sticky=tk.W+tk.E)
        
        # 绑定型号变化事件
        self.edit_model_var.trace_add('write', self.on_model_change)
        
        # 第三行：规格
        tk.Label(edit_frame, text="规格:", font=("Microsoft YaHei", 10)).grid(row=2, column=0, sticky=tk.W, pady=2)
        self.edit_spec_var = tk.StringVar()
        self.edit_spec_entry = tk.Entry(edit_frame, textvariable=self.edit_spec_var, 
                                       font=("Microsoft YaHei", 10), width=20)
        self.edit_spec_entry.grid(row=2, column=1, pady=2, padx=(5, 0), sticky=tk.W+tk.E)
        
        # 第四行：结构描述
        tk.Label(edit_frame, text="结构描述:", font=("Microsoft YaHei", 10)).grid(row=3, column=0, sticky=tk.W, pady=2)
        self.edit_structure_var = tk.StringVar()
        self.edit_structure_entry = tk.Entry(edit_frame, textvariable=self.edit_structure_var, 
                                            font=("Microsoft YaHei", 10), width=20)
        self.edit_structure_entry.grid(row=3, column=1, pady=2, padx=(5, 0), sticky=tk.W+tk.E)
        
        # 第五行：备注
        tk.Label(edit_frame, text="备注:", font=("Microsoft YaHei", 10)).grid(row=4, column=0, sticky=tk.W, pady=2)
        self.edit_remarks_var = tk.StringVar()
        self.edit_remarks_entry = tk.Entry(edit_frame, textvariable=self.edit_remarks_var, 
                                          font=("Microsoft YaHei", 10), width=20)
        self.edit_remarks_entry.grid(row=4, column=1, pady=2, padx=(5, 0), sticky=tk.W+tk.E)
        
        # 配置列权重
        edit_frame.columnconfigure(1, weight=1)
        
        # 按钮行
        button_row = tk.Frame(edit_frame)
        button_row.grid(row=5, column=0, columnspan=2, pady=(10, 0), sticky=tk.W+tk.E)
        
        tk.Button(button_row, text="✅ 更新", command=self.update_selected_result,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 9, "bold"), 
                 width=10).pack(side=tk.LEFT, padx=(0, 5))
        
        tk.Button(button_row, text="🔄 重置", command=self.reset_edit_fields,
                 bg="#FF9800", fg="white", font=("Microsoft YaHei", 9, "bold"), 
                 width=10).pack(side=tk.LEFT)
        
        # 统计信息区域
        stats_frame = ttk.LabelFrame(right_panel, text="📊 统计信息", padding=10)
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.stats_label = tk.Label(stats_frame, text="等待解析...", 
                                   font=("Microsoft YaHei", 10), fg="gray", wraplength=280, justify=tk.LEFT)
        self.stats_label.pack(fill=tk.X)
        
        # 使用说明区域
        help_frame = ttk.LabelFrame(right_panel, text="💡 使用说明", padding=10)
        help_frame.pack(fill=tk.BOTH, expand=True)
        
        help_text = """操作指南：

• 单击结果行：高亮对应输入文本
• 双击结果行：进入编辑模式
• 修改型号：结构自动更新
• 点击更新：保存修改
• 点击重置：恢复原值

颜色说明：
🟢 绿色：高置信度(≥90%)
🟡 橙色：中置信度(70-90%)
🔴 红色：低置信度(<70%)"""
        
        help_label = tk.Label(help_frame, text=help_text, font=("Microsoft YaHei", 9), 
                             fg="#666", justify=tk.LEFT, wraplength=280)
        help_label.pack(fill=tk.BOTH, expand=True)
        
        # 存储解析结果和当前选中项
        self.parsing_results = []
        self.selected_result_index = -1
        
        # 初始化行号
        self.update_line_numbers()
        
        # 配置颜色标签
        self.results_tree.tag_configure("high_confidence", background="#E8F5E8")
        self.results_tree.tag_configure("medium_confidence", background="#FFF3E0")
        self.results_tree.tag_configure("low_confidence", background="#FFEBEE")
        # 存储解析结果和当前选中项
        self.parsing_results = []
        self.selected_result_index = -1
        
        # 初始化行号
        self.update_line_numbers()
    
    def start_intelligent_parsing(self):
        """开始智能解析"""
        try:
            # 获取输入文本
            text_content = self.input_text.get("1.0", tk.END).strip()
            if not text_content:
                messagebox.showwarning("警告", "请输入要解析的文本！")
                return
            
            texts = [line.strip() for line in text_content.split('\n') if line.strip()]
            
            # 清空现有结果
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            self.parsing_results = []
            
            # 解析每行文本，记录行号
            line_number = 1
            for i, line in enumerate(text_content.split('\n')):
                if line.strip():  # 只处理非空行
                    result = self.parse_single_text(line.strip())
                    result['line_number'] = line_number
                    result['original_line_index'] = i  # 原始行索引（包含空行）
                    self.parsing_results.append(result)
                    
                    # 添加到表格
                    self.add_result_to_tree(result)
                    line_number += 1
            
            # 更新统计信息
            self.update_parsing_stats()
            
            messagebox.showinfo("完成", f"解析完成！共处理 {len(self.parsing_results)} 条记录。")
            
        except Exception as e:
            messagebox.showerror("错误", f"解析失败：{str(e)}")
    
    def update_line_numbers(self, event=None):
        """更新行号显示"""
        try:
            # 获取文本内容
            content = self.input_text.get("1.0", tk.END)
            lines = content.split('\n')
            
            # 生成行号
            line_numbers = []
            for i in range(len(lines)):
                if i < len(lines) - 1 or lines[i].strip():  # 最后一行如果为空则不显示行号
                    line_numbers.append(str(i + 1))
                else:
                    line_numbers.append("")
            
            # 更新行号显示
            self.line_numbers.config(state='normal')
            self.line_numbers.delete("1.0", tk.END)
            self.line_numbers.insert("1.0", '\n'.join(line_numbers))
            self.line_numbers.config(state='disabled')
            
        except Exception as e:
            print(f"更新行号错误: {str(e)}")
    
    def sync_scroll(self, *args):
        """同步滚动行号和文本"""
        self.input_text.yview(*args)
        self.line_numbers.yview(*args)
    
    def on_text_scroll(self, event):
        """处理鼠标滚轮事件"""
        self.line_numbers.yview_scroll(int(-1*(event.delta/120)), "units")
        return "break"
    
    def on_result_select(self, event):
        """当选择结果行时，高亮对应的输入文本行"""
        try:
            selection = self.results_tree.selection()
            if not selection:
                self.clear_text_highlight()
                self.clear_edit_fields()
                return
            
            # 获取选中项的数据
            item = self.results_tree.item(selection[0])
            values = item['values']
            
            if not values:
                return
            
            line_number = int(values[0])  # 第一列是行号
            self.selected_result_index = line_number - 1
            
            # 高亮对应的文本行
            self.highlight_text_line(line_number)
            
            # 填充编辑字段
            if self.selected_result_index < len(self.parsing_results):
                result = self.parsing_results[self.selected_result_index]
                self.fill_edit_fields(result)
            
        except Exception as e:
            print(f"选择结果错误: {str(e)}")
    
    def on_result_double_click(self, event):
        """双击结果行时进入编辑模式"""
        try:
            # 确保有选中的项
            selection = self.results_tree.selection()
            if not selection:
                return
            
            # 获取选中项的数据
            item = self.results_tree.item(selection[0])
            values = item['values']
            
            if not values:
                return
            
            line_number = int(values[0])  # 第一列是行号
            self.selected_result_index = line_number - 1
            
            # 高亮对应的文本行
            self.highlight_text_line(line_number)
            
            # 填充编辑字段
            if self.selected_result_index < len(self.parsing_results):
                result = self.parsing_results[self.selected_result_index]
                self.fill_edit_fields(result)
            
            # 聚焦到第一个编辑字段并选中内容
            self.edit_voltage_entry.focus_set()
            self.edit_voltage_entry.select_range(0, tk.END)
            
            # 显示提示信息
            messagebox.showinfo("编辑模式", f"正在编辑第 {line_number} 行的解析结果\n\n可以修改各个字段，然后点击'✅ 更新'保存修改")
            
        except Exception as e:
            print(f"双击编辑错误: {str(e)}")
            messagebox.showerror("错误", f"进入编辑模式失败：{str(e)}")
    
    def highlight_text_line(self, line_number):
        """高亮指定行的文本"""
        try:
            # 清除之前的高亮
            self.clear_text_highlight()
            
            # 高亮指定行
            start_index = f"{line_number}.0"
            end_index = f"{line_number}.end"
            
            self.input_text.tag_add("highlight", start_index, end_index)
            self.input_text.tag_config("highlight", background="#FFE082", foreground="#E65100", relief="raised", borderwidth=1)
            
            # 滚动到该行
            self.input_text.see(start_index)
            
            print(f"高亮第 {line_number} 行文本")
            
        except Exception as e:
            print(f"高亮文本行错误: {str(e)}")
    
    def clear_text_highlight(self):
        """清除文本高亮"""
        try:
            self.input_text.tag_remove("highlight", "1.0", tk.END)
        except Exception as e:
            print(f"清除高亮错误: {str(e)}")
    
    def fill_edit_fields(self, result):
        """填充编辑字段"""
        try:
            self.edit_voltage_var.set(result.get('voltage', ''))
            self.edit_model_var.set(result.get('model', ''))
            self.edit_spec_var.set(result.get('specification', ''))
            self.edit_structure_var.set(result.get('structure', ''))
            self.edit_remarks_var.set(result.get('remarks', ''))
        except Exception as e:
            print(f"填充编辑字段错误: {str(e)}")
    
    def clear_edit_fields(self):
        """清空编辑字段"""
        self.edit_voltage_var.set('')
        self.edit_model_var.set('')
        self.edit_spec_var.set('')
        self.edit_structure_var.set('')
        self.edit_remarks_var.set('')
        self.selected_result_index = -1
    
    def on_model_change(self, *args):
        """当型号改变时，自动更新结构描述"""
        try:
            model = self.edit_model_var.get().strip()
            if not model:
                return
            
            # 从系统的参数卡片中获取对应的结构描述
            structure = self.get_structure_by_model(model)
            if structure:
                current_structure = self.edit_structure_var.get().strip()
                # 只有当前结构为空或者与推断结构不同时才更新
                if not current_structure or current_structure != structure:
                    self.edit_structure_var.set(structure)
                    print(f"型号 {model} 自动更新结构为: {structure}")
                
        except Exception as e:
            print(f"型号变化处理错误: {str(e)}")
    
    def get_structure_by_model(self, model):
        """根据型号获取结构描述"""
        try:
            # 特殊型号直接返回
            if model == 'PABC':
                return 'PABC'
            elif model == 'H1Z2Z2.K':
                return 'TAC/XLPO/XLPO'
            
            # 首先尝试从数据库中查找
            search_results = self.code_manager.search_by_alias(model)
            
            if search_results:
                best_match = search_results[0]
                spec_data = best_match.get('spec_data')
                
                if spec_data and len(spec_data) > 9 and spec_data[9]:
                    return spec_data[9]  # structure_string字段
            
            # 如果数据库中没有，尝试根据型号推断结构
            return self.infer_structure_by_model(model)
            
        except Exception as e:
            print(f"获取结构描述错误: {str(e)}")
            return ""
    
    def infer_structure_by_model(self, model):
        """根据型号推断结构描述"""
        try:
            # 解析型号组件
            parts = model.split('.')
            base_model = parts[0] if parts else model
            voltage_suffix = parts[1] if len(parts) > 1 else ''
            
            # 基础结构推断
            structure_parts = []
            
            # 导体材质
            if 'YJLV' in base_model or 'AL' in base_model:
                structure_parts.append('AL')
            elif 'PABC' in base_model:
                return 'PABC'  # 裸铜线特殊处理
            elif 'H1Z2Z2' in base_model:
                return 'TAC/XLPO/XLPO'  # 光伏电缆特殊处理
            else:
                structure_parts.append('CU')
            
            # 绝缘材料
            if 'NH' in base_model:
                if 'YJ' in base_model:
                    structure_parts.append('MT/XLPE')
                else:
                    structure_parts.append('MT/PVC')
            elif 'YJ' in base_model:
                structure_parts.append('XLPE')
            elif 'BV' in base_model:
                structure_parts.append('PVC')
            else:
                structure_parts.append('PVC')
            
            # 屏蔽（中压电缆）
            if voltage_suffix == 'MV':
                structure_parts.append('CTS')
            
            # 内护套（铠装电缆）
            has_armor = any(armor in base_model for armor in ['22', '32', '62', '72', '23'])
            if has_armor:
                if '23' in base_model:
                    structure_parts.append('HDPE')
                else:
                    structure_parts.append('PVC')
            
            # 铠装
            if '22' in base_model or '23' in base_model:
                structure_parts.append('STA')
            elif '32' in base_model:
                structure_parts.append('SWA')
            elif '62' in base_model:
                structure_parts.append('SSTA')
            elif '72' in base_model:
                structure_parts.append('AWA')
            
            # 外护套
            if 'BV' in base_model and not has_armor:
                # BV线缆通常没有护套
                pass
            elif '23' in base_model:
                structure_parts.append('HDPE')
            elif 'WDZ' in base_model:
                structure_parts.append('LSZH')
            else:
                structure_parts.append('PVC')
            
            return '/'.join(structure_parts)
            
        except Exception as e:
            print(f"推断结构描述错误: {str(e)}")
            return ""
    
    def update_selected_result(self):
        """更新选中的解析结果"""
        try:
            if self.selected_result_index < 0 or self.selected_result_index >= len(self.parsing_results):
                messagebox.showwarning("警告", "请先选择要更新的结果行！")
                return
            
            # 获取编辑字段的值
            voltage = self.edit_voltage_var.get().strip()
            model = self.edit_model_var.get().strip()
            specification = self.edit_spec_var.get().strip()
            structure = self.edit_structure_var.get().strip()
            remarks = self.edit_remarks_var.get().strip()
            
            # 验证必填字段
            if not voltage or not model:
                messagebox.showwarning("警告", "电压等级和型号不能为空！")
                return
            
            # 更新解析结果
            result = self.parsing_results[self.selected_result_index]
            old_values = {
                'voltage': result.get('voltage', ''),
                'model': result.get('model', ''),
                'specification': result.get('specification', ''),
                'structure': result.get('structure', ''),
                'remarks': result.get('remarks', '')
            }
            
            result['voltage'] = voltage
            result['model'] = model
            result['specification'] = specification
            result['structure'] = structure
            result['remarks'] = remarks
            result['confidence'] = 1.0  # 人工校验后置信度为100%
            
            # 更新表格显示
            self.refresh_results_tree()
            
            # 重新选中该行
            items = self.results_tree.get_children()
            if self.selected_result_index < len(items):
                self.results_tree.selection_set(items[self.selected_result_index])
                self.results_tree.focus(items[self.selected_result_index])
            
            # 显示更新信息
            changes = []
            for key in ['voltage', 'model', 'specification', 'structure', 'remarks']:
                if old_values[key] != result[key]:
                    changes.append(f"{key}: '{old_values[key]}' → '{result[key]}'")
            
            if changes:
                change_text = "\n".join(changes)
                messagebox.showinfo("更新成功", f"第 {result.get('line_number', 0)} 行结果已更新！\n\n变更内容：\n{change_text}")
            else:
                messagebox.showinfo("提示", "没有检测到任何变更。")
            
        except Exception as e:
            messagebox.showerror("错误", f"更新失败：{str(e)}")
    
    def reset_edit_fields(self):
        """重置编辑字段为原始解析结果"""
        try:
            if self.selected_result_index < 0 or self.selected_result_index >= len(self.parsing_results):
                return
            
            result = self.parsing_results[self.selected_result_index]
            self.fill_edit_fields(result)
            
        except Exception as e:
            print(f"重置编辑字段错误: {str(e)}")
    
    def refresh_results_tree(self):
        """刷新结果表格显示"""
        try:
            # 清空现有结果
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            # 重新添加所有结果
            for result in self.parsing_results:
                self.add_result_to_tree(result)
            
            # 更新统计信息
            self.update_parsing_stats()
            
        except Exception as e:
            print(f"刷新结果表格错误: {str(e)}")
    
    def add_result_to_tree(self, result):
        """添加解析结果到表格"""
        try:
            line_number = result.get('line_number', 0)
            voltage = result.get('voltage', '')
            model = result.get('model', '')
            specification = result.get('specification', '')
            structure = result.get('structure', '')
            remarks = result.get('remarks', '')
            confidence = result.get('confidence', 0.0)
            match_source = result.get('match_source', '直接识别')
            
            # 根据置信度和匹配来源选择标签
            if confidence >= 0.85:
                tag = "high_confidence"  # 绿色 - 直接识别成功
            elif confidence >= 0.55 or match_source == '结构匹配':
                tag = "medium_confidence"  # 橙色 - 结构匹配或中等置信度
            else:
                tag = "low_confidence"  # 红色 - 低置信度或无法识别
            
            # 插入到表格
            self.results_tree.insert("", "end", 
                                   values=(line_number, voltage, model, specification, structure, remarks),
                                   tags=(tag,))
            
        except Exception as e:
            print(f"添加结果到表格错误: {str(e)}")
    
    def update_parsing_stats(self):
        """更新解析统计信息"""
        try:
            if not self.parsing_results:
                self.stats_label.config(text="等待解析...")
                return
            
            total = len(self.parsing_results)
            high_confidence = sum(1 for r in self.parsing_results if r.get('confidence', 0) >= 0.85)
            medium_confidence = sum(1 for r in self.parsing_results if 0.55 <= r.get('confidence', 0) < 0.85 or r.get('match_source') == '结构匹配')
            low_confidence = sum(1 for r in self.parsing_results if r.get('confidence', 0) < 0.55 and r.get('match_source') != '结构匹配')
            
            avg_confidence = sum(r.get('confidence', 0) for r in self.parsing_results) / total
            
            stats_text = f"共 {total} 条 | 高置信度: {high_confidence} | 中置信度: {medium_confidence} | 低置信度: {low_confidence} | 平均置信度: {avg_confidence:.1%}"
            self.stats_label.config(text=stats_text)
            
        except Exception as e:
            print(f"更新统计信息错误: {str(e)}")
    
    def clear_input_text(self):
        """清空输入文本"""
        self.input_text.delete("1.0", tk.END)
        self.update_line_numbers()
        self.clear_text_highlight()
        self.clear_edit_fields()
    
    def copy_parsing_results(self):
        """复制解析结果到剪贴板"""
        try:
            if not self.parsing_results:
                messagebox.showwarning("警告", "没有解析结果可复制！")
                return
            
            # 构建表格数据
            headers = ["电压等级", "报价型号", "规格", "结构描述", "备注"]
            lines = ["\t".join(headers)]
            
            for result in self.parsing_results:
                line = "\t".join([
                    result.get('voltage', ''),
                    result.get('model', ''),
                    result.get('specification', ''),
                    result.get('structure', ''),
                    result.get('remarks', '')
                ])
                lines.append(line)
            
            # 复制到剪贴板
            result_text = "\n".join(lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(result_text)
            
            messagebox.showinfo("成功", f"已复制 {len(self.parsing_results)} 条解析结果到剪贴板！\n可直接粘贴到Excel中。")
            
        except Exception as e:
            messagebox.showerror("错误", f"复制失败：{str(e)}")
    
    def parse_single_text(self, text):
        """解析单条文本 - 使用增强的解析逻辑"""
        import re
        
        result = {
            'voltage': '',
            'model': '',
            'specification': '',
            'structure': '',
            'remarks': '',
            'confidence': 0.0
        }
        
        if not text or not text.strip():
            return result
        
        text_upper = text.upper().strip()
        
        try:
            # 使用增强的解析逻辑（传递原始文本，不是大写文本）
            enhanced_result = self.enhanced_parse_text(text)
            
            # 如果增强解析识别出了完整结构，优先使用增强结果
            if (enhanced_result.get('structure') and 
                len(enhanced_result['structure'].split('/')) >= 3):
                return enhanced_result
            
            # 如果增强解析成功（置信度较高），使用其结果
            if enhanced_result['confidence'] > 0.7:
                return enhanced_result
            
            # 否则使用原有的解析逻辑作为备选
            fallback_result = self.fallback_parse_text(text)
            
            # 比较两个结果，选择更好的
            if (enhanced_result.get('structure') and 
                len(enhanced_result['structure'].split('/')) >= len(fallback_result.get('structure', '').split('/'))):
                return enhanced_result
            else:
                return fallback_result
            
        except Exception as e:
            print(f"解析错误: {str(e)}")
            return self.fallback_parse_text(text)
    
    def enhanced_parse_text(self, text):
        """增强的文本解析逻辑（基于实际业务数据）"""
        result = {
            'voltage': '',
            'model': '',
            'specification': '',
            'structure': '',
            'remarks': '',
            'confidence': 0.0,
            'match_source': '直接识别'  # 新增字段，标记匹配来源
        }
        
        # 1. 电压等级识别（基于实际数据模式）
        voltage = self.extract_voltage_enhanced(text)
        result['voltage'] = voltage
        
        # 2. 规格识别（基于实际数据模式）
        specification = self.extract_specification_enhanced(text)
        result['specification'] = specification
        
        # 3. 型号识别（基于实际数据模式）
        model = self.extract_model_enhanced(text, voltage)
        result['model'] = model
        
        # 4. 结构识别（基于实际数据模式）
        structure = self.extract_structure_enhanced(text, model)
        result['structure'] = structure
        
        # 5. 备注识别
        remarks = self.extract_remarks_enhanced(text)
        result['remarks'] = remarks
        
        # 6. 智能匹配逻辑 - 优化型号识别
        if not model or model in ['YJV.LV', 'YJLV.LV']:  # 只有默认型号才需要结构匹配
            if structure and structure != 'CU/PVC' and len(structure.split('/')) >= 3:  # 有完整结构信息
                # 通过结构搜索找到最匹配的型号，传入电压信息进行精确匹配
                structure_match = self.find_model_by_structure(structure, voltage)
                if structure_match:
                    result['model'] = structure_match['model']
                    result['match_source'] = structure_match['source']
                    # 如果是电压+结构匹配，置信度稍高一些
                    if structure_match['source'] == '结构+电压匹配':
                        result['confidence'] = min(structure_match['confidence'] * 0.85, 0.9)
                    else:
                        result['confidence'] = min(structure_match['confidence'] * 0.75, 0.8)
                else:
                    # 结构识别出来但没找到匹配型号，进一步降低置信度
                    result['confidence'] = 0.6
                    result['match_source'] = '结构推测'
            elif structure and len(structure.split('/')) >= 2:  # 有部分结构信息
                # 尝试结构匹配，但置信度更低
                structure_match = self.find_model_by_structure(structure, voltage)
                if structure_match:
                    result['model'] = structure_match['model']
                    result['match_source'] = structure_match['source']
                    if structure_match['source'] == '结构+电压匹配':
                        result['confidence'] = min(structure_match['confidence'] * 0.7, 0.8)
                    else:
                        result['confidence'] = min(structure_match['confidence'] * 0.6, 0.7)
                else:
                    result['confidence'] = 0.5
                    result['match_source'] = '部分结构推测'
            else:
                # 型号和结构都识别不出来，置信度降到红色区间 (0.3-0.5)
                result['confidence'] = 0.4
                result['match_source'] = '低置信度推测'
        else:
            # 7. 计算正常置信度（直接识别成功）
            result['confidence'] = self.calculate_confidence_enhanced(result)
            result['match_source'] = '直接识别'
        
        return result
    
    def extract_voltage_enhanced(self, text):
        """增强的电压提取 - 精准映射系统"""
        # 特殊型号的电压识别
        if re.search(r'\bH1Z2Z2[-.]?K\b', text):
            return 'DC1500V'
        
        # 光伏电缆电压识别
        solar_keywords = [
            'XLPO', 'Solar Cable', 'PV solar', 'Câble solaire', 
            'Solar wire', 'Photovoltaic', 'PV cable'
        ]
        if any(keyword.upper() in text.upper() for keyword in solar_keywords):
            return 'DC1500V'
        
        # 裸铜线无电压等级
        bare_copper_keywords = [
            'bare copper conductor', 'bare copper wire', 'bare copper',
            '铜绞线', '铜缆'
        ]
        if any(keyword.upper() in text.upper() for keyword in bare_copper_keywords):
            return 'N/A'
        
        # 建立精准的电压映射字典
        voltage_mappings = {
            # 标准格式映射
            '0.6/1.0KV': '0.6/1kV',
            '0.6/1KV': '0.6/1kV',
            '0.6/1': '0.6/1kV',
            '1KV': '0.6/1kV',
            '1.0KV': '0.6/1kV',
            
            '1.8/3.0KV': '1.8/3kV',
            '1.8/3KV': '1.8/3kV',
            '1.8/3': '1.8/3kV',
            '3KV': '1.8/3kV',
            '3.0KV': '1.8/3kV',
            
            '1.9/3.3KV': '1.9/3.3kV',
            '3.3KV': '1.9/3.3kV',
            
            '6/10KV': '6/10kV',
            '6.6KV': '6/10kV',
            '10KV': '6/10kV',
            
            '8.7/15KV': '8.7/15kV',
            '15KV': '8.7/15kV',
            
            '12/20KV': '12/20kV',
            '12.7/22KV': '12.7/22kV',
            '22KV': '12.7/22kV',
            '20KV': '12/20kV',
            
            '18/30KV': '18/30kV',
            '18/30': '18/30kV',
            '30KV': '18/30kV',
            
            '19/33KV': '19/33kV',
            '19/33': '19/33kV',
            '33KV': '19/33kV',
            
            '21/35KV': '21/35kV',
            '26/35KV': '26/35kV',
            '26/35': '26/35kV',
            '35KV': '26/35kV',
            
            # 低压格式
            '450/750V': '450/750V',
            '0.45/0.75KV': '450/750V',
            '750V': '450/750V',
            
            # 直流格式
            'DC1500V': 'DC1500V',
            'DC 1500V': 'DC1500V',
            
            # 中文千伏格式
            '10千伏': '6/10kV',
            '35千伏': '26/35kV',
        }
        
        # 电压提取模式（按优先级排序）
        voltage_patterns = [
            # NH电缆特殊格式：NHBV-450/750V
            r'NH[A-Z]*[-]?(\d+/\d+V)',
            # 中文千伏格式
            r'(\d+(?:\.\d+)?千伏)',
            # 完整格式带V：数字/数字V（优先匹配）
            r'(\d+(?:\.\d+)?/\d+(?:\.\d+)?V)',
            # 完整格式：数字/数字KV 或 数字/数字 kV
            r'(\d+(?:\.\d+)?/\d+(?:\.\d+)?\s*[Kk][Vv])',
            # 直流格式：DC数字V
            r'(DC\s*\d+V)',
            # 不完整格式：数字/数字（后面跟特定字符或结尾）
            r'(\d+(?:\.\d+)?/\d+(?:\.\d+)?)(?=[-\s]|$|[^\d\.])',
            # 单一电压：数字KV
            r'(\d+(?:\.\d+)?\s*[Kk][Vv])',
            # 单一电压：数字V
            r'(\d+V)(?!\d)',
        ]
        
        # 预处理文本：移除干扰字符
        clean_text = text.upper()
        # 移除括号内容，如 (36 kV), (40)kV
        clean_text = re.sub(r'\([^)]*\)', '', clean_text)
        # 统一空格
        clean_text = re.sub(r'\s+', ' ', clean_text)
        
        # 按优先级匹配电压
        for pattern in voltage_patterns:
            matches = re.findall(pattern, clean_text)
            for match in matches:
                # 标准化匹配结果
                normalized_match = match.strip().upper()
                normalized_match = re.sub(r'\s+', '', normalized_match)  # 移除空格
                
                # 直接映射查找
                if normalized_match in voltage_mappings:
                    return voltage_mappings[normalized_match]
                
                # 模糊匹配处理
                if 'DC' in normalized_match:
                    return normalized_match.replace(' ', '')
                elif '/' in normalized_match:
                    # 处理格式化
                    if 'KV' in normalized_match:
                        return normalized_match.lower().replace('kv', 'kV')
                    elif normalized_match.endswith('V'):
                        # 已经有V后缀，直接返回
                        return normalized_match
                    else:
                        # 补充kV后缀
                        return normalized_match + 'kV'
                elif normalized_match.endswith('V') and not normalized_match.endswith('KV'):
                    # 处理单独的V格式
                    if '750V' in normalized_match:
                        return '450/750V'
        
        # 基于型号的默认电压推断
        if any(keyword in text.upper() for keyword in ['BV', 'BVR', 'RVV', 'THW', 'H07V-R', '接地线', '接地电缆', 'GREEN/YELLOW']):
            return '450/750V'
        elif any(keyword in text.upper() for keyword in ['YJV', 'YJLV']) and not re.search(r'\d+/\d+', text):
            return '0.6/1kV'
        
        return '0.6/1kV'  # 默认低压
    
    def extract_specification_enhanced(self, text):
        """增强的规格提取 - 修复版"""
        import re
        
        # 预处理：处理特殊情况
        text_processed = text.upper()
        
        # 1. 处理3.5C x 300 Sq.mm这种3C+1E结构
        if re.search(r'3\.5C\s*[X×]\s*(\d+)', text_processed, re.IGNORECASE):
            match = re.search(r'3\.5C\s*[X×]\s*(\d+)', text_processed, re.IGNORECASE)
            if match:
                size = int(match.group(1))
                # 3.5C = 3C+1E，1E通常是主芯的一半
                earth_size = size // 2
                return f"3x{size}+1x{earth_size}"
        
        # 2. 处理Single core情况
        if 'SINGLE CORE' in text_processed:
            # 查找截面积
            size_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:MM\s*SQ|SQMM|MM²|SQ\.?MM)', text_processed, re.IGNORECASE)
            if size_match:
                size = size_match.group(1)
                # 去掉小数点后的0
                if '.' in size and size.endswith('.0'):
                    size = size[:-2]
                return f"1x{size}"
        
        # 3. 处理3C情况（3 core）- 修复120SQMM识别问题
        # 先处理3C-95 sqmm格式
        if re.search(r'(\d+)C\s*[-]\s*(\d+(?:\.\d+)?)\s*(?:SQMM|MM²?|SQ\.?MM)', text_processed, re.IGNORECASE):
            match = re.search(r'(\d+)C\s*[-]\s*(\d+(?:\.\d+)?)\s*(?:SQMM|MM²?|SQ\.?MM)', text_processed, re.IGNORECASE)
            if match:
                cores = match.group(1)
                size = match.group(2)
                # 去掉小数点后的0
                if '.' in size and size.endswith('.0'):
                    size = size[:-2]
                return f"{cores}x{size}"
        
        # 处理3C, 120SQMM或3C 120SQMM格式
        if re.search(r'(\d+)C\s*[,]?\s*(\d+(?:\.\d+)?)\s*(?:SQMM|MM²?|SQ\.?MM)', text_processed, re.IGNORECASE):
            match = re.search(r'(\d+)C\s*[,]?\s*(\d+(?:\.\d+)?)\s*(?:SQMM|MM²?|SQ\.?MM)', text_processed, re.IGNORECASE)
            if match:
                cores = match.group(1)
                size = match.group(2)
                # 去掉小数点后的0
                if '.' in size and size.endswith('.0'):
                    size = size[:-2]
                return f"{cores}x{size}"
        
        # 4. 处理2C-4.0 mm²格式
        if re.search(r'(\d+)C\s*[-]\s*(\d+(?:\.\d+)?)\s*(?:MM²?)', text_processed, re.IGNORECASE):
            match = re.search(r'(\d+)C\s*[-]\s*(\d+(?:\.\d+)?)\s*(?:MM²?)', text_processed, re.IGNORECASE)
            if match:
                cores = match.group(1)
                size = match.group(2)
                # 去掉小数点后的0
                if '.' in size and size.endswith('.0'):
                    size = size[:-2]
                return f"{cores}x{size}"
        
        # 5. 处理120SQMM这种格式（修复2.5被识别为2120SQMM的问题）
        if re.search(r'(\d+(?:\.\d+)?)\s*SQMM', text_processed, re.IGNORECASE):
            match = re.search(r'(\d+(?:\.\d+)?)\s*SQMM', text_processed, re.IGNORECASE)
            if match:
                size = match.group(1)
                # 去掉小数点后的0
                if '.' in size and size.endswith('.0'):
                    size = size[:-2]
                
                # 检查是否有芯数信息
                core_match = re.search(r'(\d+)C', text_processed, re.IGNORECASE)
                if core_match:
                    cores = core_match.group(1)
                    return f"{cores}x{size}"
                else:
                    # 检查电缆类型，决定是否添加1x前缀
                    cable_type = self._get_cable_type_from_text(text)
                    if cable_type in ['BV', 'BVR', 'PABC', 'H1Z2Z2-K']:
                        return size  # 这些类型不需要1x前缀
                    else:
                        return f"1x{size}"
        
        # 6. 处理H1Z2Z2-K-DC1500V-1×4格式（光伏线特殊处理）
        if re.search(r'H1Z2Z2-K.*?[-]\s*1\s*[×X]\s*(\d+(?:\.\d+)?)', text_processed, re.IGNORECASE):
            match = re.search(r'H1Z2Z2-K.*?[-]\s*1\s*[×X]\s*(\d+(?:\.\d+)?)', text_processed, re.IGNORECASE)
            if match:
                size = match.group(1)
                # 去掉小数点后的0
                if '.' in size and size.endswith('.0'):
                    size = size[:-2]
                return size  # 光伏线不需要1x前缀
        
        # 7. 处理3×95 mm2+1×50 mm2格式
        if re.search(r'(\d+)\s*[×X]\s*(\d+(?:\.\d+)?)\s*(?:mm2?|MM²?)\s*\+\s*(\d+)\s*[×X]\s*(\d+(?:\.\d+)?)\s*(?:mm2?|MM²?)', text_processed, re.IGNORECASE):
            match = re.search(r'(\d+)\s*[×X]\s*(\d+(?:\.\d+)?)\s*(?:mm2?|MM²?)\s*\+\s*(\d+)\s*[×X]\s*(\d+(?:\.\d+)?)\s*(?:mm2?|MM²?)', text_processed, re.IGNORECASE)
            if match:
                cores1, size1, cores2, size2 = match.groups()
                # 去掉小数点后的0
                if '.' in size1 and size1.endswith('.0'):
                    size1 = size1[:-2]
                if '.' in size2 and size2.endswith('.0'):
                    size2 = size2[:-2]
                return f"{cores1}x{size1}+{cores2}x{size2}"
        
        # 8. 规格模式（按优先级，基于实际数据）
        spec_patterns = [
            # 复杂格式优先：3x300+1x150, 3×240+1×120等
            r'(\d+)\s*[×X]\s*(\d+(?:\.\d+)?)\s*\+\s*(\d+)\s*[×X]\s*(\d+(?:\.\d+)?)',
            # NHBV特殊格式：NHBV-450/750V1x2.5
            r'NHBV[-]?\d+/\d+V(\d+x\d+(?:\.\d+)?)',
            # BVR4mm2 格式分离
            r'\bBVR(\d+(?:\.\d+)?mm2?)\b',
            # AWG格式
            r'(\d+)\s*AWG',
            # 英文格式：300 mm² 3C 或 3C 300 mm²
            r'(\d+(?:\.\d+)?)\s*mm²?\s*(\d+)C',  # 300 mm² 3C
            r'(\d+)C\s*(\d+(?:\.\d+)?)\s*mm²?',  # 3C 300 mm²
            r'(\d+)C.*?(\d+(?:\.\d+)?)\s*mm²?',  # 3C ... 300 mm²
            # 标准格式：3x2.5, 3×240等
            r'(\d+)\s*[CX×*]\s*(\d+(?:\.\d+)?)\s*(?:MM²?|SQMM)?',
            # 单独的mm²格式
            r'(\d+(?:\.\d+)?)\s*(?:mm²?|MM²?)',
        ]
        
        for pattern in spec_patterns:
            matches = re.findall(pattern, text_processed, re.IGNORECASE)
            if matches:
                match = matches[0]
                
                # 特殊处理BVR4mm2格式
                if 'BVR' in text_processed and isinstance(match, str) and 'MM' in match.upper():
                    # 提取数字部分
                    spec_num = re.search(r'(\d+(?:\.\d+)?)', match)
                    if spec_num:
                        size = spec_num.group(1)
                        # 去掉小数点后的0
                        if '.' in size and size.endswith('.0'):
                            size = size[:-2]
                        return size  # BVR不需要1x前缀
                
                # 特殊处理AWG格式
                if 'AWG' in text_processed and isinstance(match, str) and match.isdigit():
                    return match + 'AWG'
                
                if isinstance(match, tuple):
                    # 处理复杂格式：3x300+1x150
                    if len(match) == 4:  # 复杂格式 (cores1, size1, cores2, size2)
                        cores1, size1, cores2, size2 = match
                        # 去掉小数点后的0
                        if '.' in size1 and size1.endswith('.0'):
                            size1 = size1[:-2]
                        if '.' in size2 and size2.endswith('.0'):
                            size2 = size2[:-2]
                        return f"{cores1}x{size1}+{cores2}x{size2}"
                    
                    # 处理英文格式：300 mm² 3C 或 3C 300 mm²
                    elif len(match) == 2 and match[0] and match[1]:
                        # 判断哪个是芯数，哪个是截面
                        try:
                            val1, val2 = float(match[0]), float(match[1])
                            if val1 <= 10 and val2 > 10:  # val1是芯数，val2是截面
                                size = str(int(val2)) if val2.is_integer() else str(val2)
                                return f"{int(val1)}x{size}"
                            elif val2 <= 10 and val1 > 10:  # val2是芯数，val1是截面
                                size = str(int(val1)) if val1.is_integer() else str(val1)
                                return f"{int(val2)}x{size}"
                            else:
                                # 默认第一个是芯数
                                size1 = str(int(val1)) if val1.is_integer() else str(val1)
                                size2 = str(int(val2)) if val2.is_integer() else str(val2)
                                return f"{size1}x{size2}"
                        except ValueError:
                            # 如果转换失败，使用原始值
                            return f"{match[0]}x{match[1]}"
                    
                    # 处理标准格式：芯数x截面积
                    if len(match) == 2 and match[0] and match[1]:
                        cores = match[0]
                        size = match[1]
                        # 去掉小数点后的0
                        if '.' in size and size.endswith('.0'):
                            size = size[:-2]
                        return f"{cores}x{size}"
                else:
                    spec = match
                    # 标准化格式，移除换行符和多余空格
                    spec = spec.replace('\n', '').replace('\r', '').strip()
                    spec = spec.replace('C', 'x').replace('X', 'x').replace('×', 'x').replace('*', 'x')
                    spec = spec.replace('MM²', '').replace('MM', '').replace('SQMM', '').replace('mm²', '').replace('mm', '')
                    
                    # 去掉小数点后的0
                    if '.' in spec and spec.endswith('.0'):
                        spec = spec[:-2]
                    
                    # 检查电缆类型，决定是否添加1x前缀
                    if spec.isdigit() or re.match(r'^\d+\.\d+$', spec):
                        cable_type = self._get_cable_type_from_text(text)
                        if cable_type in ['BV', 'BVR', 'PABC', 'H1Z2Z2-K']:
                            return spec  # 这些类型不需要1x前缀
                        else:
                            return f"1x{spec}"
                    
                    return spec
        
        return ''
    
    def _get_cable_type_from_text(self, text: str) -> str:
        """从文本中获取电缆类型"""
        text_upper = text.upper()
        
        # 光伏线
        if any(keyword in text_upper for keyword in ['H1Z2Z2', 'PV1-F', 'SOLAR']):
            return 'H1Z2Z2-K'
        
        # 布电线
        if 'BVR' in text_upper:
            return 'BVR'
        elif 'BV' in text_upper:
            return 'BV'
        
        # 裸导线
        if 'PABC' in text_upper:
            return 'PABC'
        
        # 控制电缆
        if any(keyword in text_upper for keyword in ['RVV', 'KVV']):
            return 'RVV'
        
        # 默认为电力电缆
        return 'YJV'
    
    def extract_model_enhanced(self, text, voltage):
        """增强的型号提取（基于分层判断逻辑）"""
        try:
            # 导入高级型号识别器
            from advanced_model_recognition import AdvancedModelRecognition
            
            # 创建识别器实例
            recognizer = AdvancedModelRecognition()
            
            # 使用高级识别逻辑
            result = recognizer.recognize_model(text)
            
            # 返回识别的型号
            return result.model_name
            
        except ImportError:
            # 如果无法导入高级识别器，使用原有逻辑
            return self.extract_model_enhanced_fallback(text, voltage)
        except Exception as e:
            print(f"高级型号识别失败: {str(e)}")
            return self.extract_model_enhanced_fallback(text, voltage)
    
    def extract_model_enhanced_fallback(self, text, voltage):
        """原有的型号提取逻辑（作为备选）"""
        # 特殊型号直接识别
        if re.search(r'\bH1Z2Z2[-.]?K\b', text):
            return 'H1Z2Z2.K'
        elif re.search(r'\bPABC\b', text):
            return 'PABC'
        elif re.search(r'\bRVV\b', text):
            return 'RVV'
        
        # 接地线识别（特殊处理）- 优先级最高
        if ('Green/Yellow' in text and 'Insulated Copper Cable' in text) or \
           any(keyword in text for keyword in ['黄绿接地线', '接地电缆', '接地线', 'H07V-R']):
            return 'BV.LV'
        
        # 光伏电缆识别（多种描述）
        solar_keywords = [
            'XLPO', 'Solar Cable', 'PV solar', 'Câble solaire', 
            'Solar wire', 'Photovoltaic', 'PV cable'
        ]
        if any(keyword.upper() in text.upper() for keyword in solar_keywords):
            return 'H1Z2Z2.K'
        
        # 裸铜线识别（多种描述）- 优先级高于接地线
        bare_copper_keywords = [
            'bare copper conductor', 'bare copper wire', 'bare copper',
            '铜绞线', '铜缆'
        ]
        if any(keyword.upper() in text.upper() for keyword in bare_copper_keywords):
            return 'PABC'
        
        # 检查前缀（更精确的匹配）
        prefix = ''
        special_performance = ''
        
        # 检查特殊性能指标
        if re.search(r'\bFSY', text) or re.search(r'\bFYS', text):  # FSY或FYS都可能表示防鼠蚁
            special_performance = '防鼠蚁'
        
        # 检查阻燃前缀
        if re.search(r'\bZRC[-.]?', text) or re.search(r'\bZC[-.]?', text):
            prefix = 'ZC'
        elif re.search(r'\bZR[-.]?', text) and not re.search(r'\bZRC[-.]?', text):
            prefix = 'ZC'  # ZR也映射为ZC
        elif re.search(r'\bWDZN[-.]?', text):
            prefix = 'WDZN'
        elif re.search(r'\bWDZC[-.]?', text):
            prefix = 'WDZC'
        elif re.search(r'\bWDZ[-.]?', text) or '低烟无卤' in text or '无卤低烟' in text:
            prefix = 'WDZ'
        elif re.search(r'\bNH[-.]?', text) or re.search(r'^NH[A-Z]', text):
            prefix = 'NH'
        
        # 检查基础型号（更精确的匹配，包括铠装）
        base_model = ''
        armor_suffix = ''
        
        # 优先检查特殊线缆类型和规格分离
        if re.search(r'\bBVR(\d+(?:\.\d+)?mm2?)\b', text, re.IGNORECASE):
            # BVR4mm2 格式，需要分离
            base_model = 'BVR'
        elif re.search(r'\bNHBV\b', text):
            # NHBV格式
            base_model = 'BV'
        elif re.search(r'\b(THW|BV|BVR)\b', text):
            if 'THW' in text or 'BV' in text:
                base_model = 'BV'
            elif 'BVR' in text:
                base_model = 'BVR'
        elif re.search(r'\bBYJ\b', text):
            base_model = 'BYJ'
        elif re.search(r'\bYJLHY23\b', text):
            base_model = 'YJLHY'
            armor_suffix = '23'
        elif re.search(r'\bYJLV22\b', text):  # 先匹配带铠装的
            base_model = 'YJLV'
            armor_suffix = '22'
        elif re.search(r'\bYJLV32\b', text):
            base_model = 'YJLV'
            armor_suffix = '32'
        elif re.search(r'\bYJLV62\b', text):
            base_model = 'YJLV'
            armor_suffix = '62'
        elif re.search(r'\bYJLV72\b', text):
            base_model = 'YJLV'
            armor_suffix = '72'
        elif re.search(r'\bYJLV23\b', text):
            base_model = 'YJLV'
            armor_suffix = '23'
        elif re.search(r'\bYJLV\b', text):
            base_model = 'YJLV'
        elif re.search(r'\bYJY23\b', text):  # 先匹配带铠装的YJY
            base_model = 'YJY'
            armor_suffix = '23'
        elif re.search(r'\bYJV22\b', text):  # 先匹配带铠装的
            base_model = 'YJV'
            armor_suffix = '22'
        elif re.search(r'\bYJV32\b', text):
            base_model = 'YJV'
            armor_suffix = '32'
        elif re.search(r'\bYJV62\b', text):
            base_model = 'YJV'
            armor_suffix = '62'
        elif re.search(r'\bYJV72\b', text):
            base_model = 'YJV'
            armor_suffix = '72'
        elif re.search(r'\bYJV23\b', text):
            base_model = 'YJV'
            armor_suffix = '23'
        elif re.search(r'\bYJVR?\b', text):
            base_model = 'YJV'
        elif re.search(r'\bYJY\b', text):
            base_model = 'YJY'
        
        # 如果没有从型号中提取到铠装，再检查独立的铠装标识
        if not armor_suffix:
            if re.search(r'\b22\b', text) or 'ARMOURED' in text:
                armor_suffix = '22'
            elif re.search(r'\b32\b', text):
                armor_suffix = '32'
            elif re.search(r'\b62\b', text):
                armor_suffix = '62'
            elif re.search(r'\b72\b', text):
                armor_suffix = '72'
            elif re.search(r'\b23\b', text):
                armor_suffix = '23'
        
        # 电压后缀
        voltage_suffix = ''
        if voltage in ['0.6/1kV', '450/750V']:
            voltage_suffix = 'LV'
        elif voltage in ['1.8/3kV', '1.9/3.3kV']:
            voltage_suffix = 'LV.3kV'  # 特殊的3kV低压
        elif voltage in ['6/10kV', '8.7/15kV', '12/20kV', '12.7/22kV', '18/30kV', '19/33kV', '21/35kV', '26/35kV']:
            voltage_suffix = 'MV'
        
        # 构建型号
        model_parts = []
        if prefix:
            model_parts.append(prefix)
        if base_model:
            if armor_suffix:
                # 将铠装后缀直接附加到基础型号上
                model_parts.append(base_model + armor_suffix)
            else:
                model_parts.append(base_model)
        
        # 只有标准电缆类型才添加电压后缀，特殊电缆类型不添加
        special_cable_types = ['BV', 'BVR', 'H1Z2Z2.K', 'PABC', 'HDBC', 'RVV']
        should_add_voltage_suffix = True
        
        if base_model in special_cable_types:
            should_add_voltage_suffix = False
        elif any(special_type in '.'.join(model_parts) for special_type in special_cable_types):
            should_add_voltage_suffix = False
        
        if voltage_suffix and should_add_voltage_suffix:
            model_parts.append(voltage_suffix)
        
        # 如果没有基础型号，根据文本推断
        if not base_model:
            # 英文格式的型号推断
            if 'AL' in text or 'ALUMINIUM' in text:
                inferred_base = 'YJLV'
            else:
                inferred_base = 'YJV'
            
            # 检查铠装类型（英文格式）
            inferred_armor = ''
            if 'SWA' in text or 'STEEL WIRE ARMOURED' in text:
                inferred_armor = '32'
            elif 'STA' in text or 'STEEL TAPE ARMOURED' in text or 'ARMOURED' in text:
                inferred_armor = '22'
            
            # 构建推断型号
            if inferred_armor:
                model_parts = [inferred_base + inferred_armor]
            else:
                model_parts = [inferred_base]
            
            # 推断的型号通常是标准电缆类型，需要添加电压后缀
            if voltage_suffix:
                model_parts.append(voltage_suffix)
        
        if model_parts:
            return '.'.join(model_parts)
        
        # 默认返回
        return 'YJV.LV'
    
    def extract_structure_enhanced(self, text, model):
        """增强的结构提取（基于实际数据模式）"""
        structure_parts = []
        
        # 特殊型号的结构处理
        if 'H1Z2Z2.K' in model:
            return 'TAC/XLPO/XLPO'
        elif 'PABC' in model:
            return 'PABC'  # 裸铜线只返回型号
        
        # 检查是否直接包含结构描述（如 "AL/XLPE/CTS/PVC"）
        structure_pattern = r'([A-Z]+(?:/[A-Z]+){2,})'
        structure_match = re.search(structure_pattern, text.upper())
        if structure_match:
            potential_structure = structure_match.group(1)
            # 验证是否是有效的电缆结构
            structure_components = potential_structure.split('/')
            valid_materials = ['CU', 'AL', 'TAC', 'XLPE', 'PVC', 'XLPO', 'LSZH', 'HDPE', 'CTS', 'CWS', 'STA', 'SWA', 'AWA', 'SSTA', 'MT']
            if len(structure_components) >= 3 and all(comp in valid_materials for comp in structure_components):
                return potential_structure
        
        # 检查是否有明确的电缆特征，如果没有则返回空结构
        cable_indicators = [
            # 中文电缆特征
            '电缆', '导体', '绝缘', '护套', '铠装', '芯', '铜芯', '铝芯', '交联', '聚乙烯', '聚氯乙烯',
            # 英文电缆特征
            'CABLE', 'CONDUCTOR', 'INSULATED', 'SHEATHED', 'ARMOURED', 'COPPER', 'ALUMINUM', 'ALUMINIUM',
            'XLPE', 'PVC', 'CORE', 'WIRE',
            # 型号特征
            'YJV', 'YJLV', 'BV', 'BVR', 'RVV', 'KVV', 'NH', 'ZR', 'ZC', 'WDZ'
        ]
        
        # 如果文本中没有任何电缆特征，返回空结构
        if not any(indicator in text.upper() for indicator in cable_indicators) and not model:
            return ''
        
        # 如果只有型号但没有其他特征，且型号是默认的，也返回空结构
        if model in ['YJV.LV', 'YJLV.LV'] and not any(indicator in text.upper() for indicator in cable_indicators[:15]):  # 只检查明确的电缆特征，不包括型号
            return ''
        
        # 导体（基于实际数据模式）
        if 'AL' in text or 'ALUMINIUM' in text or 'YJLV' in model:
            structure_parts.append('AL')
        elif 'TAC' in model or 'H1Z2Z2' in model:
            structure_parts.append('TAC')
        elif 'Cl5' in text or 'YJVR' in model or 'RVV' in model:
            structure_parts.append('Cl5 CU')
        else:
            structure_parts.append('CU')
        
        # 绝缘（基于实际数据模式）
        if 'NH' in model or 'NH' in text:
            if 'XLPE' in text or 'YJ' in model:
                structure_parts.append('MT/XLPE')
            else:
                structure_parts.append('MT/PVC')
        elif 'XLPE' in text or 'YJ' in model:
            structure_parts.append('XLPE')
        elif 'XLPO' in text or 'H1Z2Z2' in model or 'BYJ' in model:
            structure_parts.append('XLPO')
        elif 'BV' in model or 'THW' in text:
            # BV线缆使用PVC绝缘
            structure_parts.append('PVC')
        else:
            structure_parts.append('PVC')
        
        # 屏蔽（中压电缆，基于实际数据）
        if any(v in text for v in ['1.8/3', '6/10KV', '8.7/15KV', '26/35KV', '19/33']) or 'MV' in model:
            if 'CWS' in text:
                structure_parts.append('CWS')
            else:
                structure_parts.append('CTS')
        
        # 内护套（基于实际数据模式和英文格式）
        has_armor = any(armor in model for armor in ['22', '32', '62', '72', '23']) or any(armor in text for armor in ['SWA', 'STA', 'ARMOURED'])
        
        is_lszh = 'LSZH' in text or 'WDZ' in model or '低烟无卤' in text or '无卤低烟' in text
        if is_lszh:
            if has_armor:
                structure_parts.append('LSZH')
        elif has_armor:
            # YJY23使用HDPE内护套
            if 'YJY23' in model:
                structure_parts.append('HDPE')
            else:
                structure_parts.append('PVC')
        
        # 铠装（基于实际数据模式）
        if '22' in model or '23' in model or ('STA' in text and 'SWA' not in text):
            structure_parts.append('STA')
        elif '32' in model or 'SWA' in text:
            structure_parts.append('SWA')
        elif '62' in model:
            structure_parts.append('SSTA')
        elif '72' in model:
            structure_parts.append('AWA')
        
        # 外护套（基于实际数据模式）
        # BV线缆通常没有护套，只有绝缘
        if 'BV' in model and not any(armor in model for armor in ['22', '32', '62', '72', '23']):
            # BV线缆不添加护套
            pass
        elif is_lszh:
            structure_parts.append('LSZH')
        elif 'HDPE' in text or '23' in model:
            structure_parts.append('HDPE')
        elif 'H1Z2Z2' in model or 'BYJ' in model:
            structure_parts.append('XLPO')
        else:
            structure_parts.append('PVC')
        
        return '/'.join(structure_parts)
    
    def extract_remarks_enhanced(self, text):
        """增强的备注提取"""
        remarks = []
        
        # 特殊性能指标
        if re.search(r'\bFSY', text) or re.search(r'\bFYS', text):  # FSY或FYS都可能表示防鼠蚁
            remarks.append('防鼠蚁')
        
        # 阻燃等级
        if re.search(r'\bZA\b', text):
            remarks.append('ZA')
        elif re.search(r'\bZB\b', text):
            remarks.append('ZB')
        elif re.search(r'\bZC\b', text):
            remarks.append('ZC')
        elif re.search(r'\bZR\b', text):
            remarks.append('ZC')  # ZR也映射为ZC
        
        # 耐火
        if re.search(r'\bNH\b', text):
            remarks.append('NH')
        
        # 低烟无卤
        if 'LSZH' in text or 'WDZ' in text:
            remarks.append('低烟无卤')
        
        # 特殊要求
        if '防鼠' in text:
            remarks.append('防鼠')
        if '防白蚁' in text:
            remarks.append('防白蚁')
        if '耐油' in text:
            remarks.append('耐油')
        if '阻燃' in text:
            remarks.append('阻燃')
        
        return ' '.join(remarks)
    
    def find_model_by_structure(self, structure, voltage=None):
        """通过结构搜索找到最匹配的型号，考虑电压等级匹配"""
        try:
            # 使用现有的结构搜索功能
            search_results = self.code_manager.search_by_structure(structure)
            
            if search_results:
                # 如果提供了电压信息，优先选择电压匹配的结果
                if voltage:
                    voltage_matched_results = []
                    for result in search_results:
                        spec_data = result.get('spec_data')
                        if spec_data and len(spec_data) > 1:
                            result_voltage = spec_data[1]  # 电压等级在索引1
                            if result_voltage == voltage:
                                voltage_matched_results.append(result)
                    
                    # 如果有电压匹配的结果，从中选择置信度最高的
                    if voltage_matched_results:
                        best_result = max(voltage_matched_results, key=lambda x: x.get('confidence', 0))
                        
                        # 获取最佳显示名称
                        spec_data = best_result.get('spec_data')
                        aliases = best_result.get('aliases', [])
                        spec_id = best_result.get('spec_id')
                        
                        display_name = self.get_best_display_name(spec_id, aliases, spec_data)
                        
                        # 为1.8/3kV和1.9/3.3kV电压添加3kV后缀
                        if voltage in ['1.8/3kV', '1.9/3.3kV'] and not display_name.endswith('.3kV'):
                            if '.' in display_name:
                                # 如果已有后缀，替换为3kV
                                parts = display_name.split('.')
                                display_name = f"{parts[0]}.{parts[1]}.3kV"
                            else:
                                # 如果没有后缀，添加LV.3kV
                                display_name = f"{display_name}.LV.3kV"
                        
                        return {
                            'model': display_name,
                            'confidence': best_result.get('confidence', 0.7),
                            'spec_id': spec_id,
                            'source': '结构+电压匹配'
                        }
                
                # 如果没有电压匹配的结果，或者没有提供电压，选择置信度最高的结果
                best_result = max(search_results, key=lambda x: x.get('confidence', 0))
                
                # 获取最佳显示名称
                spec_data = best_result.get('spec_data')
                aliases = best_result.get('aliases', [])
                spec_id = best_result.get('spec_id')
                
                display_name = self.get_best_display_name(spec_id, aliases, spec_data)
                
                # 为1.8/3kV和1.9/3.3kV电压添加3kV后缀（即使没有精确匹配）
                if voltage in ['1.8/3kV', '1.9/3.3kV'] and not display_name.endswith('.3kV'):
                    if '.' in display_name:
                        # 如果已有后缀，替换为3kV
                        parts = display_name.split('.')
                        if len(parts) >= 2:
                            display_name = f"{parts[0]}.{parts[1]}.3kV"
                        else:
                            display_name = f"{display_name}.LV.3kV"
                    else:
                        # 如果没有后缀，添加LV.3kV
                        display_name = f"{display_name}.LV.3kV"
                
                return {
                    'model': display_name,
                    'confidence': best_result.get('confidence', 0.7),
                    'spec_id': spec_id,
                    'source': '结构匹配'
                }
            
            return None
            
        except Exception as e:
            print(f"结构匹配搜索失败: {str(e)}")
            return None
    
    def calculate_confidence_enhanced(self, result):
        """增强的置信度计算"""
        confidence = 0.0
        
        # 电压识别置信度 (25%)
        if result['voltage']:
            voltage_text = result['voltage'].strip()
            if re.match(r'^\d+(\.\d+)?/\d+(\.\d+)?kV$', voltage_text):  # 标准格式如 0.6/1kV
                confidence += 0.25
            elif re.match(r'^\d+kV$', voltage_text):  # 简单格式如 10kV
                confidence += 0.20
            elif 'kV' in voltage_text or 'V' in voltage_text:  # 包含电压单位
                confidence += 0.15
            else:
                confidence += 0.10  # 识别到但格式不标准
        
        # 型号识别置信度 (30%)
        if result['model']:
            model_text = result['model'].strip()
            # 检查是否是标准型号格式
            if re.match(r'^[A-Z]{2,6}(\d+)?(\.[A-Z]{2,3})?$', model_text):  # 如 YJV22.LV
                confidence += 0.30
            elif re.match(r'^[A-Z]{2,4}\d*$', model_text):  # 如 YJV22
                confidence += 0.25
            elif len(model_text) >= 3 and model_text.isalpha():  # 纯字母型号
                confidence += 0.20
            else:
                confidence += 0.15  # 识别到但不确定
        
        # 规格识别置信度 (25%)
        if result['specification']:
            spec_text = result['specification'].strip()
            # 检查规格格式的准确性
            if re.match(r'^\d+x\d+(\.\d+)?(\+\d+x\d+(\.\d+)?)*$', spec_text):  # 标准格式如 3x120+1x70
                confidence += 0.25
            elif re.match(r'^\d+x\d+(\.\d+)?$', spec_text):  # 简单格式如 3x120
                confidence += 0.22
            elif re.match(r'^\d+(\.\d+)?$', spec_text):  # 只有截面积如 120
                confidence += 0.18
            else:
                confidence += 0.12  # 识别到但格式不标准
        
        # 结构识别置信度 (20%)
        if result['structure']:
            structure_text = result['structure'].strip()
            # 检查结构描述的完整性
            structure_parts = structure_text.split('/')
            if len(structure_parts) >= 3:  # 完整结构如 CU/XLPE/PVC
                confidence += 0.20
            elif len(structure_parts) == 2:  # 部分结构
                confidence += 0.15
            else:
                confidence += 0.10  # 简单结构
        
        # 综合评估奖励
        filled_fields = sum(1 for field in ['voltage', 'model', 'specification', 'structure'] 
                           if result.get(field, '').strip())
        
        if filled_fields >= 4:  # 所有字段都有
            confidence += 0.05
        elif filled_fields >= 3:  # 大部分字段有
            confidence += 0.03
        
        # 确保置信度在合理范围内
        return min(max(confidence, 0.1), 1.0)
        """增强的置信度计算"""
        confidence = 0.0
        
        # 电压识别置信度
        if result['voltage']:
            confidence += 0.25
        
        # 型号识别置信度
        if result['model']:
            confidence += 0.30
        
        # 规格识别置信度
        if result['specification']:
            confidence += 0.25
        
        # 结构识别置信度
        if result['structure']:
            confidence += 0.20
        
        return min(confidence, 1.0)
    
    def fallback_parse_text(self, text):
        """原有的解析逻辑作为备选"""
        import re
        
        result = {
            'voltage': '',
            'model': '',
            'specification': '',
            'structure': '',
            'remarks': '',
            'confidence': 0.0
        }
        
        # 电压等级提取
        voltage_patterns = [
            r'(\d+(?:\.\d+)?/\d+(?:\.\d+)?kV)',  # 8.7/15kV, 0.6/1kV
            r'(\d+(?:\.\d+)?千伏)',               # 10千伏
            r'(DC\s*\d+V)',                       # DC 1500V
        ]
        
        for pattern in voltage_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                voltage = match.group(1)
                # 标准化电压表示
                if '千伏' in voltage:
                    voltage_num = voltage.replace('千伏', '')
                    if voltage_num == '10':
                        voltage = '6/10kV'
                    elif voltage_num == '35':
                        voltage = '26/35kV'
                result['voltage'] = voltage
                break
        
        # 规格提取
        spec_patterns = [
            r'(\d+[xX×]\d+(?:\+\d+[xX×]\d+)*(?:mm²?)?)',  # 3x240+1x120
            r'(\d+mm²?)',                                  # 240mm²
        ]
        
        for pattern in spec_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                result['specification'] = match.group(1).replace('×', 'x').replace('X', 'x')
                break
        
        # 型号识别和匹配
        model_match = self.match_cable_model(text)
        if model_match:
            result['model'] = model_match['model']
            result['structure'] = model_match['structure']
            result['confidence'] = model_match['confidence']
        else:
            # 如果没有匹配到，尝试根据描述推断型号
            inferred_model = self.infer_model_from_description(text, result['voltage'])
            if inferred_model:
                result['model'] = inferred_model
                result['structure'] = self.infer_basic_structure(inferred_model, text)
                result['confidence'] = 0.6
        
        # 特殊性能识别
        special_patterns = [
            r'(ZA|ZB|ZC|ZR)(?:级)?',              # 阻燃等级
            r'(NH)(?:级)?',                       # 耐火
            r'(防鼠|防白蚁|耐油|耐酸碱)',          # 特殊要求
            r'(低烟无卤|LSZH|WDZ)',               # 低烟无卤
        ]
        
        remarks = []
        for pattern in special_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            remarks.extend(matches)
        
        result['remarks'] = ' '.join(set(remarks))
        
        return result
        import re
        
        result = {
            'original_text': text,
            'voltage': '',
            'model': '',
            'specification': '',
            'structure': '',
            'remarks': '',
            'confidence': 0.0
        }
        
        # 电压等级提取
        voltage_patterns = [
            r'(\d+(?:\.\d+)?/\d+(?:\.\d+)?kV)',  # 8.7/15kV, 0.6/1kV
            r'(\d+(?:\.\d+)?千伏)',               # 10千伏
            r'(DC\s*\d+V)',                       # DC 1500V
        ]
        
        for pattern in voltage_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                voltage = match.group(1)
                # 标准化电压表示
                if '千伏' in voltage:
                    voltage_num = voltage.replace('千伏', '')
                    if voltage_num == '10':
                        voltage = '6/10kV'
                    elif voltage_num == '35':
                        voltage = '26/35kV'
                result['voltage'] = voltage
                break
        
        # 规格提取
        spec_patterns = [
            r'(\d+[xX×]\d+(?:\+\d+[xX×]\d+)*(?:mm²?)?)',  # 3x240+1x120
            r'(\d+mm²?)',                                  # 240mm²
        ]
        
        for pattern in spec_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                result['specification'] = match.group(1).replace('×', 'x').replace('X', 'x')
                break
        
        # 型号识别和匹配
        model_match = self.match_cable_model(text)
        if model_match:
            result['model'] = model_match['model']
            result['structure'] = model_match['structure']
            result['confidence'] = model_match['confidence']
        else:
            # 如果没有匹配到，尝试根据描述推断型号
            inferred_model = self.infer_model_from_description(text, result['voltage'])
            if inferred_model:
                result['model'] = inferred_model
                result['structure'] = self.infer_basic_structure(inferred_model, text)
                result['confidence'] = 0.6
        
        # 特殊性能识别
        special_patterns = [
            r'(ZA|ZB|ZC|ZR)(?:级)?',              # 阻燃等级
            r'(NH)(?:级)?',                       # 耐火
            r'(防鼠|防白蚁|耐油|耐酸碱)',          # 特殊要求
            r'(低烟无卤|LSZH|WDZ)',               # 低烟无卤
        ]
        
        remarks = []
        for pattern in special_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            remarks.extend(matches)
        
        result['remarks'] = ' '.join(set(remarks))
        
        return result
    
    def match_cable_model(self, text):
        """匹配电缆型号"""
        # 使用现有的编码管理器进行匹配
        try:
            # 尝试通过别名搜索
            search_results = self.code_manager.search_by_alias(text)
            
            if search_results:
                best_match = search_results[0]
                spec_data = best_match.get('spec_data')
                spec_id = best_match['spec_id']
                
                # 获取正确的型号名称
                model = ""
                
                # 检查spec_data是否存在且有效
                if spec_data and len(spec_data) > 10 and spec_data[10]:  # product_model字段
                    model = spec_data[10]
                else:
                    # 尝试从备注中提取型号
                    remarks = best_match.get('remarks', '')
                    if '产品型号:' in remarks:
                        model = remarks.split('产品型号:')[-1].strip()
                    elif '匹配别名:' in remarks:
                        model = remarks.split('匹配别名:')[-1].strip()
                    else:
                        # 尝试根据规格数据生成型号名称
                        if spec_data and len(spec_data) > 3:
                            generated_model = self.generate_model_from_spec(spec_data)
                            if generated_model:
                                model = generated_model
                            else:
                                # 最后使用规格ID
                                model = spec_id
                        else:
                            model = spec_id
                
                # 获取正确的结构字符串
                structure = ""
                # 检查spec_data是否存在且有效
                if spec_data and len(spec_data) > 9 and spec_data[9]:  # structure_string字段
                    structure = spec_data[9]
                elif spec_data:
                    # 根据规格数据构建结构字符串
                    structure = self.build_structure_from_spec(spec_data)
                else:
                    # 如果没有spec_data，尝试基础推断
                    structure = self.infer_basic_structure("", text)
                
                return {
                    'model': model,
                    'structure': structure,
                    'confidence': best_match.get('confidence', 0.8)
                }
            
            # 如果没有找到匹配，尝试基础型号识别
            model_patterns = [
                r'([A-Z]{2,6}(?:\d{1,2})?(?:\.[A-Z]{2})?)',  # YJV, YJV22
                r'(NH[-.]?[A-Z]{2,6})',                      # NH-YJV
                r'(WDZ[-.]?[A-Z]{2,6})',                     # WDZ-YJV
            ]
            
            for pattern in model_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    model = match.group(1).upper()
                    
                    # 基础结构推断
                    structure = self.infer_basic_structure(model, text)
                    
                    return {
                        'model': model,
                        'structure': structure,
                        'confidence': 0.6
                    }
            
        except Exception as e:
            print(f"匹配错误: {str(e)}")
        
        return None
    
    def build_structure_from_spec(self, spec_data):
        """根据规格数据构建结构字符串（按照参数卡片格式）"""
        parts = []
        
        try:
            # spec_data格式（更新后包含product_model）: 
            # [category, voltage_rating, conductor, insulation, shield_type, 
            #  armor, outer_sheath, is_fire_resistant, special_performance, structure_string, product_model]
            # 索引:  0         1              2         3           4            
            #        5      6             7                8                  9              10
            
            # 导体
            if len(spec_data) > 2 and spec_data[2]:
                parts.append(spec_data[2])
            
            # 绝缘（考虑耐火）
            if len(spec_data) > 3 and spec_data[3] and spec_data[3] not in ['None', '无']:
                if len(spec_data) > 7 and spec_data[7]:  # is_fire_resistant
                    parts.append(f"MT/{spec_data[3]}")
                else:
                    parts.append(spec_data[3])
            
            # 屏蔽
            if len(spec_data) > 4 and spec_data[4] and spec_data[4] not in ['None', '无']:
                parts.append(spec_data[4])
            
            # 内护套 - 注意：在当前的查询中没有inner_sheath字段，跳过
            
            # 铠装
            if len(spec_data) > 5 and spec_data[5] and spec_data[5] not in ['None', '无']:
                parts.append(spec_data[5])
            
            # 外护套
            if len(spec_data) > 6 and spec_data[6] and spec_data[6] not in ['None', '无']:
                parts.append(spec_data[6])
            
        except Exception as e:
            print(f"构建结构字符串错误: {str(e)}")
        
        return "/".join(parts)
    
    def generate_model_from_spec(self, spec_data):
        """根据规格数据生成型号名称"""
        try:
            # spec_data格式: [category, voltage_rating, conductor, insulation, shield_type, 
            #                armor, outer_sheath, is_fire_resistant, special_performance, structure_string, product_model]
            
            if not spec_data or len(spec_data) < 4:
                return None
            
            parts = []
            
            # 耐火前缀
            if len(spec_data) > 7 and spec_data[7]:  # is_fire_resistant
                parts.append("NH")
            
            # 基础型号
            base_model = ""
            
            # 导体类型
            conductor = spec_data[2] if len(spec_data) > 2 else ""
            insulation = spec_data[3] if len(spec_data) > 3 else ""
            
            if insulation == "XLPE":
                if conductor == "AL":
                    base_model = "YJLV"
                else:
                    base_model = "YJV"
            elif insulation == "PVC":
                if conductor == "AL":
                    base_model = "VLV"
                else:
                    base_model = "VV"
            elif insulation == "LSZH":
                base_model = "WDZ-YJV" if conductor != "AL" else "WDZ-YJLV"
            
            if base_model:
                parts.append(base_model)
            else:
                # 如果无法确定基础型号，使用通用型号
                parts.append("YJV" if conductor != "AL" else "YJLV")
            
            # 铠装类型
            armor = spec_data[5] if len(spec_data) > 5 else ""
            if armor == "STA":
                parts.append("22")
            elif armor == "SWA":
                parts.append("32")
            elif armor == "SSTA":
                parts.append("62")
            elif armor == "AWA":
                parts.append("72")
            
            # 电压等级后缀
            voltage = spec_data[1] if len(spec_data) > 1 else ""
            if voltage in ["0.6/1kV", "1kV"]:
                parts.append("LV")
            elif voltage in ["6/10kV", "8.7/15kV", "12/20kV"]:
                parts.append("MV")
            elif voltage in ["26/35kV"]:
                parts.append("MV")
            
            # 确保至少返回基础型号
            if len(parts) >= 1:
                return ".".join(parts)
            else:
                return "YJV"  # 默认型号
            
        except Exception as e:
            print(f"生成型号名称错误: {str(e)}")
            return "YJV"  # 出错时返回默认型号而不是None
    
    def infer_basic_structure(self, model, text):
        """基础结构推断（按照参数卡片格式）"""
        parts = []
        
        # 导体推断
        if 'L' in model and model.startswith('YJ'):
            parts.append("AL")
        elif model.startswith('YJ') or '铜芯' in text:
            parts.append("CU")
        
        # 绝缘推断
        if 'YJ' in model or '交联聚乙烯' in text:
            # 检查是否耐火
            if 'NH' in model or '耐火' in text:
                parts.append("MT/XLPE")
            else:
                parts.append("XLPE")
        elif 'V' in model or '聚氯乙烯' in text:
            parts.append("PVC")
        
        # 屏蔽推断（中压电缆通常有屏蔽）
        voltage_in_text = text.upper()
        if any(v in voltage_in_text for v in ['6/10KV', '8.7/15KV', '12/20KV', '26/35KV', '10千伏', '35千伏']):
            if 'P2' in model:
                parts.append("CTS")
            elif 'P' in model:
                parts.append("CWS")
            else:
                parts.append("CTS")  # 中压默认铜带屏蔽
        
        # 铠装推断
        if '22' in model or '钢带铠装' in text:
            parts.append("STA")
        elif '32' in model or '钢丝铠装' in text:
            parts.append("SWA")
        elif '72' in model:
            parts.append("AWA")
        elif '23' in model:
            parts.append("STA")  # YJV23类型
        
        # 护套推断
        if model.endswith('V') or '聚氯乙烯护套' in text:
            parts.append("PVC")
        elif model.endswith('Y') or '聚乙烯护套' in text or '23' in model:
            parts.append("HDPE")
        elif 'LSZH' in text.upper() or 'WDZ' in model:
            parts.append("LSZH")
        else:
            parts.append("PVC")  # 默认PVC护套
        
        return "/".join(parts)
    
    def infer_model_from_description(self, text, voltage):
        """根据描述推断型号"""
        # 检查关键词
        text_upper = text.upper()
        
        # 耐火电缆
        if '耐火' in text or 'NH' in text_upper:
            if '铝芯' in text or 'AL' in text_upper:
                return 'NH-YJLV'
            else:
                return 'NH-YJV'
        
        # 低烟无卤
        if '低烟无卤' in text or 'WDZ' in text_upper or 'LSZH' in text_upper:
            if '铝芯' in text or 'AL' in text_upper:
                return 'WDZ-YJLV'
            else:
                return 'WDZ-YJV'
        
        # 根据结构特征推断
        has_armor = any(armor in text for armor in ['钢带铠装', '钢丝铠装', '铠装'])
        has_al = '铝芯' in text or 'AL' in text_upper
        has_xlpe = '交联聚乙烯' in text or 'XLPE' in text_upper
        
        if has_xlpe:
            if has_al:
                if has_armor:
                    if '钢带铠装' in text:
                        return 'YJLV22'
                    elif '钢丝铠装' in text:
                        return 'YJLV32'
                    else:
                        return 'YJLV22'  # 默认钢带
                else:
                    return 'YJLV'
            else:  # 铜芯
                if has_armor:
                    if '钢带铠装' in text:
                        return 'YJV22'
                    elif '钢丝铠装' in text:
                        return 'YJV32'
                    else:
                        return 'YJV22'  # 默认钢带
                else:
                    return 'YJV'
        
        # 默认返回空
        return ""

    def create_project_list_interface(self):
            """创建项目清单列表界面"""
            # 创建滚动区域容器
            scroll_container = tk.Frame(self.project_list_frame, bg="#f0f0f0")
            scroll_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # 创建主滚动框架
            main_canvas = tk.Canvas(scroll_container, bg="#f0f0f0")
            main_scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=main_canvas.yview)
            scrollable_frame = tk.Frame(main_canvas, bg="#f0f0f0")

            scrollable_frame.bind(
                "<Configure>",
                lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
            )

            main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            main_canvas.configure(yscrollcommand=main_scrollbar.set)

            # 标题
            title_label = tk.Label(scrollable_frame, text="项目清单列表统一管理", 
                                  font=("Microsoft YaHei", 16, "bold"), bg="#f0f0f0")
            title_label.pack(pady=15)

            # 筛选和搜索区域
            filter_frame = ttk.LabelFrame(scrollable_frame, text="🔍 筛选和搜索", padding=15)
            filter_frame.pack(fill=tk.X, padx=30, pady=(0, 15))

            # 第一行筛选
            filter_row1 = tk.Frame(filter_frame)
            filter_row1.pack(fill=tk.X, pady=5)

            # 月份筛选
            tk.Label(filter_row1, text="月份:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
            self.list_filter_month_var = tk.StringVar(value="全部")
            self.month_filter_combo = ttk.Combobox(filter_row1, textvariable=self.list_filter_month_var,
                                                   state="readonly", width=12)
            self.month_filter_combo.pack(side=tk.LEFT, padx=(5, 15))
            self.month_filter_combo.bind('<<ComboboxSelected>>', self.on_month_filter_change)

            # 项目筛选
            tk.Label(filter_row1, text="项目:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
            self.list_filter_project_var = tk.StringVar(value="全部")
            self.project_list_filter_combo = ttk.Combobox(filter_row1, textvariable=self.list_filter_project_var,
                                                         state="readonly", width=20)
            self.project_list_filter_combo.pack(side=tk.LEFT, padx=(5, 15))
            self.project_list_filter_combo.bind('<<ComboboxSelected>>', self.apply_list_filters)

            # 电压等级筛选
            tk.Label(filter_row1, text="电压等级:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
            self.list_filter_voltage_var = tk.StringVar(value="全部")
            self.voltage_filter_combo = ttk.Combobox(filter_row1, textvariable=self.list_filter_voltage_var,
                                                   state="readonly", width=12)
            self.voltage_filter_combo.pack(side=tk.LEFT, padx=(5, 15))
            self.voltage_filter_combo.bind('<<ComboboxSelected>>', self.apply_list_filters)

            # 搜索框
            tk.Label(filter_row1, text="搜索:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
            self.list_search_var = tk.StringVar()
            search_entry = tk.Entry(filter_row1, textvariable=self.list_search_var, width=25, font=("Microsoft YaHei", 11))
            search_entry.pack(side=tk.LEFT, padx=5)
            search_entry.bind('<KeyRelease>', self.apply_list_filters)

            # 清空筛选按钮
            tk.Button(filter_row1, text="清空筛选", command=self.clear_list_filters,
                     bg="#FF9800", fg="white", font=("Microsoft YaHei", 9), width=8).pack(side=tk.LEFT, padx=5)

            # 统计信息
            self.list_stats_label = tk.Label(filter_frame, text="", font=("Microsoft YaHei", 10), fg="#666")
            self.list_stats_label.pack(pady=5)

            # 清单列表
            list_frame = ttk.LabelFrame(scrollable_frame, text="📋 项目清单数据", padding=15)
            list_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=15)

            # 创建Treeview显示清单数据
            columns = ("项目编号", "行号", "电压等级", "报价型号", "结构描述", "备注")
            self.project_list_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

            # 初始化排序状态
            self.list_sort_column = None
            self.list_sort_reverse = False

            # 设置列标题和宽度（添加排序功能）
            for col in columns:
                self.project_list_tree.heading(col, text=col, 
                                              command=lambda c=col: self.sort_project_list(c))
            
            self.project_list_tree.column("项目编号", width=120)
            self.project_list_tree.column("行号", width=60)
            self.project_list_tree.column("电压等级", width=100)
            self.project_list_tree.column("报价型号", width=150)
            self.project_list_tree.column("结构描述", width=200)
            self.project_list_tree.column("备注", width=150)

            # 滚动条
            scrollbar_list = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.project_list_tree.yview)
            self.project_list_tree.configure(yscrollcommand=scrollbar_list.set)

            self.project_list_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar_list.pack(side=tk.RIGHT, fill=tk.Y)

            # 打包主滚动组件
            main_canvas.pack(side="left", fill="both", expand=True)
            main_scrollbar.pack(side="right", fill="y")

            # 绑定鼠标滚轮事件
            def _on_mousewheel(event):
                main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            main_canvas.bind_all("<MouseWheel>", _on_mousewheel)

            # 按钮框架 - 固定在窗口底部（不在滚动区域内）
            button_frame = tk.Frame(self.project_list_frame, bg="#f0f0f0")
            button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10, padx=20)
            
            # 创建按钮容器，居中显示
            button_container = tk.Frame(button_frame, bg="#f0f0f0")
            button_container.pack()

            tk.Button(button_container, text="🔄 刷新列表", command=self.refresh_project_list,
                     bg="#2196F3", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
            tk.Button(button_container, text="📊 导出清单", command=self.export_project_list,
                     bg="#4CAF50", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)
            tk.Button(button_container, text="🗑️ 清空列表", command=self.clear_project_list,
                     bg="#f44336", fg="white", font=("Microsoft YaHei", 10), width=12).pack(side=tk.LEFT, padx=5)

            # 加载项目清单数据
            self.load_project_list_data()


    def load_project_list_data(self):
        """加载项目清单数据"""
        try:
            # 获取所有项目的清单数据
            project_lists = self.config.get("project_lists", {})
            all_list_data = []
            projects = set()
            voltages = set()
            months = set()  # 新增：收集月份信息
            
            # 汇总所有项目的清单数据
            for project_code, project_data in project_lists.items():
                list_data = project_data.get("data", [])
                
                # 提取月份信息（从项目编号格式 yyyymmxxxx-LL）
                month_str = self.extract_month_from_project_code(project_code)
                if month_str:
                    months.add(month_str)
                
                for idx, item in enumerate(list_data, 1):
                    # 添加行号和项目编号
                    list_item = {
                        "项目编号": project_code,
                        "月份": month_str,  # 新增：添加月份字段
                        "行号": idx,
                        "电压等级": item.get("电压等级", ""),
                        "报价型号": item.get("报价型号", ""),
                        "结构描述": item.get("结构描述", ""),
                        "备注": item.get("备注", "")
                    }
                    all_list_data.append(list_item)
                    projects.add(project_code)
                    if item.get("电压等级"):
                        voltages.add(item.get("电压等级"))
            
            # 更新月份筛选选项
            month_options = ["全部"] + sorted(list(months), reverse=True)  # 按时间倒序
            self.month_filter_combo['values'] = month_options
            
            # 存储所有项目编号（用于级联筛选）
            self.all_projects_by_month = {}
            for project_code in projects:
                month_str = self.extract_month_from_project_code(project_code)
                if month_str not in self.all_projects_by_month:
                    self.all_projects_by_month[month_str] = []
                self.all_projects_by_month[month_str].append(project_code)
            
            # 更新项目筛选选项（根据当前选择的月份）
            self.update_project_filter_options()
            
            voltage_options = ["全部"] + sorted(list(voltages))
            self.voltage_filter_combo['values'] = voltage_options
            
            # 存储原始数据
            self.all_project_list_data = all_list_data
            
        except Exception as e:
            messagebox.showerror("错误", f"加载项目清单数据失败：{str(e)}")
    
    def update_project_filter_options(self):
        """更新项目筛选选项（根据选择的月份进行级联筛选）"""
        try:
            # 获取当前选择的月份
            month_filter = self.list_filter_month_var.get() if hasattr(self, 'list_filter_month_var') else "全部"
            
            # 根据月份筛选项目
            if month_filter == "全部":
                # 显示所有项目
                available_projects = []
                for projects_list in self.all_projects_by_month.values():
                    available_projects.extend(projects_list)
            else:
                # 只显示选中月份的项目
                available_projects = self.all_projects_by_month.get(month_filter, [])
            
            # 更新项目筛选下拉框
            project_options = ["全部"] + sorted(available_projects)
            self.project_list_filter_combo['values'] = project_options
            
            # 如果当前选择的项目不在新的选项中，重置为"全部"
            current_project = self.list_filter_project_var.get() if hasattr(self, 'list_filter_project_var') else "全部"
            if current_project not in project_options:
                self.list_filter_project_var.set("全部")
        
        except Exception as e:
            print(f"更新项目筛选选项失败: {e}")
    
    def on_month_filter_change(self, event=None):
        """月份筛选变化时的处理（级联更新项目筛选选项）"""
        # 更新项目筛选选项
        self.update_project_filter_options()
        # 应用筛选
        self.apply_list_filters()
    
    def extract_month_from_project_code(self, project_code):
        """从项目编号中提取月份（格式：yyyymmxxxx-LL）"""
        try:
            # 匹配格式：yyyymmxxxx-LL
            match = re.match(r'^(\d{4})(\d{2})\d{4}-[A-Z]+$', project_code)
            if match:
                year = match.group(1)
                month = match.group(2)
                return f"{year}-{month}"  # 返回格式：2025-12
            return None
        except Exception as e:
            print(f"提取月份失败 {project_code}: {e}")
            return None

    def load_single_project_list_data(self, project_code):
        """加载指定项目的清单数据"""
        try:
            project_lists = self.config.get("project_lists", {})
            project_data = project_lists.get(project_code, {})
            return project_data.get("data", [])
        except Exception as e:
            print(f"加载项目 {project_code} 清单数据失败: {e}")
            return []
    
    def generate_list_item_hash(self, item):
        """生成清单项的唯一哈希值（不包含规格字段）"""
        # 使用四个关键字段生成哈希（去掉规格字段）
        key_fields = [
            item.get('电压等级', '').strip(),
            item.get('报价型号', '').strip(),
            item.get('结构描述', '').strip(),
            item.get('备注', '').strip()
        ]
        
        # 创建标准化的字符串
        normalized_string = '|'.join(key_fields).lower()
        
        # 生成MD5哈希
        return hashlib.md5(normalized_string.encode('utf-8')).hexdigest()
    
    def deduplicate_list_data(self, new_list_data, project_code):
        """对清单数据进行项目内去重"""
        print(f"🔍 对项目 {project_code} 进行项目内去重...")
        
        # 项目内去重：只去除当前项目数据中的重复项
        seen_hashes = set()
        deduplicated_data = []
        duplicate_count = 0
        
        for item in new_list_data:
            item_hash = self.generate_list_item_hash(item)
            
            if item_hash not in seen_hashes:
                # 项目内唯一项，添加到结果中
                deduplicated_data.append(item)
                seen_hashes.add(item_hash)
            else:
                duplicate_count += 1
        
        print(f"   原始数据: {len(new_list_data)} 条")
        print(f"   去重后: {len(deduplicated_data)} 条")
        print(f"   项目内重复项: {duplicate_count} 条")
        
        return deduplicated_data, duplicate_count
    
    def get_all_unique_list_items(self):
        """获取所有项目的去重清单项"""
        project_lists = self.config.get("project_lists", {})
        unique_items = {}  # hash -> (item, projects)
        
        for project_code, project_data in project_lists.items():
            list_data = project_data.get("data", [])
            for item in list_data:
                item_hash = self.generate_list_item_hash(item)
                
                if item_hash not in unique_items:
                    unique_items[item_hash] = {
                        'item': item,
                        'projects': [project_code],
                        'first_seen': project_data.get('last_updated', 'Unknown')
                    }
                else:
                    # 记录在哪些项目中出现
                    if project_code not in unique_items[item_hash]['projects']:
                        unique_items[item_hash]['projects'].append(project_code)
        
        return unique_items
    def apply_list_filters(self, event=None):
        """应用清单筛选（增强型号搜索，支持月份筛选）"""
        try:
            if not hasattr(self, 'all_project_list_data'):
                return
            
            # 获取筛选条件
            month_filter = self.list_filter_month_var.get()  # 新增：月份筛选
            project_filter = self.list_filter_project_var.get()
            voltage_filter = self.list_filter_voltage_var.get()
            search_text = self.list_search_var.get().strip()
            
            # 应用筛选
            filtered_data = []
            for item in self.all_project_list_data:
                # 月份筛选（新增）
                if month_filter != "全部" and item.get("月份", "") != month_filter:
                    continue
                
                # 项目筛选
                if project_filter != "全部" and item.get("项目编号", "") != project_filter:
                    continue
                
                # 电压等级筛选
                if voltage_filter != "全部" and item.get("电压等级", "") != voltage_filter:
                    continue
                
                # 增强搜索筛选
                if search_text:
                    if self.enhanced_search_match(item, search_text):
                        filtered_data.append(item)
                else:
                    filtered_data.append(item)
            
            # 更新显示
            self.update_project_list_display(filtered_data)
            
            # 更新统计信息
            total_count = len(self.all_project_list_data)
            filtered_count = len(filtered_data)
            unique_models = len(set(item.get("报价型号", "") for item in filtered_data if item.get("报价型号", "")))
            unique_projects = len(set(item.get("项目编号", "") for item in filtered_data if item.get("项目编号", "")))
            
            stats_text = f"显示 {filtered_count}/{total_count} 条记录，涉及 {unique_projects} 个项目，{unique_models} 个型号"
            self.list_stats_label.config(text=stats_text)
            
        except Exception as e:
            print(f"应用筛选失败: {str(e)}")
    
    def enhanced_search_match(self, item, search_text):
        """增强搜索匹配（支持型号、结构描述、备注的模糊搜索）"""
        search_text = search_text.upper().strip()
        
        # 获取可搜索的字段
        searchable_fields = [
            item.get('报价型号', ''),
            item.get('结构描述', ''),
            item.get('备注', ''),
            item.get('电压等级', ''),
            item.get('项目编号', '')
        ]
        
        # 组合所有可搜索文本
        combined_text = ' '.join(str(field) for field in searchable_fields).upper()
        
        # 支持多关键词搜索（空格分隔）
        search_keywords = search_text.split()
        
        # 所有关键词都必须匹配
        for keyword in search_keywords:
            if keyword not in combined_text:
                return False
        
        return True
    def update_project_list_display(self, data):
        """更新项目清单显示（默认按项目编号倒序排序）"""
        # 清空现有数据
        for item in self.project_list_tree.get_children():
            self.project_list_tree.delete(item)
        
        # 按项目编号倒序排序（最新项目在前）
        sorted_data = sorted(data, key=lambda x: x.get("项目编号", ""), reverse=True)
        
        # 添加排序后的数据
        for item in sorted_data:
            self.project_list_tree.insert("", "end", values=(
                item["项目编号"],
                item["行号"],
                item["电压等级"],
                item["报价型号"],
                item["结构描述"],
                item["备注"]
            ))
    
    def sort_project_list(self, col):
        """按指定列排序项目清单列表"""
        try:
            # 如果点击同一列，则反转排序顺序
            if self.list_sort_column == col:
                self.list_sort_reverse = not self.list_sort_reverse
            else:
                self.list_sort_column = col
                self.list_sort_reverse = False
            
            # 获取所有数据
            data = []
            for item in self.project_list_tree.get_children():
                values = self.project_list_tree.item(item)['values']
                data.append((item, values))
            
            # 获取列索引
            columns = ("项目编号", "行号", "电压等级", "报价型号", "结构描述", "备注")
            col_index = columns.index(col)
            
            # 定义排序键函数
            def sort_key(item):
                value = item[1][col_index]
                # 处理数字列
                if col == "行号":
                    try:
                        return int(value) if value else 0
                    except:
                        return 0
                # 处理文本列（忽略大小写）
                return str(value).lower() if value else ""
            
            # 排序
            data.sort(key=sort_key, reverse=self.list_sort_reverse)
            
            # 重新插入排序后的数据
            for index, (item, values) in enumerate(data):
                self.project_list_tree.move(item, '', index)
            
            # 更新列标题显示排序指示器
            for c in columns:
                if c == col:
                    indicator = " ▼" if self.list_sort_reverse else " ▲"
                    self.project_list_tree.heading(c, text=c + indicator)
                else:
                    self.project_list_tree.heading(c, text=c, 
                                                  command=lambda c=c: self.sort_project_list(c))
            
        except Exception as e:
            messagebox.showerror("错误", f"排序失败：{str(e)}")

    def clear_list_filters(self):
        """清空清单筛选"""
        self.list_filter_month_var.set("全部")  # 新增：清空月份筛选
        self.list_filter_project_var.set("全部")
        self.list_filter_voltage_var.set("全部")
        self.list_search_var.set("")
        self.apply_list_filters()

    def refresh_project_list(self):
        """刷新项目清单列表"""
        self.load_project_list_data()
        messagebox.showinfo("提示", "项目清单列表已刷新")

    def export_project_list(self):
        """导出项目清单数据"""
        try:
            # 获取当前显示的数据
            displayed_data = []
            for item_id in self.project_list_tree.get_children():
                item = self.project_list_tree.item(item_id)
                displayed_data.append(item['values'])
            
            if not displayed_data:
                messagebox.showwarning("警告", "没有数据可导出！")
                return
            
            # 选择保存位置
            filename = filedialog.asksaveasfilename(
                title="导出项目清单数据",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel文件", "*.xlsx"),
                    ("CSV文件", "*.csv"),
                    ("所有文件", "*.*")
                ]
            )
            
            if not filename:
                return
            
            # 准备数据（与树形控件列定义一致）
            columns = ["项目编号", "行号", "电压等级", "报价型号", "结构描述", "备注"]
            
            if filename.endswith('.csv'):
                # 导出CSV
                import csv
                with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(columns)
                    writer.writerows(displayed_data)
            else:
                # 导出Excel
                df = pd.DataFrame(displayed_data, columns=columns)
                df.to_excel(filename, index=False, sheet_name='项目清单汇总')
            
            messagebox.showinfo("成功", f"项目清单数据已导出到：\n{filename}\n\n导出记录数：{len(displayed_data)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def clear_project_list(self):
        """清空项目清单数据"""
        if messagebox.askyesno("确认清空", "确定要清空所有项目清单数据吗？\n\n注意：这将删除所有已导入的清单数据，但不会影响项目文件夹。"):
            try:
                self.config["project_lists"] = {}
                self.save_config()
                self.load_project_list_data()
                messagebox.showinfo("提示", "项目清单数据已清空")
            except Exception as e:
                messagebox.showerror("错误", f"清空失败：{str(e)}")
    

    def import_projects_from_folder(self):
        """从文件夹导入项目（支持单个项目文件夹和批量导入）"""
        from tkinter import filedialog, messagebox
        
        # 选择项目文件夹
        folder_path = filedialog.askdirectory(
            title="选择项目文件夹或包含项目文件夹的目录",
            initialdir=self.config.get("default_project_folder", "")
        )
        
        if not folder_path:
            return
        
        try:
            # 智能检测选择的文件夹类型
            imported_projects = self.smart_scan_projects(folder_path)
            
            if not imported_projects:
                messagebox.showinfo("提示", "在选择的目录中没有找到有效的项目文件夹。")
                return
            
            # 显示导入预览对话框
            self.show_import_preview_dialog(imported_projects, folder_path)
            
        except Exception as e:
            messagebox.showerror("错误", f"扫描项目文件夹时出错：{str(e)}")
    

    def smart_scan_projects(self, folder_path):
        """智能扫描项目（支持单个项目文件夹和批量扫描）"""
        imported_projects = []
        
        try:
            # 首先检查选择的文件夹是否本身就是一个项目文件夹
            folder_name = os.path.basename(folder_path)
            single_project = self.check_single_project_folder(folder_path, folder_name)
            
            if single_project:
                # 选择的是单个项目文件夹
                print(f"✅ 检测到单个项目文件夹: {folder_name}")
                imported_projects.append(single_project)
            else:
                # 选择的是包含多个项目的文件夹，进行批量扫描
                print(f"🔍 批量扫描项目文件夹: {folder_path}")
                imported_projects = self.scan_project_folders(folder_path)
        
        except Exception as e:
            print(f"智能扫描失败: {e}")
        
        return imported_projects
    
    def check_single_project_folder(self, folder_path, folder_name):
        """检查是否为单个项目文件夹"""
        try:
            # 1. 检查文件夹名称是否符合项目命名规范
            project_info = self.parse_project_folder_name(folder_name)
            if not project_info:
                return None
            
            # 2. 检查文件夹内是否包含项目相关的子文件夹
            has_project_structure = False
            
            if os.path.isdir(folder_path):
                for item in os.listdir(folder_path):
                    item_path = os.path.join(folder_path, item)
                    if os.path.isdir(item_path):
                        # 检查是否包含项目相关文件夹
                        if any(keyword in item for keyword in ['清单', '定额', '技术规范', '项目资料']):
                            has_project_structure = True
                            break
            
            if not has_project_structure:
                return None
            
            # 3. 分析项目内容
            print(f"📊 分析单个项目: {project_info['code']}")
            analysis = self.analyze_project_structure(folder_path)
            
            model_count = analysis.get("model_count", 0)
            spec_count = analysis.get("spec_count", 0)
            list_data = analysis.get("list_data", [])
            
            # 4. 检查是否应该跳过分析
            skip_analysis = self.should_skip_project_analysis(project_info['code'])
            
            project_data = {
                'code': project_info['code'],
                'name': project_info['name'],
                'manager': project_info['manager'],
                'folder_path': folder_path,
                'model_count': model_count,
                'spec_count': spec_count,
                'list_data': list_data,
                'skip_analysis': skip_analysis,
                'import_type': 'single'  # 标记为单个项目导入
            }
            
            print(f"✅ 单个项目分析完成: {model_count}个型号, {spec_count}个技术规范")
            return project_data
            
        except Exception as e:
            print(f"检查单个项目文件夹失败 {folder_name}: {e}")
            return None
    def scan_project_folders(self, root_folder):
        """扫描项目文件夹"""
        imported_projects = []
        
        try:
            for item in os.listdir(root_folder):
                item_path = os.path.join(root_folder, item)
                
                # 只处理文件夹
                if not os.path.isdir(item_path):
                    continue
                
                # 解析文件夹名称
                project_info = self.parse_project_folder_name(item)
                if not project_info:
                    continue
                
                # 统计清单文件和技术规范
                model_count, spec_count, list_data = self.analyze_project_folder(item_path)
                
                # 检查是否应该跳过分析
                skip_analysis = self.should_skip_project_analysis(project_info['code'])
                
                project_data = {
                    'code': project_info['code'],
                    'name': project_info['name'],
                    'manager': project_info['manager'],
                    'folder_path': item_path,
                    'model_count': model_count,
                    'spec_count': spec_count,
                    'list_data': list_data,
                    'skip_analysis': skip_analysis
                }
                
                imported_projects.append(project_data)
        
        except Exception as e:
            print(f"扫描文件夹时出错: {e}")
        
        return imported_projects
    
    
    def parse_folder_name(self, folder_name, folder_path):
        """解析项目文件夹名称并分析内容（兼容测试接口）"""
        try:
            # 解析文件夹名称
            project_info = self.parse_project_folder_name(folder_name)
            if not project_info:
                return None
            
            # 分析项目文件夹内容
            analysis = self.analyze_project_structure(folder_path)
            model_count = analysis.get("model_count", 0)
            spec_count = analysis.get("spec_count", 0)
            list_data = analysis.get("list_data", [])
            
            # 返回完整的项目信息
            result = {
                'code': project_info['code'],
                'name': project_info['name'],
                'manager': project_info['manager'],
                'folder_path': folder_path,
                'model_count': model_count,
                'spec_count': spec_count,
                'list_data': list_data
            }
            
            return result
            
        except Exception as e:
            print(f"解析项目文件夹失败 {folder_name}: {e}")
            return None

    def parse_project_folder_name(self, folder_name):
        """解析项目文件夹名称"""
        # 移除开头的括号内容（如：（陈颖））
        clean_name = re.sub(r'^[（(][^）)]*[）)]\s*', '', folder_name)
        
        # 尝试多种分隔符模式
        patterns = [
            r'^([A-Z0-9\-_]+)\s+(.+?)\s+([^\s]+)$',  # 项目编号 项目名称 业务员
            r'^([A-Z0-9\-_]+)[-_\s](.+?)[-_\s]([^\s\-_]+)$',  # 项目编号-项目名称-业务员
        ]
        
        for pattern in patterns:
            match = re.match(pattern, clean_name)
            if match:
                return {
                    'code': match.group(1).strip(),
                    'name': match.group(2).strip(),
                    'manager': match.group(3).strip()
                }
        
        return None
    
    def analyze_project_folder(self, folder_path):
        """分析项目文件夹，统计清单和技术规范"""
        model_count = 0
        spec_count = 0
        list_data = []
        unique_models = set()  # 用于去重统计
        
        try:
            # 查找清单定额文件夹
            quota_folder = None
            for item in os.listdir(folder_path):
                if '清单' in item or '定额' in item:
                    quota_folder = os.path.join(folder_path, item)
                    break
            
            if quota_folder and os.path.isdir(quota_folder):
                # 统计Excel文件并解析清单数据
                for file in os.listdir(quota_folder):
                    if file.endswith(('.xlsx', '.xls')):
                        file_path = os.path.join(quota_folder, file)
                        try:
                            # 解析Excel文件
                            parsed_data = self.parse_excel_list_file(file_path)
                            if parsed_data:
                                list_data.extend(parsed_data)
                                
                                # 统计去重的报价型号
                                for item in parsed_data:
                                    model = item.get('报价型号', '').strip()
                                    if model and model not in ['', 'nan', 'None']:
                                        unique_models.add(model)
                                        
                        except Exception as e:
                            print(f"解析Excel文件失败 {file}: {e}")
            
            # 型号数量为去重后的数量
            model_count = len(unique_models)
            
            # 查找技术规范文件夹 - 修复：只查找名为"技术规范"的文件夹
            spec_folder = None
            for item in os.listdir(folder_path):
                if '技术规范' in item:  # 只匹配包含"技术规范"的文件夹
                    spec_folder = os.path.join(folder_path, item)
                    break
            
            # 修复：只有找到技术规范文件夹才统计Word文档
            if spec_folder and os.path.isdir(spec_folder):
                # 统计技术规范文件 - 修复：只统计Word文档，不包括PDF
                for file in os.listdir(spec_folder):
                    if file.endswith(('.docx', '.doc')):  # 移除了.pdf
                        spec_count += 1
                        print(f"✓ 统计技术规范文件: {file}")
            else:
                print(f"ℹ️ 项目 {os.path.basename(folder_path)} 没有找到技术规范文件夹，技术规范数量为0")
        
        except Exception as e:
            print(f"分析项目文件夹失败 {folder_path}: {e}")
        
        return model_count, spec_count, list_data
    
    def should_skip_project_analysis(self, project_code):
        """检查是否应该跳过项目分析"""
        skip_list = self.config.get("skip_analysis_projects", [])
        return project_code in skip_list
    
    def parse_excel_list_file(self, file_path):
        """解析Excel清单文件（修复版 - 支持深度搜索连续五列）"""
        try:
            import pandas as pd
            import warnings
            
            # 抑制openpyxl警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                
                # 读取Excel文件的所有工作表
                excel_file = pd.ExcelFile(file_path)
                all_data = []
                
                for sheet_name in excel_file.sheet_names:
                    try:
                        # 读取工作表（不设置header，以便深度搜索）
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                        
                        if df.empty:
                            continue
                        
                        # 查找连续的五列：电压等级 报价型号 规格 结构描述 备注
                        header_row, start_col = self.find_five_consecutive_columns(df)
                        
                        if header_row is not None and start_col is not None:
                            # 提取数据
                            sheet_data = self.extract_five_column_data(df, header_row, start_col, file_path, sheet_name)
                            if sheet_data:
                                all_data.extend(sheet_data)
                    
                    except Exception as e:
                        print(f"解析工作表 {sheet_name} 失败: {e}")
                        continue
                
                return all_data
        
        except Exception as e:
            print(f"解析Excel文件失败 {file_path}: {e}")
            return []
    
    def find_five_consecutive_columns(self, df, max_search_rows=50):
        """查找连续的五列：电压等级 报价型号 规格 结构描述 备注"""
        target_columns = ['电压等级', '报价型号', '规格', '结构描述', '备注']
        search_rows = min(max_search_rows, len(df))
        
        for row_idx in range(search_rows):
            # 获取这一行的所有值
            row_values = []
            for col_idx in range(len(df.columns)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    row_values.append(cell_str)
                else:
                    row_values.append("")
            
            # 查找连续的五列
            for start_col in range(len(row_values) - 4):  # 至少需要5列
                consecutive_values = row_values[start_col:start_col + 5]
                
                # 检查这五个连续的值是否匹配目标列名
                matches = 0
                for i, target_col in enumerate(target_columns):
                    if target_col in consecutive_values[i]:
                        matches += 1
                
                if matches == 5:  # 完全匹配
                    return row_idx, start_col
        
        return None, None
    
    def extract_five_column_data(self, df, header_row, start_col, file_path, sheet_name):
        """提取五列数据"""
        list_data = []
        target_columns = ['电压等级', '报价型号', '规格', '结构描述', '备注']
        
        for row_idx in range(header_row + 1, len(df)):
            row_data = {}
            
            # 提取五列数据
            for i, col_name in enumerate(target_columns):
                col_idx = start_col + i
                if col_idx < len(df.columns):
                    cell_value = df.iloc[row_idx, col_idx]
                    if pd.notna(cell_value):
                        row_data[col_name] = str(cell_value).strip()
                    else:
                        row_data[col_name] = ""
                else:
                    row_data[col_name] = ""
            
            # 检查是否为有效数据行
            if self.is_valid_list_row(row_data):
                row_data['来源文件'] = os.path.basename(file_path)
                row_data['来源工作表'] = sheet_name
                list_data.append(row_data)
            elif self.is_data_end_marker(row_data):
                break
        
        return list_data
    
    def is_valid_list_row(self, row_data):
        """检查是否为有效的清单数据行"""
        # 至少报价型号或规格不能为空
        model = row_data.get('报价型号', '').strip()
        spec = row_data.get('规格', '').strip()
        
        if not model and not spec:
            return False
        
        # 排除明显的表头重复
        if model in ['报价型号', '电压等级', '规格', '结构描述', '备注']:
            return False
        
        # 排除明显的非数据行
        exclude_patterns = ['合计', '小计', '总计', '备注说明', '注：']
        for pattern in exclude_patterns:
            if pattern in model or pattern in spec:
                return False
        
        return True
    
    def is_data_end_marker(self, row_data):
        """检查是否为数据结束标记"""
        end_markers = ['合计', '小计', '总计', 'total', 'sum', '备注说明', '注：']
        
        for value in row_data.values():
            if value and any(marker in value.lower() for marker in end_markers):
                return True
        
        return False

    def show_import_preview_dialog(self, imported_projects, source_folder):
        """显示导入预览对话框"""
        dialog = tk.Toplevel(self.root)
        # 根据导入类型设置标题
        if len(imported_projects) == 1 and imported_projects[0].get('import_type') == 'single':
            dialog.title(f"导入单个项目: {imported_projects[0]['code']}")
        else:
            dialog.title("批量项目导入预览")
        dialog.geometry("1000x700")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (1000 // 2)
        y = (dialog.winfo_screenheight() // 2) - (700 // 2)
        dialog.geometry(f"1000x700+{x}+{y}")
        
        # 主框架
        main_frame = tk.Frame(dialog, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        title_label = tk.Label(main_frame, text="项目导入预览", 
                              font=("Microsoft YaHei", 16, "bold"), bg="#f0f0f0", fg="#333")
        title_label.pack(pady=(0, 10))
        
        # 信息标签
        info_label = tk.Label(main_frame, text=f"扫描路径：{source_folder}\n找到 {len(imported_projects)} 个项目文件夹", 
                             font=("Microsoft YaHei", 11), bg="#f0f0f0", fg="#666")
        info_label.pack(pady=(0, 15))
        
        # 检查冲突
        conflicts = []
        recent_projects = self.config.get("recent_projects", [])
        for project in imported_projects:
            for existing in recent_projects:
                if existing.get("code") == project["code"]:
                    conflicts.append({
                        "project": project,
                        "existing": existing
                    })
                    break
        
        if conflicts:
            conflict_label = tk.Label(main_frame, 
                                    text=f"⚠️ 发现 {len(conflicts)} 个项目编号冲突，将更新现有记录", 
                                    font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#ff6600")
            conflict_label.pack(pady=(0, 10))
        
        # 项目列表框架
        list_frame = tk.Frame(main_frame, bg="#f0f0f0")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 创建Treeview
        columns = ("项目编号", "项目名称", "业务员", "型号数量", "技术规范数量", "状态")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题和宽度
        tree.heading("项目编号", text="项目编号")
        tree.heading("项目名称", text="项目名称")
        tree.heading("业务员", text="业务员")
        tree.heading("型号数量", text="型号数量")
        tree.heading("技术规范数量", text="技术规范数量")
        tree.heading("状态", text="状态")
        
        tree.column("项目编号", width=120)
        tree.column("项目名称", width=300)
        tree.column("业务员", width=100)
        tree.column("型号数量", width=80)
        tree.column("技术规范数量", width=100)
        tree.column("状态", width=100)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 填充数据
        for i, project in enumerate(imported_projects):
            # 确定状态
            status = "新增"
            for conflict in conflicts:
                if conflict["project"]["code"] == project["code"]:
                    status = "更新"
                    break
            
            if project.get('skip_analysis', False):
                status += " (跳过分析)"
            
            tree.insert("", "end", values=(
                project["code"],
                project["name"][:40] + "..." if len(project["name"]) > 40 else project["name"],
                project["manager"],
                project["model_count"],
                project["spec_count"],
                status
            ))
        
        # 按钮框架
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        def confirm_import():
            """确认导入"""
            try:
                imported_count = 0
                updated_count = 0
                
                for project in imported_projects:
                    # 检查是否已存在
                    existing = False
                    for existing_project in recent_projects:
                        if existing_project.get("code") == project["code"]:
                            existing = True
                            updated_count += 1
                            break
                    
                    if not existing:
                        imported_count += 1
                    
                    # 添加到近期项目列表
                    self.add_recent_project(
                        project["code"],
                        project["name"],
                        project["manager"],
                        project["folder_path"],
                        project["model_count"],
                        project["spec_count"]
                    )
                    
                    # 保存清单数据
                    if project.get("list_data"):
                        self.save_project_list_data(project["code"], project["list_data"])
                
                dialog.destroy()
                messagebox.showinfo("导入完成", 
                                  f"项目导入完成！\n\n新增项目：{imported_count} 个\n更新项目：{updated_count} 个")
                
            except Exception as e:
                messagebox.showerror("导入失败", f"导入项目时出错：{str(e)}")
        
        def cancel_import():
            """取消导入"""
            dialog.destroy()
        
        # 按钮
        tk.Button(button_frame, text="✅ 确认导入", command=confirm_import,
                 bg="#4CAF50", fg="white", font=("Microsoft YaHei", 12), width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="❌ 取消", command=cancel_import,
                 bg="#f44336", fg="white", font=("Microsoft YaHei", 12), width=15).pack(side=tk.LEFT, padx=10)
        
        # 统计信息
        stats_label = tk.Label(button_frame, 
                              text=f"总计：{len(imported_projects)} 个项目，冲突：{len(conflicts)} 个", 
                              font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#666")
        stats_label.pack(side=tk.RIGHT, padx=10)
    
    def save_project_list_data(self, project_code, list_data, enable_deduplication=True):
        """保存项目清单数据（支持项目内去重，保留创建时间）"""
        if not list_data:
            return {"success": False, "message": "无数据需要保存"}
        
        original_count = len(list_data)
        duplicate_count = 0
        
        # 如果启用去重功能
        if enable_deduplication:
            list_data, duplicate_count = self.deduplicate_list_data(list_data, project_code)
        
        # 获取现有项目数据以保留创建时间
        project_lists = self.config.get("project_lists", {})
        existing_project = project_lists.get(project_code, {})
        
        # 保留原有的创建时间，如果不存在则创建新的
        created_time = existing_project.get("created_time")
        if not created_time:
            created_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 保存去重后的数据
        project_lists[project_code] = {
            "data": list_data,
            "created_time": created_time,  # 保留或新建创建时间
            "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # 更新时间
            "original_count": original_count,
            "duplicate_count": duplicate_count
        }
        
        self.config["project_lists"] = project_lists
        self.save_config()
        
        result = {
            "success": True,
            "original_count": original_count,
            "saved_count": len(list_data),
            "duplicate_count": duplicate_count,
            "created_time": created_time,
            "message": f"保存成功：原始{original_count}条，项目内去重后{len(list_data)}条，重复{duplicate_count}条"
        }
        
        print(f"✅ {result['message']}")
        print(f"   创建时间: {created_time}")
        return result


    def get_unique_models_from_list_data(self, list_data):
        """从清单数据中提取去重的型号列表（用于测试和验证）"""
        unique_models = set()
        for item in list_data:
            model = item.get('报价型号', '').strip()
            if model and model not in ['', 'nan', 'None']:
                unique_models.add(model)
        return unique_models
    
    def validate_model_count_consistency(self, folder_path):
        """验证型号数量统计的一致性（用于测试）"""
        try:
            analysis = self.analyze_project_structure(folder_path)
            model_count = analysis.get("model_count", 0)
            spec_count = analysis.get("spec_count", 0)
            list_data = analysis.get("list_data", [])
            unique_models = self.get_unique_models_from_list_data(list_data)
            
            # 检查一致性
            is_consistent = model_count == len(unique_models)
            
            return {
                'is_consistent': is_consistent,
                'model_count': model_count,
                'unique_models_count': len(unique_models),
                'unique_models': unique_models,
                'total_rows': len(list_data),
                'spec_count': spec_count
            }
        except Exception as e:
            return {
                'is_consistent': False,
                'error': str(e)
            }


    def analyze_project_structure(self, project_path):
        """分析项目文件夹结构（最终版实现）"""
        analysis = {
            "list_count": 0,
            "spec_count": 0,
            "model_count": 0,
            "list_files": [],
            "list_data": [],
            "unique_models": set(),
            "has_quota_folder": False,
            "has_spec_folder": False
        }
        
        try:
            # 检查清单定额文件夹
            quota_path = os.path.join(project_path, "清单定额")
            if os.path.exists(quota_path):
                analysis["has_quota_folder"] = True
                
                # 查找清单文件
                for file in os.listdir(quota_path):
                    if file.endswith(('.xlsx', '.xls', '.xlsm')) and self.is_list_file(file):
                        file_path = os.path.join(quota_path, file)
                        try:
                            # 解析Excel文件
                            file_data = self.parse_excel_list_file(file_path)
                            if file_data:
                                # 文件信息
                                file_info = {
                                    "filename": file,
                                    "row_count": len(file_data),
                                    "has_expected_structure": True,
                                    "columns": ["电压等级", "报价型号", "规格", "结构描述", "备注"]
                                }
                                
                                # 检查是否是多工作表文件
                                if isinstance(file_data, dict) and 'sheets' in file_data:
                                    # 多工作表文件
                                    file_info.update({
                                        "sheet_count": file_data['sheet_count'],
                                        "valid_sheets": file_data['valid_sheets'],
                                        "sheet_names": file_data['sheet_names'],
                                        "total_rows": len(file_data['data']),
                                        "unique_models": len(file_data['unique_models'])
                                    })
                                    analysis["list_data"].extend(file_data['data'])
                                    analysis["unique_models"].update(file_data['unique_models'])
                                else:
                                    # 单工作表文件
                                    analysis["list_data"].extend(file_data)
                                    # 提取报价型号
                                    for item in file_data:
                                        model = item.get('报价型号', '').strip()
                                        if model and model not in ['', 'nan', 'None', '报价型号']:
                                            analysis["unique_models"].add(model)
                                
                                analysis["list_files"].append(file_info)
                                analysis["list_count"] += len(file_data) if isinstance(file_data, list) else len(file_data.get('data', []))
                                
                        except Exception as e:
                            print(f"解析清单文件失败 {file}: {e}")
            
            # 检查技术规范文件夹
            spec_path = os.path.join(project_path, "技术规范")
            if os.path.exists(spec_path):
                analysis["has_spec_folder"] = True
                
                # 统计Word文档数量
                for file in os.listdir(spec_path):
                    if file.endswith(('.doc', '.docx')):
                        analysis["spec_count"] += 1
            
            # 型号数量为去重后的数量
            analysis["model_count"] = len(analysis["unique_models"])
            
        except Exception as e:
            print(f"分析项目文件夹失败 {project_path}: {e}")
        
        return analysis
    
    def is_list_file(self, filename):
        """判断文件是否为清单文件"""
        list_keywords = ['清单', 'list', 'List', '报价']
        return any(keyword in filename for keyword in list_keywords)
    
    

if __name__ == "__main__":
    app = CableDesignSystemV4()
    app.root.mainloop()