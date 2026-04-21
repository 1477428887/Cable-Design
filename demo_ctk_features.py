#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CustomTkinter 版本特性演示
展示新界面的主要改进点
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox

# 设置外观
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class CTKFeaturesDemo:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("CustomTkinter 特性演示")
        self.root.geometry("800x600")
        
        self.create_demo_interface()
        
    def create_demo_interface(self):
        """创建演示界面"""
        # 标题
        title = ctk.CTkLabel(self.root, 
                            text="电缆设计系统 CustomTkinter 版本特性", 
                            font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=20)
        
        # 特性展示区域
        features_frame = ctk.CTkScrollableFrame(self.root, label_text="主要改进特性")
        features_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 特性1：现代化按钮
        feature1_frame = ctk.CTkFrame(features_frame)
        feature1_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(feature1_frame, text="1. 现代化按钮设计", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        button_demo_frame = ctk.CTkFrame(feature1_frame)
        button_demo_frame.pack(pady=10)
        
        ctk.CTkButton(button_demo_frame, text="主要按钮", 
                     command=lambda: messagebox.showinfo("演示", "这是主要按钮")).pack(side="left", padx=10)
        ctk.CTkButton(button_demo_frame, text="次要按钮", 
                     fg_color="transparent", border_width=2,
                     command=lambda: messagebox.showinfo("演示", "这是次要按钮")).pack(side="left", padx=10)
        
        # 特性2：现代化输入框
        feature2_frame = ctk.CTkFrame(features_frame)
        feature2_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(feature2_frame, text="2. 现代化输入框", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        input_demo_frame = ctk.CTkFrame(feature2_frame)
        input_demo_frame.pack(pady=10)
        
        ctk.CTkEntry(input_demo_frame, placeholder_text="输入项目编号...", 
                    width=200).pack(side="left", padx=10)
        ctk.CTkEntry(input_demo_frame, placeholder_text="输入项目名称...", 
                    width=300).pack(side="left", padx=10)
        
        # 特性3：下拉框和选择器
        feature3_frame = ctk.CTkFrame(features_frame)
        feature3_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(feature3_frame, text="3. 现代化选择器", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        selector_demo_frame = ctk.CTkFrame(feature3_frame)
        selector_demo_frame.pack(pady=10)
        
        ctk.CTkComboBox(selector_demo_frame, values=["低压电力电缆", "中压电力电缆", "控制电缆"], 
                       width=150).pack(side="left", padx=10)
        
        radio_var = tk.StringVar(value="铜")
        ctk.CTkRadioButton(selector_demo_frame, text="铜导体", 
                          variable=radio_var, value="铜").pack(side="left", padx=10)
        ctk.CTkRadioButton(selector_demo_frame, text="铝导体", 
                          variable=radio_var, value="铝").pack(side="left", padx=10)
        
        # 特性4：外观模式切换
        feature4_frame = ctk.CTkFrame(features_frame)
        feature4_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(feature4_frame, text="4. 外观模式切换", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        appearance_frame = ctk.CTkFrame(feature4_frame)
        appearance_frame.pack(pady=10)
        
        ctk.CTkLabel(appearance_frame, text="外观模式:").pack(side="left", padx=10)
        appearance_menu = ctk.CTkOptionMenu(appearance_frame, 
                                           values=["Light", "Dark", "System"],
                                           command=self.change_appearance)
        appearance_menu.pack(side="left", padx=10)
        
        ctk.CTkLabel(appearance_frame, text="界面缩放:").pack(side="left", padx=10)
        scaling_menu = ctk.CTkOptionMenu(appearance_frame, 
                                        values=["80%", "90%", "100%", "110%", "120%"],
                                        command=self.change_scaling)
        scaling_menu.pack(side="left", padx=10)
        
        # 特性5：进度条和滑块
        feature5_frame = ctk.CTkFrame(features_frame)
        feature5_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(feature5_frame, text="5. 进度条和滑块", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        progress_frame = ctk.CTkFrame(feature5_frame)
        progress_frame.pack(pady=10)
        
        self.progress = ctk.CTkProgressBar(progress_frame, width=300)
        self.progress.pack(side="left", padx=10)
        self.progress.set(0.7)
        
        self.slider = ctk.CTkSlider(progress_frame, from_=0, to=1, 
                                   command=self.update_progress, width=200)
        self.slider.pack(side="left", padx=10)
        self.slider.set(0.7)
        
        # 特性6：文本框
        feature6_frame = ctk.CTkFrame(features_frame)
        feature6_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(feature6_frame, text="6. 现代化文本框", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        self.textbox = ctk.CTkTextbox(feature6_frame, height=100, width=600)
        self.textbox.pack(pady=10)
        self.textbox.insert("1.0", "这是一个现代化的文本框，支持滚动和多行文本编辑。\n" +
                                  "CustomTkinter 提供了更好的视觉效果和用户体验。\n" +
                                  "界面更加简洁、现代，符合当前的设计趋势。")
        
        # 底部按钮
        bottom_frame = ctk.CTkFrame(self.root)
        bottom_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkButton(bottom_frame, text="启动完整版本", 
                     command=self.launch_full_version,
                     font=ctk.CTkFont(size=14, weight="bold"),
                     height=40).pack(side="right", padx=10)
        
        ctk.CTkButton(bottom_frame, text="关闭演示", 
                     command=self.root.quit,
                     fg_color="gray", hover_color="darkgray",
                     height=40).pack(side="right", padx=10)
    
    def change_appearance(self, mode):
        """改变外观模式"""
        ctk.set_appearance_mode(mode)
        
    def change_scaling(self, scaling):
        """改变缩放"""
        scaling_float = int(scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(scaling_float)
        
    def update_progress(self, value):
        """更新进度条"""
        self.progress.set(value)
        
    def launch_full_version(self):
        """启动完整版本"""
        try:
            import subprocess
            subprocess.Popen(["python", "cable_design_system_v4_ctk.py"])
            messagebox.showinfo("启动", "正在启动完整版本...")
        except Exception as e:
            messagebox.showerror("错误", f"启动失败：{str(e)}")
    
    def run(self):
        """运行演示"""
        self.root.mainloop()

if __name__ == "__main__":
    demo = CTKFeaturesDemo()
    demo.run()