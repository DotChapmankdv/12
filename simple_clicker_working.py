#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单可靠的自动点击工具
"""

import pyautogui
import tkinter as tk
from tkinter import messagebox, filedialog
import time
import threading
import json
import os
import sys
import subprocess
import ctypes
import ctypes.wintypes

class SimpleClickerWorking:
    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("智能自动点击工具 Pro")

        # 设置真实缩放因子（90%缩放，确保完整显示）
        self.scale_factor = 0.90

        # 原版设计尺寸（完整界面所需的最小尺寸）
        design_width = 1800
        design_height = 1300

        # 计算缩放后的窗口大小
        scaled_width = int(design_width * self.scale_factor)
        scaled_height = int(design_height * self.scale_factor)

        # 设置窗口大小 - 确保能完整显示所有内容
        self.root.geometry(f"{scaled_width}x{scaled_height}")
        self.root.resizable(True, True)
        self.root.minsize(int(scaled_width * 0.9), int(scaled_height * 0.9))

        # 窗口居中
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (scaled_width // 2)
        y = (screen_height // 2) - (scaled_height // 2)
        self.root.geometry(f"{scaled_width}x{scaled_height}+{x}+{y}")

        # 字体和尺寸缩放辅助函数
        def scaled_size(original_size):
            """计算缩放后的尺寸"""
            return max(1, int(original_size * self.scale_factor))

        def scaled_font(size):
            """返回缩放后的字体大小"""
            return max(8, int(size * self.scale_factor))

        self.scaled_size = scaled_size
        self.scaled_font = scaled_font

        # 数据
        self.actions = {}
        self.current_action = None
        self.is_running = False

        # 设置GUI
        self.setup_gui()

    def scaled_size(self, original_size):
        """计算缩放后的尺寸"""
        return int(original_size * self.scale_factor)

    def scaled_font(self, original_size):
        """计算缩放后的字体大小"""
        return int(original_size * self.scale_factor)

    def setup_gui(self):
        """设置GUI界面（75%缩放版本，包含完整三块区域）"""

        # 主容器 - 超紧凑间距
        main_container = tk.Frame(self.root, bg='#f8f9fa')
        main_container.pack(fill=tk.BOTH, expand=True, padx=self.scaled_size(20), pady=self.scaled_size(5))

        # 动作设置区域 - 缩放版本
        action_frame = tk.LabelFrame(main_container, text="动作设置",
                                   font=("Arial", self.scaled_font(14), "bold"),
                                   fg='#2c3e50', bg='white')
        action_frame.pack(fill=tk.X, pady=(self.scaled_size(5), self.scaled_size(3)))

        self.status_labels = {}

        # 创建12个动作
        action_colors = {
            1: '#3498db',  # 蓝色
            2: '#27ae60',  # 绿色
            3: '#f39c12',  # 橙色
            4: '#9b59b6',  # 紫色
            5: '#e74c3c',  # 红色 - 特殊动作
            6: '#16a085',  # 青色
            7: '#d35400',  # 深橙色
            8: '#8e44ad',  # 深紫色 - 文件操作
            9: '#2980b9',  # 深蓝色 - 复制操作
            10: '#c0392b', # 深红色 - URL操作
            11: '#f39c12', # 橙色 - 右击操作
            12: '#27ae60'  # 绿色 - 键盘粘贴
        }

        for i in range(1, 13):
            # 动作行 - 缩放版本
            row_frame = tk.Frame(action_frame, bg='white')
            row_frame.pack(fill=tk.X, padx=self.scaled_size(20), pady=self.scaled_size(4))

            # 动作名
            action_name = f"动作{i}"

            # 设置动作描述
            if i == 5:
                name_label_text = f"[延迟] {action_name} (延迟5秒后执行)"
                text_color = '#e74c3c'
            elif i == 6:
                name_label_text = f"[延迟] {action_name} (延迟5秒后执行)"
                text_color = '#16a085'
            elif i == 7:
                name_label_text = f"[最后] {action_name} (最后动作)"
                text_color = '#d35400'
            elif i == 8:
                name_label_text = f"[文档] {action_name} (双击打开TXT文档)"
                text_color = '#8e44ad'
            elif i == 9:
                name_label_text = f"[复制] {action_name} (智能文档识别复制)"
                text_color = '#2980b9'
            elif i == 10:
                name_label_text = f"[链接] {action_name} (添加URL)"
                text_color = '#c0392b'
            elif i == 11:
                name_label_text = f"[新增] {action_name} (自动右击)"
                text_color = '#f39c12'
            elif i == 12:
                name_label_text = f"[新增] {action_name} (键盘粘贴)"
                text_color = '#27ae60'
            else:
                name_label_text = f"{action_name}:"
                text_color = 'black'

            bg_color = action_colors[i]

            # 缩放后的控件尺寸
            label_width = self.scaled_size(18)
            status_width = self.scaled_size(100)
            button_width = self.scaled_size(12)
            button_width_small = self.scaled_size(9)
            button_width_medium = self.scaled_size(10)

            tk.Label(row_frame, text=name_label_text,
                    font=("Arial", self.scaled_font(12), "bold"),
                    width=label_width, anchor='w',
                    bg='white', fg=text_color).pack(side=tk.LEFT, padx=(0, self.scaled_size(12)))

            # 状态显示 - 缩放版本
            status_var = tk.StringVar(value="未设置区域")
            self.status_labels[action_name] = status_var

            status_label = tk.Label(row_frame, textvariable=status_var,
                                   font=("Arial", self.scaled_font(11)),
                                   fg='#2c3e50', bg='white',
                                   width=status_width, anchor='w')
            status_label.pack(side=tk.LEFT, padx=(0, self.scaled_size(15)))

            # 框选按钮 - 缩放版本
            select_btn = tk.Button(row_frame, text="框选区域",
                                  font=("Arial", self.scaled_font(10), "bold"),
                                  bg=bg_color, fg='white',
                                  width=button_width,
                                  cursor="hand2",
                                  command=lambda idx=i, name=action_name: self.select_area(name))
            select_btn.pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

            # 测试按钮 - 缩放版本
            test_btn = tk.Button(row_frame, text="测试",
                               font=("Arial", self.scaled_font(10), "bold"),
                               bg='#95a5a6', fg='white',
                               width=button_width_small,
                               cursor="hand2",
                               command=lambda idx=i, name=action_name: self.test_click(name))
            test_btn.pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

            # 执行按钮 - 缩放版本
            exec_btn = tk.Button(row_frame, text="执行",
                               font=("Arial", self.scaled_font(10), "bold"),
                               bg='#27ae60', fg='white',
                               width=button_width_small,
                               cursor="hand2",
                               command=lambda idx=i, name=action_name: self.execute_action(name))
            exec_btn.pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

            # 添加获取鼠标位置按钮（仅对动作8）
            if i == 8:
                mouse_btn = tk.Button(row_frame, text="获取鼠标",
                                     font=("Arial", self.scaled_font(10), "bold"),
                                     bg='#e67e22', fg='white',
                                     width=button_width_medium,
                                     cursor="hand2",
                                     command=lambda idx=i, name=action_name: self.get_mouse_position(name))
                mouse_btn.pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

            # 添加选择文档按钮（仅对动作9）
            if i == 9:
                doc_btn = tk.Button(row_frame, text="选择文档",
                                   font=("Arial", self.scaled_font(10), "bold"),
                                   bg='#9b59b6', fg='white',
                                   width=button_width_medium,
                                   cursor="hand2",
                                   command=lambda idx=i, name=action_name: self.select_document(name))
                doc_btn.pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

      
        # 批量操作区域 - 缩放版本
        batch_frame = tk.LabelFrame(main_container, text="批量操作",
                                   font=("Arial", self.scaled_font(14), "bold"),
                                   fg='#2c3e50', bg='white')
        batch_frame.pack(fill=tk.X, pady=(self.scaled_size(3), self.scaled_size(3)))

        batch_row = tk.Frame(batch_frame, bg='white')
        batch_row.pack(pady=self.scaled_size(4))

        # 批量操作按钮尺寸
        batch_button_width = self.scaled_size(20)
        batch_button_width_small = self.scaled_size(17)
        batch_button_width_tiny = self.scaled_size(14)

        tk.Button(batch_row, text="[顺序] 顺序执行全部",
                 font=("Arial", self.scaled_font(13), "bold"),
                 bg='#3498db', fg='white',
                 width=batch_button_width,
                 cursor="hand2",
                 command=self.execute_all).pack(side=tk.LEFT, padx=(self.scaled_size(25), self.scaled_size(15)))

        tk.Button(batch_row, text="[循环] 循环执行",
                 font=("Arial", self.scaled_font(13), "bold"),
                 bg='#9b59b6', fg='white',
                 width=batch_button_width_small,
                 cursor="hand2",
                 command=self.loop_execute).pack(side=tk.LEFT, padx=self.scaled_size(15))

        self.stop_btn = tk.Button(batch_row, text="[停止] 停止执行",
                                 font=("Arial", self.scaled_font(13), "bold"),
                                 bg='#e74c3c', fg='white',
                                 width=batch_button_width_small,
                                 cursor="hand2",
                                 command=self.stop_execution,
                                 state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=self.scaled_size(15))

        # 范围执行区域 - 缩放版本
        range_frame = tk.Frame(batch_frame, bg='white')
        range_frame.pack(pady=(self.scaled_size(3), self.scaled_size(2)))

        # 范围执行标签
        tk.Label(range_frame, text="[范围] 范围执行:",
                font=("Arial", self.scaled_font(12), "bold"),
                fg='#2c3e50', bg='white').pack(side=tk.LEFT, padx=(self.scaled_size(25), self.scaled_size(15)))

        # 起始动作
        tk.Label(range_frame, text="从动作:",
                font=("Arial", self.scaled_font(11), "bold"),
                bg='white').pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

        self.start_action = tk.StringVar(value="8")
        start_spinbox = tk.Spinbox(range_frame, from_=1, to=12, width=self.scaled_size(6),
                                   textvariable=self.start_action,
                                   font=("Arial", self.scaled_font(11)))
        start_spinbox.pack(side=tk.LEFT, padx=(0, self.scaled_size(15)))

        # 结束动作
        tk.Label(range_frame, text="到动作:",
                font=("Arial", self.scaled_font(11), "bold"),
                bg='white').pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

        self.end_action = tk.StringVar(value="12")
        end_spinbox = tk.Spinbox(range_frame, from_=1, to=12, width=self.scaled_size(6),
                                 textvariable=self.end_action,
                                 font=("Arial", self.scaled_font(11)))
        end_spinbox.pack(side=tk.LEFT, padx=(0, self.scaled_size(15)))

        # 执行按钮
        tk.Button(range_frame, text="[执行] 执行范围",
                 font=("Arial", self.scaled_font(12), "bold"),
                 bg='#e67e22', fg='white',
                 width=batch_button_width_tiny,
                 cursor="hand2",
                 command=self.execute_range).pack(side=tk.LEFT, padx=(self.scaled_size(15), 0))

        # 设置区域 - 缩放版本
        settings_frame = tk.LabelFrame(main_container, text="设置",
                                     font=("Arial", self.scaled_font(14), "bold"),
                                     fg='#2c3e50', bg='white')
        settings_frame.pack(fill=tk.X, pady=(self.scaled_size(3), self.scaled_size(5)))

        settings_row = tk.Frame(settings_frame, bg='white')
        settings_row.pack(pady=self.scaled_size(3))

        tk.Label(settings_row, text="执行前延迟(秒):",
                font=("Arial", self.scaled_font(12), "bold"),
                bg='white').pack(side=tk.LEFT, padx=(self.scaled_size(20), self.scaled_size(8)))

        self.pre_delay = tk.DoubleVar(value=1.0)
        tk.Spinbox(settings_row, from_=0, to=5, increment=0.1,
                  textvariable=self.pre_delay, width=self.scaled_size(10),
                  font=("Arial", self.scaled_font(11))).pack(side=tk.LEFT, padx=(0, self.scaled_size(25)))

        tk.Label(settings_row, text="执行后延迟(秒):",
                font=("Arial", self.scaled_font(12), "bold"),
                bg='white').pack(side=tk.LEFT, padx=(0, self.scaled_size(8)))

        self.post_delay = tk.DoubleVar(value=0.5)
        tk.Spinbox(settings_row, from_=0, to=5, increment=0.1,
                  textvariable=self.post_delay, width=self.scaled_size(10),
                  font=("Arial", self.scaled_font(11))).pack(side=tk.LEFT, padx=self.scaled_size(8))

        # 底部控制面板 - 缩放版本，简洁风格
        bottom_container = tk.Frame(main_container, bg='#2c3e50', relief=tk.RIDGE, bd=2)
        bottom_container.pack(fill=tk.X, pady=(self.scaled_size(5), self.scaled_size(8)))

        # 控制面板标题
        title_frame = tk.Frame(bottom_container, bg='#2c3e50')
        title_frame.pack(fill=tk.X, padx=self.scaled_size(20), pady=(self.scaled_size(3), self.scaled_size(2)))

        title_label = tk.Label(title_frame, text="[控制] 控制面板",
                              font=("Arial", self.scaled_font(16), "bold"),
                              fg='white', bg='#2c3e50')
        title_label.pack()

        # 按钮容器 - 优化的大尺寸布局
        buttons_frame = tk.Frame(bottom_container, bg='#2c3e50')
        buttons_frame.pack(fill=tk.X, padx=self.scaled_size(40), pady=self.scaled_size(6))

        # 控制面板按钮尺寸 - 增大以更好利用空间
        control_button_width = self.scaled_size(22)
        control_button_height = self.scaled_size(5)
        control_button_font = self.scaled_font(16)

        # 创建4个控制按钮，充分利用新增空间
        tk.Button(buttons_frame, text="[保存]\n保存配置",
                 font=("Arial", control_button_font, "bold"),
                 bg='#27ae60', fg='white',
                 width=control_button_width, height=control_button_height,
                 cursor="hand2",
                 relief=tk.RAISED, bd=3,
                 command=self.save_config).pack(side=tk.LEFT, padx=(0, self.scaled_size(25)), expand=True, fill=tk.X)

        tk.Button(buttons_frame, text="[加载]\n加载配置",
                 font=("Arial", control_button_font, "bold"),
                 bg='#3498db', fg='white',
                 width=control_button_width, height=control_button_height,
                 cursor="hand2",
                 relief=tk.RAISED, bd=3,
                 command=self.load_config).pack(side=tk.LEFT, padx=self.scaled_size(25), expand=True, fill=tk.X)

        tk.Button(buttons_frame, text="[清除]\n清除所有",
                 font=("Arial", control_button_font, "bold"),
                 bg='#e67e22', fg='white',
                 width=control_button_width, height=control_button_height,
                 cursor="hand2",
                 relief=tk.RAISED, bd=3,
                 command=self.clear_all).pack(side=tk.LEFT, padx=self.scaled_size(25), expand=True, fill=tk.X)

        tk.Button(buttons_frame, text="[帮助]\n使用帮助",
                 font=("Arial", control_button_font, "bold"),
                 bg='#95a5a6', fg='white',
                 width=control_button_width, height=control_button_height,
                 cursor="hand2",
                 relief=tk.RAISED, bd=3,
                 command=self.show_help).pack(side=tk.LEFT, padx=self.scaled_size(25), expand=True, fill=tk.X)

        # 底部信息栏 - 简化版本
        info_frame = tk.Frame(bottom_container, bg='#2c3e50')
        info_frame.pack(fill=tk.X, padx=self.scaled_size(20), pady=(self.scaled_size(3), self.scaled_size(5)))

        info_label = tk.Label(info_frame,
                             text="[范围] 智能自动点击工具 Pro - 75%缩放版 | "
                                  "[桌面] 远程桌面优化 | "
                                  "[执行] 12个智能动作",
                             font=("Arial", self.scaled_font(10)),
                             fg='#bdc3c7', bg='#2c3e50')
        info_label.pack()

        # 状态栏
        self.status_var = tk.StringVar(value="准备就绪")
        status_bar = tk.Frame(self.root, bg='#2c3e50', height=30)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        status_bar.pack_propagate(False)

        tk.Label(status_bar, textvariable=self.status_var,
                font=("Arial", 10),
                fg='white', bg='#2c3e50').pack(expand=True)

    def select_area(self, action_name):
        """选择区域"""
        self.current_action = action_name
        self.status_var.set(f"正在为{action_name}选择区域... ESC取消")

        # 隐藏主窗口
        self.root.withdraw()

        # 创建选择窗口
        self.selection_win = tk.Toplevel()
        self.selection_win.attributes('-fullscreen', True)
        self.selection_win.attributes('-topmost', True)
        self.selection_win.attributes('-alpha', 0.3)
        self.selection_win.configure(bg='black')

        # 画布
        self.canvas = tk.Canvas(self.selection_win, bg='black', cursor="crosshair")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 说明文字
        self.canvas.create_text(
            self.canvas.winfo_screenwidth()//2, 80,
            text=f"为 {action_name} 选择区域",
            fill="white", font=("Arial", 20, "bold"),
            tags="info"
        )

        self.canvas.create_text(
            self.canvas.winfo_screenwidth()//2, 120,
            text="按住鼠标左键拖动选择区域 | ESC键取消",
            fill="white", font=("Arial", 14),
            tags="info2"
        )

        # 绑定事件
        self.canvas.bind('<Button-1>', self.on_start)
        self.canvas.bind('<B1-Motion>', self.on_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_end)
        self.selection_win.bind('<Escape>', self.cancel_select)

        self.start_x = None
        self.start_y = None
        self.rect = None

    def on_start(self, event):
        """开始选择"""
        self.start_x = event.x
        self.start_y = event.y
        self.canvas.delete("info")
        self.canvas.delete("info2")

    def on_drag(self, event):
        """拖动"""
        if self.rect:
            self.canvas.delete(self.rect)

        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, event.x, event.y,
            outline='red', width=3
        )

    def on_end(self, event):
        """结束选择"""
        if self.rect:
            self.canvas.delete(self.rect)

        # 计算区域
        x1 = min(self.start_x, event.x)
        y1 = min(self.start_y, event.y)
        x2 = max(self.start_x, event.x)
        y2 = max(self.start_y, event.y)

        width = x2 - x1
        height = y2 - y1

        if width >= 20 and height >= 20:
            # 保存
            self.actions[self.current_action] = {
                'x': x1, 'y': y1, 'width': width, 'height': height
            }

            center_x = x1 + width // 2
            center_y = y1 + height // 2

            # 更新显示
            self.status_labels[self.current_action].set(
                f"[完成] 区域({x1},{y1}) 中心({center_x},{center_y}) 大小{width}x{height}"
            )

            self.status_var.set(f"{self.current_action}设置成功")
        else:
            messagebox.showwarning("警告", "选择区域太小，请重新选择")
            self.status_var.set(f"{self.current_action}设置失败")

        # 关闭选择窗口
        self.selection_win.destroy()
        self.root.deiconify()

    def cancel_select(self, event):
        """取消选择"""
        if self.selection_win:
            self.selection_win.destroy()
        self.root.deiconify()
        self.status_var.set("选择已取消")

    def test_click(self, action_name):
        """测试点击"""
        # 动作12（键盘粘贴）不需要设置区域
        if action_name != "动作12" and action_name not in self.actions:
            messagebox.showwarning("警告", f"请先设置{action_name}")
            return

        # 动作11不需要区域信息，直接跳过
        if action_name != "动作11":
            action = self.actions[action_name]
            center_x = action['x'] + action['width'] // 2
            center_y = action['y'] + action['height'] // 2

        if action_name == "动作8":
            self.status_var.set(f"测试双击{action_name}: ({center_x},{center_y})")
            # 移动并双击
            pyautogui.moveTo(center_x, center_y, duration=0.5)
            time.sleep(0.3)
            pyautogui.doubleClick()
            self.status_var.set(f"{action_name}双击测试完成")
        elif action_name == "动作9":
            self.status_var.set(f"测试{action_name}: 智能文档识别复制")
            # 测试文档识别复制功能
            success = self.test_document_copy(action_name)
            if success:
                self.status_var.set(f"{action_name}测试完成: 文档识别复制成功")
            else:
                self.status_var.set(f"{action_name}测试完成: 文档识别复制失败")
        elif action_name == "动作11":
            self.status_var.set(f"测试{action_name}: 自动右击")
            # 测试右击功能（增强版）
            action = self.actions[action_name]
            center_x = action['x'] + action['width'] // 2
            center_y = action['y'] + action['height'] // 2

            # 使用增强的右键点击方法
            success = EnhancedRightClickMethods.execute_enhanced_right_click(
                center_x, center_y,
                status_callback=self.status_var.set
            )

            if success:
                self.status_var.set(f"{action_name}测试完成: 右击测试成功")
            else:
                # 追加诊断信息，便于定位是哪一步失败
                diagnose_results = EnhancedRightClickMethods.diagnose_right_click_methods(center_x, center_y)

                detail_lines = []
                for item in diagnose_results:
                    status_flag = "✅" if item["success"] else "❌"
                    msg = f"{status_flag} 方法{item['index']}: {item['name']}"
                    if item.get("error"):
                        msg += f" (异常: {item['error']})"
                    detail_lines.append(msg)

                detail_message = "\n".join(detail_lines) if detail_lines else "未获取到诊断信息"

                self.status_var.set(f"{action_name}测试完成: 右击测试已执行（请手动验证）")
                messagebox.showinfo("右键诊断结果", f"右键诊断详情:\n{detail_message}")
        elif action_name == "动作12":
            self.status_var.set(f"测试{action_name}: 键盘粘贴")
            # 测试键盘粘贴功能
            success = self.test_keyboard_paste(action_name)
            if success:
                self.status_var.set(f"{action_name}测试完成: 键盘粘贴成功")
            else:
                self.status_var.set(f"{action_name}测试完成: 键盘粘贴失败")
        else:
            self.status_var.set(f"测试点击{action_name}: ({center_x},{center_y})")
            # 移动并单击
            pyautogui.moveTo(center_x, center_y, duration=0.5)
            time.sleep(0.3)
            pyautogui.click()
            self.status_var.set(f"{action_name}测试完成")

    def execute_action(self, action_name):
        """执行动作"""

        # 动作12（键盘粘贴）不需要设置区域
        if action_name != "动作12" and action_name not in self.actions:
            messagebox.showwarning("警告", f"请先设置{action_name}")
            return

        # 线程执行
        if action_name == "动作8":
            thread = threading.Thread(target=self._execute_double_click, args=(action_name,))
        elif action_name == "动作9":
            thread = threading.Thread(target=self._execute_document_copy, args=(action_name,))
        elif action_name == "动作11":
            thread = threading.Thread(target=self._execute_double_click, args=(action_name,))
        elif action_name == "动作12":
            thread = threading.Thread(target=self._execute_keyboard_paste, args=(action_name,))
        else:
            thread = threading.Thread(target=self._execute_single, args=(action_name,))
        thread.daemon = True
        thread.start()

    def _execute_single(self, action_name):
        """执行单个动作"""
        try:
            action = self.actions[action_name]
            center_x = action['x'] + action['width'] // 2
            center_y = action['y'] + action['height'] // 2

            self.status_var.set(f"执行{action_name}...")

            # 前延迟
            delay = self.pre_delay.get()
            if delay > 0:
                time.sleep(delay)

            # 点击
            pyautogui.moveTo(center_x, center_y, duration=0.3)
            pyautogui.click()

            # 后延迟
            after_delay = self.post_delay.get()
            if after_delay > 0:
                time.sleep(after_delay)

            self.status_var.set(f"{action_name}执行完成")

        except Exception as e:
            self.status_var.set(f"执行失败: {e}")
            messagebox.showerror("错误", f"执行失败: {e}")

    def _execute_double_click(self, action_name):
        """执行右击动作（增强版 - 多方法兼容）"""
        try:
            action = self.actions[action_name]
            center_x = action['x'] + action['width'] // 2
            center_y = action['y'] + action['height'] // 2

            self.status_var.set(f"右击执行{action_name}...")

            # 前延迟
            delay = self.pre_delay.get()
            if delay > 0:
                time.sleep(delay)

            # 使用增强的右键点击方法
            success = EnhancedRightClickMethods.execute_enhanced_right_click(
                center_x, center_y,
                status_callback=self.status_var.set
            )

            # 后延迟
            after_delay = self.post_delay.get()
            if after_delay > 0:
                time.sleep(after_delay)

            if success:
                self.status_var.set(f"{action_name}右击执行完成")
            else:
                self.status_var.set(f"{action_name}右击执行完成（请手动验证）")

        except Exception as e:
            self.status_var.set(f"右击执行失败: {e}")
            print(f"右击执行失败: {e}")
            # 不再弹出阻塞对话框，避免打断自动化流程

    def _execute_copy_paste(self, action_name):
        """执行智能复制粘贴动作"""
        try:
            action = self.actions[action_name]
            center_x = action['x'] + action['width'] // 2
            center_y = action['y'] + action['height'] // 2

            self.status_var.set(f"执行{action_name}: 智能复制粘贴...")

            # 前延迟
            delay = self.pre_delay.get()
            if delay > 0:
                time.sleep(delay)

            # 移动到目标位置
            pyautogui.moveTo(center_x, center_y, duration=0.3)
            time.sleep(0.3)

            # 先点击确保获得焦点
            self.status_var.set(f"执行{action_name}: 点击获得焦点...")
            pyautogui.click()
            time.sleep(0.5)

            # 全选
            self.status_var.set(f"执行{action_name}: 执行全选...")
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)

            # 复制
            self.status_var.set(f"执行{action_name}: 执行复制...")
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)

            # 验证是否复制成功
            try:
                clipboard_content = self.root.clipboard_get()
                if len(clipboard_content.strip()) > 0:
                    self.status_var.set(f"{action_name}执行完成: 复制成功({len(clipboard_content)}字符)")
                else:
                    self.status_var.set(f"{action_name}执行完成: 复制完成但内容为空")
            except Exception as clipboard_error:
                self.status_var.set(f"{action_name}执行完成: 复制完成(无法验证剪贴板)")
                print(f"剪贴板验证失败: {clipboard_error}")

            # 后延迟
            after_delay = self.post_delay.get()
            if after_delay > 0:
                time.sleep(after_delay)

        except Exception as e:
            self.status_var.set(f"{action_name}执行失败: {e}")
            messagebox.showerror("错误", f"{action_name}执行失败: {e}")

    
    def _execute_click_copy_action(self, action_name):
        """执行点击+复制动作"""
        try:
            action = self.actions[action_name]
            center_x = action['x'] + action['width'] // 2
            center_y = action['y'] + action['height'] // 2

            self.status_var.set(f"执行{action_name}: 双击+全选+复制...")

            # 前延迟
            delay = self.pre_delay.get()
            if delay > 0:
                time.sleep(delay)

            # 移动到目标位置
            pyautogui.moveTo(center_x, center_y, duration=0.3)
            time.sleep(0.3)

            # 双击
            self.status_var.set(f"执行{action_name}: 双击...")
            pyautogui.doubleClick()
            time.sleep(0.5)

            # 全选
            self.status_var.set(f"执行{action_name}: 全选...")
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)

            # 复制
            self.status_var.set(f"执行{action_name}: 复制...")
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)

            # 后延迟
            after_delay = self.post_delay.get()
            if after_delay > 0:
                time.sleep(after_delay)

            self.status_var.set(f"{action_name}执行完成: 已双击+全选+复制")

        except Exception as e:
            self.status_var.set(f"{action_name}执行失败: {e}")
            messagebox.showerror("错误", f"{action_name}执行失败: {e}")

    def execute_all(self):
        """执行所有"""
        if not self.actions:
            messagebox.showwarning("警告", "请先设置动作")
            return

        thread = threading.Thread(target=self._execute_all, args=(1,))
        thread.daemon = True
        thread.start()

    def execute_range(self):
        """执行指定范围的动作"""
        try:
            # 获取起始和结束动作
            start_num = int(self.start_action.get())
            end_num = int(self.end_action.get())

            # 验证范围
            if start_num < 1 or start_num > 12 or end_num < 1 or end_num > 12:
                messagebox.showwarning("警告", "动作编号必须在1-12之间")
                return

            if start_num > end_num:
                messagebox.showwarning("警告", "起始动作不能大于结束动作")
                return

            # 生成动作列表
            action_list = []
            for i in range(start_num, end_num + 1):
                action_name = f"动作{i}"
                action_list.append(action_name)

            # 检查是否有设置的动作（动作12不需要设置区域）
            has_actions = False
            for action_name in action_list:
                if action_name == "动作12" or action_name in self.actions:
                    has_actions = True
                    break

            if not has_actions:
                messagebox.showwarning("警告", f"请先设置动作{start_num}到{end_num}中的至少一个动作")
                return

            # 确认执行
            result = messagebox.askyesno("确认执行",
                f"即将执行动作{start_num}到{end_num}，共{len(action_list)}个动作\n\n确认执行吗？")
            if not result:
                return

            # 启动线程执行范围
            thread = threading.Thread(target=self._execute_range, args=(action_list,))
            thread.daemon = True
            thread.start()

        except ValueError:
            messagebox.showwarning("警告", "请输入有效的动作编号")
        except Exception as e:
            messagebox.showerror("错误", f"范围执行失败: {e}")

    def loop_execute(self):
        """循环执行"""
        if not self.actions:
            messagebox.showwarning("警告", "请先设置动作")
            return

        thread = threading.Thread(target=self._execute_all, args=(0,))
        thread.daemon = True
        thread.start()

    def _execute_range(self, action_list):
        """执行指定范围的动作"""
        self.is_running = True
        self.stop_btn.config(state=tk.NORMAL)

        try:
            self.status_var.set(f"[范围] 开始执行范围动作: {', '.join(action_list)}")

            # 逐个执行动作
            for action_name in action_list:
                if not self.is_running:
                    break

                # 检查动作是否已设置（动作12不需要设置区域）
                if action_name == "动作12" or action_name in self.actions:
                    self.status_var.set(f"[执行] 执行{action_name}...")

                    # 根据动作类型调用相应的执行方法
                    if action_name == "动作8":
                        # 动作8：打开TXT文档
                        self._execute_special_action("动作8", "open_txt")
                    elif action_name == "动作9":
                        # 动作9：智能文档识别复制
                        self._execute_document_copy("动作9")
                    elif action_name == "动作11":
                        # 动作11：自动右击
                        self._execute_double_click("动作11")
                    elif action_name == "动作12":
                        # 动作12：键盘粘贴
                        self._execute_keyboard_paste("动作12")
                    else:
                        # 其他动作：普通点击
                        self._execute_single(action_name)

                    # 动作间延迟
                    if self.is_running:
                        time.sleep(self.pre_delay.get())
                else:
                    self.status_var.set(f"[跳过] 跳过{action_name}(未设置)")
                    time.sleep(0.2)  # 短暂延迟

            if self.is_running:
                self.status_var.set("[成功] 范围执行完成！")

        except Exception as e:
            self.status_var.set(f"[错误] 范围执行失败: {e}")
            messagebox.showerror("错误", f"范围执行失败: {e}")
        finally:
            self.is_running = False
            self.stop_btn.config(state=tk.DISABLED)

    def _execute_all(self, loops):
        """执行所有动作"""
        self.is_running = True
        self.stop_btn.config(state=tk.NORMAL)
        count = 0

        try:
            while self.is_running:
                count += 1
                if loops > 0 and count > loops:
                    break

                self.status_var.set(f"[顺序] 第{count}轮执行开始...")

                # 执行动作1-3
                for action_name in ["动作1", "动作2", "动作3"]:
                    if not self.is_running:
                        break

                    if action_name in self.actions:
                        self.status_var.set(f"[执行] 执行{action_name}...")
                        self._execute_single(action_name)

                # 执行动作4
                if self.is_running and "动作4" in self.actions:
                    self.status_var.set("[执行] 执行动作4...")
                    self._execute_single("动作4")

                # 如果设置了动作5，在动作4完成后等待5秒再执行动作5
                if self.is_running and "动作5" in self.actions:
                    self.status_var.set("[延迟] 等待5秒后执行动作5...")
                    for i in range(5):
                        if not self.is_running:
                            break
                        time.sleep(1)
                        self.status_var.set(f"[延迟] 等待5秒后执行动作5... ({5-i}秒)")

                    if self.is_running:
                        self.status_var.set("[执行] 执行动作5...")
                        self._execute_single("动作5")

                # 执行动作6 (动作5后延迟5秒执行)
                if self.is_running and "动作6" in self.actions:
                    self.status_var.set("[延迟] 等待5秒后执行动作6...")
                    for i in range(5):
                        if not self.is_running:
                            break
                        time.sleep(1)
                        self.status_var.set(f"[延迟] 等待5秒后执行动作6... ({5-i}秒)")

                    if self.is_running:
                        self.status_var.set("[执行] 执行动作6...")
                        self._execute_single("动作6")

                # 执行动作7 (最后动作)
                if self.is_running and "动作7" in self.actions:
                    self.status_var.set("[执行] 执行动作7...")
                    self._execute_single("动作7")

                # 执行动作8 (打开TXT文档)
                if self.is_running and "动作8" in self.actions:
                    self.status_var.set("[文档] 执行动作8 - 打开TXT文档...")
                    self._execute_special_action("动作8", "open_txt")

                # 执行动作9 (智能文档识别复制)
                if self.is_running and "动作9" in self.actions:
                    self.status_var.set("[复制] 执行动作9 - 智能文档识别复制...")
                    self._execute_document_copy("动作9")

                # 执行动作10 (添加URL)
                if self.is_running and "动作10" in self.actions:
                    self.status_var.set("[链接] 执行动作10 - 添加URL...")
                    self._execute_single("动作10")

                # 执行动作11 (自动右击)
                if self.is_running and "动作11" in self.actions:
                    self.status_var.set("[新增] 执行动作11 - 自动右击...")
                    self._execute_double_click("动作11")

                # 执行动作12 (键盘粘贴)
                if self.is_running:
                    self.status_var.set("[新增] 执行动作12 - 键盘粘贴...")
                    self._execute_keyboard_paste("动作12")

                # 循环间延迟
                if loops == 0 and self.is_running:
                    self.status_var.set("[延迟] 循环间延迟...")
                    time.sleep(self.pre_delay.get())

            self.status_var.set("[成功] 执行完成！")

        except Exception as e:
            self.status_var.set(f"[错误] 执行失败: {e}")
            messagebox.showerror("错误", f"执行失败: {e}")
        finally:
            self.is_running = False
            self.stop_btn.config(state=tk.DISABLED)

    def _execute_special_action(self, action_name, action_type):
        """执行特殊动作"""
        try:
            if action_type == "open_txt":
                # 打开TXT文档
                file_path = filedialog.askopenfilename(
                    title="选择TXT文档",
                    filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
                )

                if file_path:
                    # 保存选择的文件路径
                    self.txt_file_path = file_path
                    print(f"动作8保存文件路径: {file_path}")  # 调试信息

                    # 用记事本打开文件
                    try:
                        if os.name == 'nt':  # Windows
                            os.startfile(file_path)
                        else:  # macOS/Linux
                            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
                            subprocess.call([opener, file_path])

                        self.status_var.set(f"[文档] 已打开: {os.path.basename(file_path)}")
                    except Exception as e:
                        self.status_var.set(f"[错误] 打开文件失败: {e}")
                else:
                    self.status_var.set("[文档] 未选择文件")

            elif action_type == "copy_txt":
                # 智能复制：优先从动作8打开的文件读取内容，其次使用快捷键
                print("动作9开始执行复制功能...")  # 调试信息

                try:
                    # 检查是否有动作8打开的文件路径
                    if hasattr(self, 'txt_file_path') and os.path.exists(self.txt_file_path):
                        print(f"动作9找到文件路径: {self.txt_file_path}")  # 调试信息
                        self.status_var.set("[复制] 从文件读取内容...")

                        try:
                            with open(self.txt_file_path, 'r', encoding='utf-8') as f:
                                content = f.read().strip()

                            print(f"文件内容长度: {len(content)} 字符")  # 调试信息
                            print(f"文件内容前100字符: {content[:100]}")  # 调试信息

                            if content:
                                # 方法1：直接使用Windows剪贴板API
                                success = False
                                try:
                                    import ctypes

                                    # Windows剪贴板API
                                    user32 = ctypes.windll.user32

                                    # 打开剪贴板
                                    if user32.OpenClipboard(0):
                                        try:
                                            # 清空剪贴板
                                            user32.EmptyClipboard()

                                            # 创建Unicode文本
                                            data = content.encode('utf-16le') + b'\x00\x00'
                                            user32.SetClipboardData(13, data)  # CF_UNICODETEXT = 13

                                            success = True
                                            self.status_var.set(f"[复制] 已复制内容 ({len(content)}字符)[Windows API]")
                                            print("使用Windows API复制成功")
                                        finally:
                                            user32.CloseClipboard()
                                    else:
                                        print("无法打开剪贴板")

                                except Exception as api_error:
                                    print(f"Windows API剪贴板失败: {api_error}")

                                # 方法2：使用tkinter剪贴板
                                if not success:
                                    try:
                                        self.root.clipboard_clear()
                                        self.root.clipboard_append(content)
                                        self.root.update()
                                        success = True
                                        self.status_var.set(f"[复制] 已复制内容 ({len(content)}字符)[tkinter]")
                                        print("使用tkinter复制成功")
                                    except Exception as e1:
                                        print(f"tkinter剪贴板失败: {e1}")

                                # 方法3：使用pyperclip（如果有安装）
                                if not success:
                                    try:
                                        import pyperclip
                                        pyperclip.copy(content)
                                        success = True
                                        self.status_var.set(f"[复制] 已复制内容 ({len(content)}字符)[pyperclip]")
                                        print("使用pyperclip复制成功")
                                    except ImportError:
                                        print("pyperclip未安装，跳过")
                                    except Exception as e2:
                                        print(f"pyperclip剪贴板失败: {e2}")

                                # 方法4：使用快捷键方式（针对远程桌面优化）
                                if not success:
                                    print("所有直接复制方法都失败，尝试远程桌面优化方式...")
                                    self.status_var.set("[复制] 使用远程桌面优化复制...")

                                    # 重新激活记事本窗口并确保获得焦点
                                    if os.name == 'nt':
                                        os.startfile(self.txt_file_path)
                                    time.sleep(3)  # 增加等待时间，确保窗口完全激活

                                    # 多次尝试确保窗口获得焦点（远程桌面需要）
                                    for attempt in range(3):
                                        try:
                                            # 点击窗口标题栏确保获得焦点（远程桌面需要）
                                            pyautogui.click()
                                            time.sleep(0.5)

                                            # 使用更可靠的快捷键方式
                                            pyautogui.hotkey('ctrl', 'a')
                                            time.sleep(1)  # 增加等待时间

                                            pyautogui.hotkey('ctrl', 'c')
                                            time.sleep(1)  # 增加等待时间

                                            # 验证是否复制成功
                                            try:
                                                clipboard_test = self.root.clipboard_get()
                                                if content[:50] in clipboard_test:
                                                    success = True
                                                    self.status_var.set(f"[复制] 远程桌面复制成功(尝试{attempt+1})")
                                                    print(f"远程桌面复制成功(第{attempt+1}次尝试)")
                                                    break
                                            except:
                                                pass

                                            if attempt < 2:
                                                print(f"第{attempt+1}次复制尝试失败，重试中...")

                                        except Exception as retry_error:
                                            print(f"第{attempt+1}次复制尝试出错: {retry_error}")
                                            if attempt < 2:
                                                time.sleep(1)  # 等待后重试

                                    if not success:
                                        self.status_var.set("[错误] 远程桌面复制失败，请手动复制")
                                        print("远程环境下复制失败，建议手动操作")
                                    else:
                                        self.status_var.set("[复制] 使用快捷键复制完成")
                                        print("使用快捷键复制完成")

                            else:
                                self.status_var.set("[复制] 文件为空")
                                print("文件内容为空")

                        except Exception as file_error:
                            print(f"读取文件失败: {file_error}")
                            self.status_var.set(f"[复制] 读取文件失败: {file_error}")

                    else:
                        print("没有找到文件路径，使用远程桌面优化方式")  # 调试信息
                        # 没有文件路径，使用通用快捷键方式（针对远程桌面优化）
                        self.status_var.set("[复制] 远程桌面执行全选+复制...")

                        # 先确保有窗口获得焦点（远程桌面需要更多时间）
                        pyautogui.press('ctrl')  # 激活当前窗口
                        time.sleep(2)  # 增加等待时间

                        # 多次尝试执行复制操作（远程桌面需要）
                        copy_success = False
                        for attempt in range(3):
                            try:
                                # 全选 (Ctrl+A)
                                pyautogui.hotkey('ctrl', 'a')
                                time.sleep(1)  # 增加等待时间

                                # 复制 (Ctrl+C)
                                pyautogui.hotkey('ctrl', 'c')
                                time.sleep(1)  # 增加等待时间

                                # 验证是否复制成功
                                try:
                                    clipboard_test = self.root.clipboard_get()
                                    if clipboard_test and len(clipboard_test) > 0:
                                        copy_success = True
                                        self.status_var.set(f"[复制] 远程桌面全选+复制成功(尝试{attempt+1})")
                                        print(f"远程桌面复制成功(第{attempt+1}次尝试)")
                                        break
                                except:
                                    pass

                                if attempt < 2:
                                    print(f"第{attempt+1}次远程复制尝试失败，重试中...")

                            except Exception as retry_error:
                                print(f"第{attempt+1}次远程复制尝试出错: {retry_error}")
                                if attempt < 2:
                                    time.sleep(1)  # 等待后重试

                        if not copy_success:
                            self.status_var.set("[错误] 远程桌面复制失败，请手动操作")
                            print("远程环境下复制失败，建议手动操作")
                        else:
                            self.status_var.set("[复制] 全选+复制完成")
                            print("通用快捷键复制完成")

                except Exception as e:
                    print(f"动作9复制失败详细错误: {e}")
                    import traceback
                    traceback.print_exc()
                    self.status_var.set(f"[错误] 复制失败: {e}")

                    # 显示错误详情对话框
                    messagebox.showerror("动作9复制失败",
                                       f"复制过程中发生错误：\n\n{e}\n\n请检查：\n1. 是否已正确选择TXT文件\n2. 文件是否包含内容\n3. 系统剪贴板是否可用")

        except Exception as e:
            self.status_var.set(f"[错误] {action_name}执行失败: {e}")
            messagebox.showerror(f"{action_name}执行失败", f"执行{action_name}时发生错误：\n\n{e}")

    def get_mouse_position(self, action_name):
        """获取当前鼠标位置"""
        try:
            # 获取当前鼠标位置
            current_x, current_y = pyautogui.position()

            # 创建一个以鼠标位置为中心的小区域（20x20像素）
            size = 20
            x = current_x - size // 2
            y = current_y - size // 2

            # 保存位置信息
            self.actions[action_name] = {
                'x': x, 'y': y, 'width': size, 'height': size
            }

            # 更新显示
            self.status_labels[action_name].set(
                f"[完成] 鼠标位置({current_x},{current_y}) 区域({x},{y}) 大小{size}x{size}"
            )

            self.status_var.set(f"{action_name}已保存鼠标位置: ({current_x},{current_y})")

        except Exception as e:
            self.status_var.set(f"获取鼠标位置失败: {e}")

    def select_document(self, action_name):
        """选择要识别复制的文档"""
        try:
            # 支持多种文档格式
            file_path = filedialog.askopenfilename(
                title="选择要识别复制的文档",
                filetypes=[
                    ("文本文件", "*.txt"),
                    ("Word文档", "*.docx;*.doc"),
                    ("PDF文档", "*.pdf"),
                    ("Excel文件", "*.xlsx;*.xls"),
                    ("所有文件", "*.*")
                ]
            )

            if file_path:
                # 保存文档路径
                if not hasattr(self, 'document_paths'):
                    self.document_paths = {}
                self.document_paths[action_name] = file_path

                # 更新状态显示
                file_name = os.path.basename(file_path)
                file_ext = os.path.splitext(file_name)[1].lower()

                if hasattr(self, 'actions') and action_name in self.actions:
                    # 更新动作信息中包含文档路径
                    self.actions[action_name]['document_path'] = file_path

                self.status_labels[action_name].set(
                    f"[完成] 已选择文档: {file_name}"
                )
                self.status_var.set(f"{action_name}已选择文档: {file_name}")

                # 尝试读取文档内容进行预览
                try:
                    content = self.read_document_content(file_path)
                    if content:
                        preview = content[:100] + "..." if len(content) > 100 else content
                        print(f"文档预览: {preview}")
                        self.status_var.set(f"{action_name}文档加载成功，内容长度: {len(content)}字符")
                    else:
                        self.status_var.set(f"{action_name}文档内容为空")
                except Exception as read_error:
                    self.status_var.set(f"{action_name}文档读取失败: {read_error}")
            else:
                self.status_var.set(f"{action_name}未选择文档")

        except Exception as e:
            self.status_var.set(f"{action_name}选择文档失败: {e}")
            messagebox.showerror("错误", f"选择文档失败: {e}")

    def read_document_content(self, file_path):
        """读取文档内容"""
        try:
            file_ext = os.path.splitext(file_path)[1].lower()

            if file_ext == '.txt':
                # 读取文本文件
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read()

            elif file_ext in ['.docx', '.doc']:
                # 读取Word文档
                try:
                    import docx
                    doc = docx.Document(file_path)
                    return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
                except ImportError:
                    print("需要安装python-docx库来读取Word文档")
                    return None
                except Exception as e:
                    print(f"读取Word文档失败: {e}")
                    return None

            elif file_ext == '.pdf':
                # 读取PDF文档
                try:
                    import PyPDF2
                    with open(file_path, 'rb') as file:
                        reader = PyPDF2.PdfReader(file)
                        text = ""
                        for page in reader.pages:
                            text += page.extract_text()
                        return text
                except ImportError:
                    print("需要安装PyPDF2库来读取PDF文档")
                    return None
                except Exception as e:
                    print(f"读取PDF文档失败: {e}")
                    return None

            elif file_ext in ['.xlsx', '.xls']:
                # 读取Excel文件
                try:
                    import pandas as pd
                    df = pd.read_excel(file_path)
                    return df.to_string()
                except ImportError:
                    print("需要安装pandas和openpyxl库来读取Excel文件")
                    return None
                except Exception as e:
                    print(f"读取Excel文件失败: {e}")
                    return None

            else:
                # 尝试作为文本文件读取
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        return f.read()
                except:
                    try:
                        with open(file_path, 'r', encoding='gbk') as f:
                            return f.read()
                    except Exception as e:
                        print(f"无法读取文件格式: {file_ext}, 错误: {e}")
                        return None

        except Exception as e:
            print(f"读取文档失败: {e}")
            return None

    def test_document_copy(self, action_name):
        """测试文档识别复制功能"""
        try:
            # 检查是否已选择文档
            if not hasattr(self, 'document_paths') or action_name not in self.document_paths:
                messagebox.showwarning("警告", f"请先为{action_name}选择文档")
                return False

            document_path = self.document_paths[action_name]
            self.status_var.set(f"测试{action_name}: 正在读取文档...")

            # 读取文档内容
            content = self.read_document_content(document_path)
            if not content:
                messagebox.showwarning("警告", "无法读取文档内容或文档为空")
                return False

            self.status_var.set(f"测试{action_name}: 文档读取成功，正在复制到剪贴板...")

            # 复制到剪贴板
            success = self.copy_to_clipboard(content)
            if success:
                self.status_var.set(f"测试{action_name}: 成功复制{len(content)}字符到剪贴板")
                return True
            else:
                messagebox.showerror("错误", "复制到剪贴板失败")
                return False

        except Exception as e:
            messagebox.showerror("错误", f"测试文档复制失败: {e}")
            return False

    def test_keyboard_paste(self, action_name):
        """测试键盘粘贴功能（远程桌面兼容版 - 不检查剪贴板）"""
        try:
            # 不再检查剪贴板内容，直接进行测试
            self.status_var.set(f"测试{action_name}: 准备执行粘贴测试...")

            # 提示用户
            result = messagebox.askyesno("粘贴测试",
                f"即将在3秒后执行Ctrl+V粘贴测试\n\n"
                f"请确保：\n"
                f"1. 光标在可粘贴位置\n"
                f"2. 已在目标程序中复制了内容\n\n"
                f"是否继续测试？")
            if not result:
                return False

            # 等待用户准备
            self.status_var.set(f"测试{action_name}: 3秒后开始粘贴测试...")
            for i in range(3, 0, -1):
                self.status_var.set(f"测试{action_name}: {i}秒后开始粘贴测试...")
                time.sleep(1)

            import pyautogui

            # 执行粘贴测试
            self.status_var.set(f"测试{action_name}: 正在执行键盘粘贴...")

            # 先点击确保获得焦点
            pyautogui.click()
            time.sleep(0.5)

            # 执行Ctrl+V
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1.0)

            self.status_var.set(f"测试{action_name}: 键盘粘贴测试完成")

            messagebox.showinfo("测试完成",
                f"已执行Ctrl+V粘贴操作\n\n"
                f"请检查目标位置是否出现了粘贴的内容\n\n"
                f"如果没有内容，请确保：\n"
                f"• 已在目标程序中复制了内容\n"
                f"• 光标位置正确\n"
                f"• 目标程序支持粘贴操作")

            return True

        except Exception as e:
            # 不再弹出阻塞对话框，只打印错误
            print(f"测试键盘粘贴失败: {e}")
            self.status_var.set(f"测试{action_name}: 粘贴测试已执行（请手动验证结果）")
            return False

    def _execute_document_copy(self, action_name):
        """执行文档识别复制动作"""
        try:
            self.status_var.set(f"执行{action_name}: 智能文档识别复制...")

            # 前延迟
            delay = self.pre_delay.get()
            if delay > 0:
                time.sleep(delay)

            # 检查是否已选择文档
            if not hasattr(self, 'document_paths') or action_name not in self.document_paths:
                messagebox.showwarning("警告", f"请先为{action_name}选择要复制的文档")
                return

            document_path = self.document_paths[action_name]
            file_name = os.path.basename(document_path)
            self.status_var.set(f"执行{action_name}: 正在读取文档 {file_name}...")

            # 读取文档内容
            content = self.read_document_content(document_path)
            if not content:
                messagebox.showerror("错误", f"无法读取文档内容或文档为空: {file_name}")
                return

            self.status_var.set(f"执行{action_name}: 文档读取成功({len(content)}字符)，正在复制...")

            # 复制到剪贴板
            success = self.copy_to_clipboard(content)
            if success:
                self.status_var.set(f"{action_name}执行完成: 已复制{len(content)}字符到剪贴板")
                print(f"成功复制文档 '{file_name}' 的内容到剪贴板")
            else:
                messagebox.showerror("错误", "复制到剪贴板失败")
                return

            # 后延迟
            after_delay = self.post_delay.get()
            if after_delay > 0:
                time.sleep(after_delay)

        except Exception as e:
            self.status_var.set(f"{action_name}执行失败: {e}")
            messagebox.showerror("错误", f"{action_name}执行失败: {e}")

    def test_paste(self, action_name):
        """测试粘贴功能"""
        try:
            self.status_var.set(f"测试{action_name}: 检查剪贴板内容...")

            # 检查剪贴板是否有内容
            try:
                clipboard_content = self.root.clipboard_get()
                if not clipboard_content or not clipboard_content.strip():
                    messagebox.showwarning("警告", "剪贴板为空或无法访问，请先复制一些内容")
                    return False
            except:
                try:
                    import pyperclip
                    clipboard_content = pyperclip.paste()
                    if not clipboard_content or not clipboard_content.strip():
                        messagebox.showwarning("警告", "剪贴板为空，请先复制一些内容")
                        return False
                except ImportError:
                    messagebox.showwarning("警告", "无法访问剪贴板，请先复制一些内容")
                    return False

            self.status_var.set(f"测试{action_name}: 剪贴板有内容({len(clipboard_content)}字符)，执行粘贴测试...")

            # 测试粘贴（先点击当前焦点位置）
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)

            self.status_var.set(f"{action_name}测试完成: 粘贴测试完成")
            return True

        except Exception as e:
            messagebox.showerror("错误", f"测试粘贴失败: {e}")
            return False

    def _execute_paste(self, action_name):
        """执行粘贴动作"""
        try:
            self.status_var.set(f"执行{action_name}: 智能粘贴...")

            # 前延迟
            delay = self.pre_delay.get()
            if delay > 0:
                time.sleep(delay)

            # 检查剪贴板内容
            clipboard_content = ""
            try:
                clipboard_content = self.root.clipboard_get()
            except:
                try:
                    import pyperclip
                    clipboard_content = pyperclip.paste()
                except ImportError:
                    pass

            if not clipboard_content or not clipboard_content.strip():
                messagebox.showwarning("警告", "剪贴板为空，无法执行粘贴操作")
                return

            self.status_var.set(f"执行{action_name}: 剪贴板有{len(clipboard_content)}字符，执行粘贴...")

            # 点击确保获得焦点
            pyautogui.click()
            time.sleep(0.3)

            # 执行粘贴
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)

            self.status_var.set(f"{action_name}执行完成: 已粘贴{len(clipboard_content)}字符")
            print(f"成功粘贴 {len(clipboard_content)} 字符")

            # 后延迟
            after_delay = self.post_delay.get()
            if after_delay > 0:
                time.sleep(after_delay)

        except Exception as e:
            self.status_var.set(f"{action_name}执行失败: {e}")
            messagebox.showerror("错误", f"{action_name}执行失败: {e}")

    def _execute_keyboard_paste(self, action_name):
        """执行键盘粘贴动作（增强版 - 多方法兼容）"""
        try:
            self.status_var.set(f"执行{action_name}: 键盘粘贴...")

            # 前延迟
            delay = self.pre_delay.get()
            if delay > 0:
                time.sleep(delay)

            # 使用增强的粘贴方法
            success = EnhancedPasteMethods.execute_enhanced_paste(
                action_name=action_name,
                status_callback=self.status_var.set,
                pre_delay=0,  # 前延迟已在上面处理
                post_delay=0  # 后延迟将在下面处理
            )

            # 后延迟
            after_delay = self.post_delay.get()
            if after_delay > 0:
                time.sleep(after_delay)

        except Exception as e:
            error_msg = f"[错误] {action_name}执行异常: {e}"
            self.status_var.set(error_msg)
            print(f"执行异常: {e}")
            # 不再弹出阻塞对话框，避免打断自动化流程
            import traceback
            traceback.print_exc()

    def _type_special_char(self, char):
        """输入特殊字符"""
        # 特殊字符映射表
        special_char_map = {
            '!': '1', '@': '2', '#': '3', '$': '4', '%': '5', '^': '6', '&': '7', '*': '8',
            '(': '9', ')': '0', '_': '-', '+': '=', '[': '[', ']': ']',
            '{': '[', '}': ']', '|': '\\', ';': ';', ':': ':',
            ',': ',', '.': '.', '<': ',', '>': '.', '?': '/', '`': '`', '~': '`',
            '"': "'", "'": "'"
        }

        if char in special_char_map:
            base_key = special_char_map[char]
            pyautogui.hotkey('shift', base_key)
        else:
            # 回退到ASCII码输入
            self._type_ascii_char(char)

    def _type_printable_char(self, char):
        """输入可打印字符"""
        try:
            # 尝试直接输入
            pyautogui.press(char)
        except:
            # 回退到ASCII码方法
            self._type_ascii_char(char)

    def _type_ascii_char(self, char):
        """使用ASCII码输入字符（终极回退方案）"""
        try:
            import ctypes
            import ctypes.wintypes

            # Windows API SendInput方法
            user32 = ctypes.windll.user32

            # 获取字符的虚拟键码
            vk_code = ctypes.windll.user32.VkKeyScanW(ord(char)) & 0xFF

            # 构造INPUT结构
            class KEYBDINPUT(ctypes.Structure):
                _fields_ = [
                    ("wVk", ctypes.wintypes.WORD),
                    ("wScan", ctypes.wintypes.WORD),
                    ("dwFlags", ctypes.wintypes.DWORD),
                    ("time", ctypes.wintypes.DWORD),
                    ("dwExtraInfo", ctypes.POINTER(ctypes.wintypes.ULONG))
                ]

            class INPUT(ctypes.Structure):
                class _INPUT(ctypes.Union):
                    _fields_ = [("ki", KEYBDINPUT)]
                _anonymous_ = ("ki",)
                _fields_ = [
                    ("type", ctypes.wintypes.DWORD),
                    ("_input", _INPUT)
                ]

            # 按键按下
            input_down = INPUT()
            input_down.type = 1  # INPUT_KEYBOARD
            input_down.ki.wVk = vk_code
            input_down.ki.dwFlags = 0

            # 按键释放
            input_up = INPUT()
            input_up.type = 1  # INPUT_KEYBOARD
            input_up.ki.wVk = vk_code
            input_up.ki.dwFlags = 2  # KEYEVENTF_KEYUP

            # 发送按键事件
            user32.SendInput(2, ctypes.byref(input_down), ctypes.sizeof(INPUT))
            time.sleep(0.01)
            user32.SendInput(2, ctypes.byref(input_up), ctypes.sizeof(INPUT))

        except Exception as ascii_error:
            print(f"ASCII码输入失败 '{char}': {ascii_error}")
            # 最后的回退：跳过该字符
            pass

    def copy_to_clipboard(self, content):
        """复制内容到剪贴板，使用多种方法"""
        try:
            if not content or not content.strip():
                print("内容为空，跳过复制")
                return False

            print(f"准备复制内容，长度: {len(content)}")

            # 方法1: 使用tkinter剪贴板
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                self.root.update()
                print("使用tkinter复制成功")
                return True
            except Exception as e1:
                print(f"tkinter剪贴板失败: {e1}")

            # 方法2: 使用Windows API
            try:
                import ctypes
                user32 = ctypes.windll.user32

                if user32.OpenClipboard(0):
                    try:
                        user32.EmptyClipboard()
                        data = content.encode('utf-16le') + b'\x00\x00'
                        user32.SetClipboardData(13, data)  # CF_UNICODETEXT
                        print("使用Windows API复制成功")
                        return True
                    finally:
                        user32.CloseClipboard()
                else:
                    print("无法打开剪贴板")
            except Exception as e2:
                print(f"Windows API剪贴板失败: {e2}")

            # 方法3: 尝试使用pyperclip
            try:
                import pyperclip
                pyperclip.copy(content)
                print("使用pyperclip复制成功")
                return True
            except ImportError:
                print("pyperclip未安装")
            except Exception as e3:
                print(f"pyperclip剪贴板失败: {e3}")

            print("所有复制方法都失败了")
            return False

        except Exception as e:
            print(f"复制到剪贴板失败: {e}")
            return False

    def stop_execution(self):
        """停止执行"""
        self.is_running = False
        self.status_var.set("停止执行中...")

    def save_config(self):
        """保存配置"""
        try:
            config = {
                'actions': self.actions,
                'pre_delay': self.pre_delay.get(),
                'post_delay': self.post_delay.get()
            }

            with open('clicker_config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)

            messagebox.showinfo("成功", "配置已保存")
            self.status_var.set("配置保存成功")

        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def load_config(self):
        """加载配置"""
        if not os.path.exists('clicker_config.json'):
            messagebox.showinfo("提示", "没有找到配置文件")
            return

        try:
            with open('clicker_config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)

            self.actions = config.get('actions', {})
            self.pre_delay.set(config.get('pre_delay', 1.0))
            self.post_delay.set(config.get('post_delay', 0.5))

            # 更新显示
            for action_name, action in self.actions.items():
                if action_name in self.status_labels:
                    center_x = action['x'] + action['width'] // 2
                    center_y = action['y'] + action['height'] // 2
                    self.status_labels[action_name].set(
                        f"[完成] 区域({action['x']},{action['y']}) 中心({center_x},{center_y}) 大小{action['width']}x{action['height']}"
                    )

            messagebox.showinfo("成功", "配置加载成功")
            self.status_var.set("配置加载成功")

        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {e}")

    def debug_clipboard(self, action_name=None):
        """空函数（已废弃）"""
        pass

    def clear_all(self):
        """清除所有"""
        if messagebox.askyesno("确认", "确定要清除所有配置吗？"):
            self.actions.clear()
            for action_name, status_label in self.status_labels.items():
                status_label.set("未设置区域")
            self.status_var.set("已清除所有配置")

    def show_help(self):
        """显示帮助"""
        help_text = """[范围] 智能自动点击工具 Pro 使用指南

[复制] 基本操作：
1. 点击"框选区域"选择目标位置
2. 点击"测试"验证位置准确性
3. 点击"执行"开始自动点击
4. 使用"顺序执行"完成自动化流程
5. [新增] 使用"范围执行"单独执行指定动作范围

[设置] 十二个动作说明：
• 动作1-3：[蓝色] 基础点击动作
• 动作4：[紫色] 常规操作动作
• 动作5：[红色] 特殊动作（动作4后等待5秒执行）
• 动作6：[青色] 后续动作（动作5后延迟5秒执行）
• 动作7：[橙色] 最终动作（点击类操作的最后一步）
• 动作8：[文档] 文件操作（双击打开TXT文档，选择文件路径）
• 动作9：[复制] 智能文档识别复制（选择文档后自动识别内容并复制到剪贴板，支持TXT、Word、PDF、Excel等格式）
• 动作10：[链接] URL操作（添加URL，点击操作）
• 动作11：[新增] 自动右击（在指定位置执行右击操作）
• 动作12：[新增] 键盘粘贴（不需要框选区域，使用Ctrl+V粘贴到当前焦点位置）

[执行] 完整执行流程：
动作1 → 动作2 → 动作3 → 动作4 → 等待5秒 → 动作5 → 等待5秒 → 动作6 → 动作7 → 动作8 → 动作9 → 动作10 → 动作11 → 动作12

[范围] 范围执行功能：
• [新增] 新增"范围执行"功能，可以单独执行指定范围内的动作
• 例如：设置从动作8到动作12，只执行这5个动作
• 范围内的动作必须已设置（动作12除外）
• 支持所有1-12动作的任意范围组合

[颜色] 颜色标识：
• 蓝色：基础操作
• 绿色：常规操作
• 紫色：中间操作
• 红色：特殊延迟操作
• 青色：后续操作
• 橙色：最终操作
• 橙色：右击操作
• 绿色：键盘粘贴

[警告] 注意事项：
• 所有动作都是可选的，可以只设置需要的动作
• 动作5的特殊延迟和动作6-7的顺序逻辑只有在设置了相应动作时才会执行
• 动作12（键盘粘贴）不需要设置区域，但需要剪贴板有内容
• 确保目标在屏幕可见且不被遮挡
• 执行期间请勿移动鼠标或键盘
• 强烈建议先使用"测试"功能验证位置

[技巧] 高级技巧：
• 可以组合不同的动作来完成复杂的工作流程
• 动作5适合页面加载等待场景
• 动作6-7适合多步骤操作的最后阶段
• 动作11适合需要右击的场景（如打开菜单、查看属性）
• 动作12适合快速粘贴操作，无需定位
• 范围执行功能让你更灵活地控制工作流程
• 所有延迟设置都会叠加生效

[帮助] 快捷操作：
• 框选区域时：按ESC键取消选择
• 执行期间：点击"停止执行"随时停止
• 动作8特殊：可点击"获取鼠标"保存当前鼠标位置
• 动作9功能：选择文档后自动识别内容并复制到剪贴板，支持多种文档格式
• [新增] 范围执行：设置起始和结束动作，点击"执行范围"即可单独运行

如有问题请重新启动程序或点击"清除所有"重置配置。"""
        messagebox.showinfo("[帮助] 使用帮助", help_text)

    def run(self):
        """运行程序"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            pass

class EnhancedPasteMethods:
    """增强的粘贴方法集合"""

    @staticmethod
    def execute_enhanced_paste(action_name, status_callback=None, pre_delay=0, post_delay=0):
        """执行增强的键盘粘贴动作（多方法兼容版）"""
        try:
            if status_callback:
                status_callback(f"执行{action_name}: 键盘粘��...")

            # 前延迟
            if pre_delay > 0:
                time.sleep(pre_delay)

            paste_success = False

            # 多种粘贴方法，按优先级尝试
            paste_methods = [
                EnhancedPasteMethods._paste_method_ctrl_v,           # 方法1：标准Ctrl+V
                EnhancedPasteMethods._paste_method_shift_ins,        # 方法2：Shift+Insert（兼容性好）
                EnhancedPasteMethods._paste_method_alt_edit_p,       # 方法3：Alt+E+P（菜单粘贴）
                EnhancedPasteMethods._paste_method_right_click,      # 方法4：右键菜单粘贴
                EnhancedPasteMethods._paste_method_send_keys,        # 方法5：Windows SendKeys
                EnhancedPasteMethods._paste_method_clipboard_type    # 方法6：剪贴板内容输入
            ]

            for method_index, paste_method in enumerate(paste_methods):
                try:
                    if status_callback:
                        status_callback(f"执行{action_name}: 尝试粘贴方法{method_index+1}")

                    if paste_method(method_index + 1):
                        paste_success = True
                        success_msg = f"[成功] {action_name}执行完成: 粘贴成功(方法{method_index+1})"
                        if status_callback:
                            status_callback(success_msg)
                        print(success_msg)
                        break
                    else:
                        if status_callback:
                            status_callback(f"执行{action_name}: 方法{method_index+1}失败，尝试下一个...")
                        time.sleep(1.0)

                except Exception as attempt_error:
                    print(f"粘贴方法{method_index+1} 失败: {attempt_error}")
                    if status_callback:
                        status_callback(f"执行{action_name}: 方法{method_index+1}异常，准备重试...")
                    time.sleep(1.0)

            # 如果所有方法都失败
            if not paste_success:
                warning_msg = f"[警告] {action_name}执行完成: 所有粘贴方法已执行（请手动验证）"
                if status_callback:
                    status_callback(warning_msg)
                print("[警告] 所有粘贴方法都已尝试，请验证结果")

            # 后延迟
            if post_delay > 0:
                time.sleep(post_delay)

            return True

        except Exception as e:
            error_msg = f"执行键盘粘贴失败: {e}"
            print(error_msg)
            if status_callback:
                status_callback(f"[错误] {action_name}执行失败: {e}")
            return False

    @staticmethod
    def _paste_method_ctrl_v(method_num):
        """方法1：标准Ctrl+V"""
        try:
            # 确保焦点
            pyautogui.click()
            time.sleep(0.5)

            # 标准Ctrl+V
            pyautogui.hotkey('ctrl', 'v', interval=0.1)
            time.sleep(1.0)
            return True
        except:
            return False

    @staticmethod
    def _paste_method_shift_ins(method_num):
        """方法2：Shift+Insert（更好的兼容性）"""
        try:
            pyautogui.click()
            time.sleep(0.5)

            # Shift+Insert粘贴（在远程桌面和终端中更可靠）
            pyautogui.hotkey('shift', 'insert', interval=0.1)
            time.sleep(1.0)
            return True
        except:
            return False

    @staticmethod
    def _paste_method_alt_edit_p(method_num):
        """方法3：Alt+E+P（编辑菜单粘贴）"""
        try:
            pyautogui.click()
            time.sleep(0.5)

            # Alt+E打开编辑菜单，然后按P粘贴
            pyautogui.hotkey('alt', 'e')
            time.sleep(0.3)
            pyautogui.press('p')
            time.sleep(1.0)
            return True
        except:
            return False

    @staticmethod
    def _paste_method_right_click(method_num):
        """方法4：右键菜单粘贴"""
        try:
            # 右键点击当前位置
            pyautogui.rightClick()
            time.sleep(0.5)

            # 尝试多种粘贴快捷键
            paste_keys = ['p', 'v']  # P或V都可能是粘贴

            for key in paste_keys:
                try:
                    pyautogui.press(key)
                    time.sleep(0.5)
                    return True
                except:
                    continue

            return False
        except:
            return False

    @staticmethod
    def _paste_method_send_keys(method_num):
        """方法5：Windows SendKeys API（更可靠）"""
        try:
            # 使用Windows API的SendKeys
            user32 = ctypes.windll.user32

            # 设置焦点
            pyautogui.click()
            time.sleep(0.5)

            # 发送Ctrl+V组合键
            user32.keybd_event(0x11, 0, 0, 0)  # Ctrl按下
            user32.keybd_event(0x56, 0, 0, 0)  # V按下
            time.sleep(0.1)
            user32.keybd_event(0x56, 0, 0x0002, 0)  # V释放
            user32.keybd_event(0x11, 0, 0x0002, 0)  # Ctrl释放

            time.sleep(1.0)
            return True
        except:
            return False

    @staticmethod
    def _paste_method_clipboard_type(method_num):
        """方法6：剪贴板API直接粘贴（如果pyperclip可用）"""
        try:
            import pyperclip

            # 获取剪贴板内容
            content = pyperclip.paste()
            if content and len(content) < 1000:  # 避免输入过长的内容
                # 使用pyautogui逐字输入（慢但可靠）
                pyautogui.click()
                time.sleep(0.5)

                # 清空当前内容（可选）
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)

                # 逐字输入剪贴板内容
                pyautogui.write(content, interval=0.01)
                time.sleep(0.5)
                return True
            return False
        except:
            return False

class EnhancedRightClickMethods:
    """增强的右键点击方法集合"""

    @staticmethod
    def _get_method_entries():
        """统一管理右键方法列表，便于诊断和日志展示"""
        return [
            ("增强时序控制", EnhancedRightClickMethods._method_enhanced_timing),
            ("Windows API", EnhancedRightClickMethods._method_windows_api),
            ("SendInput绝对坐标", EnhancedRightClickMethods._method_send_input_absolute),
            ("键盘模拟右键菜单", EnhancedRightClickMethods._method_keyboard_context_menu),
            ("管理员权限API", EnhancedRightClickMethods._method_admin_privilege_api),
            ("手动分步右键", EnhancedRightClickMethods._method_manual_right_click),
            ("click右键", EnhancedRightClickMethods._method_click_button_right),
            ("pyautogui.rightClick", EnhancedRightClickMethods._method_pyautogui_right_click),
            ("重延迟备选", EnhancedRightClickMethods._method_fallback_heavy_delay),
        ]

    @staticmethod
    def execute_enhanced_right_click(x, y, status_callback=None):
        """执行增强的右键点击（多种方法fallback）"""
        try:
            if status_callback:
                status_callback("执行右键点击...")

            # 移动到目标位置
            pyautogui.moveTo(x, y, duration=0.3)
            time.sleep(0.2)

            # 尝试多种方法（针对特定软件优化）
            methods = EnhancedRightClickMethods._get_method_entries()

            for method_index, (method_name, method) in enumerate(methods):
                try:
                    if status_callback:
                        status_callback(f"尝试右键方法{method_index+1}: {method_name}")

                    # 重新移动到目标位置（某些方法可能会移动鼠标）
                    pyautogui.moveTo(x, y, duration=0.1)
                    time.sleep(0.1)

                    result = method()
                    if result:
                        if status_callback:
                            status_callback(f"右键成功(方法{method_index+1}: {method_name})")
                        return True
                    else:
                        if status_callback:
                            status_callback(f"右键方法{method_index+1}失败，尝试下一个... ({method_name})")
                        time.sleep(0.5)

                except Exception as e:
                    print(f"右键方法{method_index+1}异常: {e}")
                    if status_callback:
                        status_callback(f"右键方法{method_index+1}异常 ({method_name})")
                    time.sleep(0.5)

            if status_callback:
                status_callback("所有右键方法都失败")
            return False

        except Exception as e:
            print(f"执行右键点击失败: {e}")
            if status_callback:
                status_callback(f"右键执行异常: {e}")
            return False

    @staticmethod
    def _method_enhanced_timing():
        """方法1：增强时序控制"""
        try:
            # 确保焦点
            time.sleep(0.1)

            # 点击左键获得焦点
            pyautogui.click()
            time.sleep(0.1)

            # 轻微移动后再右击（避免被忽略）
            pyautogui.moveRel(5, 0, duration=0.1)
            time.sleep(0.1)

            # 右击
            pyautogui.rightClick()
            return True
        except:
            return False

    @staticmethod
    def _method_windows_api():
        """方法2：Windows API直接调用"""
        try:
            user32 = ctypes.windll.user32

            # 获取当前鼠标位置
            x, y = pyautogui.position()

            # 发送右键按下消息
            user32.SetCursorPos(x, y)
            user32.mouse_event(0x0008, 0, 0, 0, 0)  # MOUSEEVENTF_RIGHTDOWN
            time.sleep(0.1)
            user32.mouse_event(0x0010, 0, 0, 0, 0)  # MOUSEEVENTF_RIGHTUP

            return True
        except Exception as e:
            print(f"Windows API 错误: {e}")
            return False

    @staticmethod
    def _method_manual_right_click():
        """方法6：手动分步按下释放"""
        try:
            pyautogui.mouseDown(button='right')
            time.sleep(0.1)
            pyautogui.mouseUp(button='right')
            return True
        except:
            return False

    @staticmethod
    def _method_click_button_right():
        """方法7：使用button参数"""
        try:
            pyautogui.click(button='right')
            return True
        except:
            return False

    @staticmethod
    def _method_pyautogui_right_click():
        """方法8：标准pyautogui右键"""
        try:
            pyautogui.rightClick()
            return True
        except:
            return False

    @staticmethod
    def _method_keyboard_context_menu():
        """方法4：键盘模拟右键菜单（适用于拦截鼠标的软件）"""
        try:
            # 先确保焦点在目标位置
            pyautogui.click()  # 左键点击获得焦点
            time.sleep(0.3)

            # 使用键盘快捷键Shift+F10或ContextMenu键模拟右键菜单
            # 方法A：Shift+F10
            pyautogui.hotkey('shift', 'f10')
            time.sleep(0.5)

            # 如果Shift+F10不工作，尝试Windows Menu键
            try:
                # 尝试发送ContextMenu键（有些键盘有这个键）
                pyautogui.press('menu')
            except:
                pass

            return True
        except:
            return False

    @staticmethod
    def _method_send_input_absolute():
        """方法3：使用SendInput发送绝对坐标右键（兼容全屏/游戏类窗口）"""
        try:
            user32 = ctypes.windll.user32

            # 获取屏幕尺寸，转换为SendInput绝对坐标（0-65535）
            screen_width = user32.GetSystemMetrics(0) - 1
            screen_height = user32.GetSystemMetrics(1) - 1
            x, y = pyautogui.position()
            abs_x = int(x * 65535 / screen_width)
            abs_y = int(y * 65535 / screen_height)

            # 定义必要结构体
            class MOUSEINPUT(ctypes.Structure):
                _fields_ = [
                    ("dx", ctypes.c_long),
                    ("dy", ctypes.c_long),
                    ("mouseData", ctypes.c_ulong),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
                ]

            class INPUT(ctypes.Structure):
                _fields_ = [
                    ("type", ctypes.c_ulong),
                    ("mi", MOUSEINPUT),
                ]

            MOUSEEVENTF_MOVE = 0x0001
            MOUSEEVENTF_ABSOLUTE = 0x8000
            MOUSEEVENTF_RIGHTDOWN = 0x0008
            MOUSEEVENTF_RIGHTUP = 0x0010

            # 移动到绝对位置后发送右键按下/抬起
            inputs = (INPUT * 3)()
            inputs[0].type = 0
            inputs[0].mi = MOUSEINPUT(abs_x, abs_y, 0, MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, 0, None)

            inputs[1].type = 0
            inputs[1].mi = MOUSEINPUT(abs_x, abs_y, 0, MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_ABSOLUTE, 0, None)

            inputs[2].type = 0
            inputs[2].mi = MOUSEINPUT(abs_x, abs_y, 0, MOUSEEVENTF_RIGHTUP | MOUSEEVENTF_ABSOLUTE, 0, None)

            sent = user32.SendInput(3, ctypes.byref(inputs), ctypes.sizeof(INPUT))
            return sent == 3
        except Exception as e:
            print(f"SendInput 绝对坐标右键失败: {e}")
            return False

    @staticmethod
    def _method_admin_privilege_api():
        """方法5：管理员权限Windows API（增强权限级别）"""
        try:
            import ctypes
            from ctypes import wintypes

            # 定义Windows API函数
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            # 尝试提升权限
            SE_PRIVILEGE_ENABLED = 0x00000002
            TOKEN_ALL_ACCESS = 0xF01FF

            # 获取当前进程令牌
            hToken = wintypes.HANDLE()
            if not user32.OpenProcessToken(kernel32.GetCurrentProcess(), TOKEN_ALL_ACCESS, ctypes.byref(hToken)):
                # 如果失败，继续使用普通方法
                pass

            # 发送低级别鼠标消息
            x, y = pyautogui.position()

            # 使用SendMessage而不是mouse_event（更直接）
            HWND = user32.WindowFromPoint((x, y))
            if HWND:
                WM_RBUTTONDOWN = 0x0204
                WM_RBUTTONUP = 0x0205
                MK_RBUTTON = 0x0002

                lparam = (y << 16) | x
                user32.SendMessageA(HWND, WM_RBUTTONDOWN, MK_RBUTTON, lparam)
                time.sleep(0.1)
                user32.SendMessageA(HWND, WM_RBUTTONUP, 0, lparam)

                return True
            else:
                # 备选方案：使用SetCursorPos + mouse_event
                user32.SetCursorPos(x, y)
                user32.mouse_event(0x0008, 0, 0, 0, 0)  # MOUSEEVENTF_RIGHTDOWN
                time.sleep(0.2)  # 增加延迟
                user32.mouse_event(0x0010, 0, 0, 0, 0)  # MOUSEEVENTF_RIGHTUP
                return True

        except Exception as e:
            print(f"管理员API方法错误: {e}")
            return False

    @staticmethod
    def _method_fallback_heavy_delay():
        """方法9：重延迟备选方案（适用于反应慢的软件）"""
        try:
            # 确保焦点
            pyautogui.click()
            time.sleep(1.0)  # 较长延迟确保获得焦点

            # 移动到目标位置
            x, y = pyautogui.position()
            pyautogui.moveTo(x, y, duration=0.5)
            time.sleep(0.5)

            # 分步执行，每步之间有较长时间延迟
            pyautogui.mouseDown(button='right')
            time.sleep(0.5)  # 长延迟
            pyautogui.mouseUp(button='right')
            time.sleep(0.5)

            # 再尝试一次标准右键
            pyautogui.rightClick()

            return True
        except:
            return False

    @staticmethod
    def diagnose_right_click_methods(x, y):
        """逐个执行所有右键方法并返回诊断结果，方便定位问题"""
        results = []
        methods = EnhancedRightClickMethods._get_method_entries()

        for method_index, (method_name, method) in enumerate(methods):
            try:
                # 确保每次诊断都回到目标位置
                pyautogui.moveTo(x, y, duration=0.1)
                time.sleep(0.1)

                success = bool(method())
                results.append({
                    "index": method_index + 1,
                    "name": method_name,
                    "success": success,
                    "error": None
                })
            except Exception as e:
                results.append({
                    "index": method_index + 1,
                    "name": method_name,
                    "success": False,
                    "error": str(e)
                })
        return results

if __name__ == "__main__":
    try:
        print("启动自动点击工具...")
        app = SimpleClickerWorking()
        app.run()
    except Exception as e:
        print(f"启动失败: {e}")
        input("按回车键退出...")