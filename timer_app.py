# -*- coding: utf-8 -*-
"""
多任务队列计时闹钟
使用Python原生tkinter库开发，无第三方依赖
配色：黑/白/灰主色调，时间数字绿色
作者：Mega_HUGO
"""

import tkinter as tk
from tkinter import messagebox
import winsound
import re
import sys
import os
import json
import subprocess
from datetime import datetime, timedelta
import calendar
import ctypes
from ctypes import windll


def resource_path(relative_path):
    """获取资源文件绝对路径，兼容开发环境和PyInstaller打包环境"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller打包后的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    # 开发环境
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def get_logs_dir():
    """获取日志目录路径"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller打包后，使用exe所在目录
        return os.path.join(os.path.dirname(sys.executable), 'logs')
    # 开发环境
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')


class TaskItem:
    """单条任务数据结构"""
    def __init__(self, task_id):
        self.task_id = task_id
        self.name = f"任务{task_id}"
        self.duration = 25  # 任务预估时长（分钟）
        self.break_count = 0  # 中途休息次数
        self.break_duration = 5  # 单次休息时长（分钟）
        self.completed = False  # 任务是否已完成
        self.skipped = False  # 任务是否被跳过
        self.early_completed = False  # 任务是否提前完成
        self.jump_terminated = False  # 是否被跳转终止
        
        # 时间追踪
        self.start_time = None  # 任务开始时间
        self.actual_duration = 0  # 实际耗时（秒）
        
        # 跳转终止时的时间记录
        self.actual_focus_minutes = 0  # 实际专注时长（分钟）
        self.actual_breaks_taken = 0  # 实际已用休息次数
        self.remaining_duration = 0  # 剩余预估时长（分钟）
        self.remaining_breaks = 0  # 剩余休息次数
        
        # UI组件引用
        self.frame = None
        self.seq_label = None  # 序号标签
        self.name_var = None
        self.duration_var = None
        self.break_count_var = None
        self.break_duration_var = None
        self.delete_btn = None
        self.drag_handle = None  # 拖拽手柄
    
    def calculate_work_segments(self):
        """计算总工作段数"""
        return self.break_count + 1
    
    def calculate_single_work_duration(self):
        """计算单段工作时长（分钟）"""
        segments = self.calculate_work_segments()
        return self.duration / segments
    
    def to_dict(self, completion_type, completed_at):
        """转换为字典用于日志记录
        
        Args:
            completion_type: 完成类型 (normal_complete/early_complete/skipped/jump_terminated)
            completed_at: 完成时间
        """
        record = {
            "id": self.task_id,
            "name": self.name,
            "duration": self.duration,
            "break_count": self.break_count,
            "break_duration": self.break_duration,
            "completion_type": completion_type,
            "actual_duration_seconds": self.actual_duration if completion_type in ("early_complete", "jump_terminated") else None,
            "completed_at": completed_at
        }
        
        # 跳转终止时额外记录详细数据
        if completion_type == "jump_terminated":
            record["actual_focus_minutes"] = self.actual_focus_minutes
            record["actual_breaks_taken"] = self.actual_breaks_taken
            record["remaining_duration"] = self.remaining_duration
            record["remaining_breaks"] = self.remaining_breaks
        
        return record


class Logger:
    """日志记录器"""
    
    def __init__(self):
        self.logs_dir = get_logs_dir()
        self._ensure_logs_dir()
    
    def _ensure_logs_dir(self):
        """确保日志目录存在"""
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)
    
    def _get_log_file_path(self, date_str=None):
        """获取指定日期的日志文件路径"""
        if date_str is None:
            date_str = datetime.now().strftime("%Y-%m-%d")
        return os.path.join(self.logs_dir, f"{date_str}.json")
    
    def log_task(self, task, completion_type):
        """记录任务完成/跳过
        
        Args:
            task: 任务对象
            completion_type: 完成类型 (normal_complete/early_complete/skipped)
        """
        now = datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        completed_at = now.strftime("%Y-%m-%d %H:%M:%S")
        
        # 读取现有日志
        log_data = self._read_log(date_str)
        
        # 添加任务记录
        task_record = task.to_dict(completion_type, completed_at)
        log_data["tasks"].append(task_record)
        log_data["updated_at"] = completed_at
        
        # 保存日志
        self._write_log(date_str, log_data)
    
    def _read_log(self, date_str):
        """读取指定日期的日志"""
        file_path = self._get_log_file_path(date_str)
        
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        
        # 返回空日志结构
        return {
            "date": date_str,
            "tasks": [],
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_at": None
        }
    
    def _write_log(self, date_str, log_data):
        """写入日志文件"""
        file_path = self._get_log_file_path(date_str)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)
    
    def get_log(self, date_str):
        """获取指定日期的日志"""
        return self._read_log(date_str)
    
    def clear_today_log(self):
        """清除当天的日志记录"""
        date_str = datetime.now().strftime("%Y-%m-%d")
        file_path = self._get_log_file_path(date_str)
        if os.path.exists(file_path):
            os.remove(file_path)
    
    def get_month_logs(self, year, month):
        """获取指定月份的所有日志"""
        logs = {}
        # 获取该月的天数
        _, days_in_month = calendar.monthrange(year, month)
        
        for day in range(1, days_in_month + 1):
            date_str = f"{year:04d}-{month:02d}-{day:02d}"
            log = self._read_log(date_str)
            if log["tasks"]:
                logs[date_str] = log
        
        return logs
    
    def get_month_statistics(self, year, month):
        """获取月度统计数据"""
        logs = self.get_month_logs(year, month)
        
        # 统计数据
        total_days = len(logs)  # 有任务的天数
        total_tasks = 0
        completed_tasks = 0  # 包含 normal_complete 和 early_complete
        early_completed_tasks = 0
        skipped_tasks = 0
        jump_terminated_tasks = 0  # 跳转终止的任务
        total_focus_minutes = 0  # 总预估专注时长
        total_break_minutes = 0  # 总休息时长
        completed_focus_minutes = 0  # 已完成任务专注时长（实际耗时）
        
        for date_str, log in logs.items():
            for task in log["tasks"]:
                total_tasks += 1
                completion_type = task.get("completion_type", "normal_complete")
                
                # 跳过的任务不参与统计
                if completion_type == "skipped":
                    skipped_tasks += 1
                    continue
                
                # 跳转终止的任务：记录实际专注时长
                if completion_type == "jump_terminated":
                    jump_terminated_tasks += 1
                    # 使用实际专注时长
                    actual_minutes = task.get("actual_focus_minutes", 0) or 0
                    if actual_minutes == 0:
                        # 兼容旧数据，从 actual_duration_seconds 计算
                        actual_seconds = task.get("actual_duration_seconds", 0) or 0
                        actual_minutes = actual_seconds / 60
                    completed_focus_minutes += actual_minutes
                    # 已用的休息时长计入
                    actual_breaks = task.get("actual_breaks_taken", 0) or 0
                    break_duration = task.get("break_duration", 5)
                    total_break_minutes += actual_breaks * break_duration
                    continue
                
                # 正常完成或提前完成
                completed_tasks += 1
                total_break_minutes += task["break_count"] * task["break_duration"]
                
                if completion_type == "early_complete":
                    early_completed_tasks += 1
                    # 提前完成：使用实际耗时
                    actual_seconds = task.get("actual_duration_seconds", 0) or 0
                    completed_focus_minutes += actual_seconds / 60
                    total_focus_minutes += task["duration"]  # 预估时长也计入
                else:
                    # 正常完成：使用预估时长
                    completed_focus_minutes += task["duration"]
                    total_focus_minutes += task["duration"]
        
        # 计算完成率（基于非跳过和非跳转终止的任务）
        non_skipped_tasks = total_tasks - skipped_tasks - jump_terminated_tasks
        completion_rate = (completed_tasks / non_skipped_tasks * 100) if non_skipped_tasks > 0 else 0
        
        # 计算平均数
        avg_daily_tasks = completed_tasks / total_days if total_days > 0 else 0
        avg_task_duration = completed_focus_minutes / completed_tasks if completed_tasks > 0 else 0
        
        return {
            "total_days": total_days,
            "total_tasks": total_tasks,
            "completed_tasks": completed_tasks,
            "early_completed_tasks": early_completed_tasks,
            "skipped_tasks": skipped_tasks,
            "jump_terminated_tasks": jump_terminated_tasks,
            "completion_rate": completion_rate,
            "total_focus_minutes": total_focus_minutes,
            "completed_focus_minutes": completed_focus_minutes,
            "total_break_minutes": total_break_minutes,
            "avg_daily_tasks": avg_daily_tasks,
            "avg_task_duration": avg_task_duration
        }
    
    def export_month_to_csv(self, year, month, file_path):
        """导出月度日志为CSV"""
        logs = self.get_month_logs(year, month)
        
        lines = []
        lines.append("日期,任务ID,任务名称,预估时长(分钟),实际耗时(分钟),休息次数,休息时长(分钟),状态,完成时间")
        
        for date_str in sorted(logs.keys()):
            log = logs[date_str]
            for task in log["tasks"]:
                completion_type = task.get("completion_type", "normal_complete")
                if completion_type == "normal_complete":
                    status_text = "已完成"
                    actual_duration = ""
                elif completion_type == "early_complete":
                    status_text = "提前完成"
                    actual_seconds = task.get("actual_duration_seconds", 0) or 0
                    actual_duration = f"{actual_seconds / 60:.1f}"
                elif completion_type == "jump_terminated":
                    status_text = "跳转终止"
                    actual_minutes = task.get("actual_focus_minutes", 0) or 0
                    if actual_minutes == 0:
                        actual_seconds = task.get("actual_duration_seconds", 0) or 0
                        actual_minutes = actual_seconds / 60
                    actual_duration = f"{actual_minutes:.1f}"
                else:
                    status_text = "已跳过"
                    actual_duration = ""
                
                lines.append(f'{date_str},{task["id"]},"{task["name"]}",{task["duration"]},{actual_duration},{task["break_count"]},{task["break_duration"]},{status_text},{task["completed_at"]}')
        
        with open(file_path, 'w', encoding='utf-8-sig') as f:
            f.write('\n'.join(lines))


class TimerApp:
    """多任务队列计时闹钟主程序"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("多任务队列计时闹钟")
        
        # 列宽定义（像素）- 统一表头和数据行
        self.col_widths = [40, 320, 90, 90, 90, 50]  # 拖拽、任务名称、预估时长、休息次数、休息时长、删除
        self.total_width = sum(self.col_widths) + 40  # 总宽度 + padding
        
        # 设置窗口图标
        try:
            icon_path = resource_path('clock_icon.ico')
            self.root.iconbitmap(icon_path)
        except Exception:
            pass  # 图标加载失败时使用默认图标
        
        self.root.resizable(True, True)  # 启用窗口调整大小
        # 窗口最小尺寸
        self.root.minsize(self.total_width + 200, 500)
        
        # 确保窗口有正常的边框样式（修复 Windows 11 边框光标问题）
        self._fix_window_border_cursor()
        
        # 绑定窗口大小变化事件
        self.root.bind("<Configure>", self._on_window_resize)
        
        # 初始化日志记录器
        self.logger = Logger()
        
        # 任务列表
        self.tasks = []
        self.task_counter = 0
        
        # 计时状态
        self.is_running = False
        self.is_paused = False
        self.current_task_index = 0
        self.current_work_segment = 0
        self.current_break_count = 0
        self.is_in_break = False
        self.remaining_seconds = 0
        self.completed_tasks = 0
        
        # 修改模式状态（用于"继续任务"功能）
        self.is_edit_mode = False
        self.saved_task_index = 0
        self.saved_work_segment = 0
        self.saved_break_count = 0
        self.saved_is_in_break = False
        
        # 任务结束休息状态（任务完成后自动5分钟休息）
        self.is_post_task_break = False  # 是否处于任务结束休息状态
        
        # 计时器ID
        self.timer_id = None
        
        # 任务选中状态（用于跳转功能）
        self.selected_task_index = None
        
        # 拖拽排序状态
        self.drag_source_row = None
        self.drag_source_task = None
        self.last_highlighted_target = None  # 拖拽时上一个高亮的目标
        
        # 创建UI
        self._create_ui()
        
        # 初始化5条默认任务
        self._init_default_tasks()
        
        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _create_ui(self):
        """创建整体UI布局"""
        # 顶部标题区
        self._create_title_area()
        
        # 任务队列区（带滚动）
        self._create_task_queue_area()
        
        # 计时与状态显示区
        self._create_timer_display_area()
        
        # 全局控制按钮区
        self._create_control_buttons_area()
        
        # 底部作者信息
        self._create_footer()
    
    def _create_title_area(self):
        """创建顶部标题区"""
        title_frame = tk.Frame(self.root, pady=10, bg="#f0f0f0")
        title_frame.pack(fill=tk.X)
        
        # 标题
        title_label = tk.Label(
            title_frame, 
            text="⏰ 多任务队列计时闹钟", 
            font=("Microsoft YaHei", 18, "bold"),
            bg="#f0f0f0", fg="#333333"
        )
        title_label.pack(side=tk.LEFT, padx=20)
        
        # 日志统计按钮（先pack，确保在最右边）
        self.log_btn = tk.Button(
            title_frame, text="📅 日志统计", 
            command=self._show_calendar_window,
            font=("Microsoft YaHei", 10),
            bg="#333333", fg="white", activebackground="#555555",
            padx=10, pady=2
        )
        self.log_btn.pack(side=tk.RIGHT, padx=10)
        
        # 任务间休整时间设置（后pack，在日志统计按钮左边）
        post_break_frame = tk.Frame(title_frame, bg="#f0f0f0")
        post_break_frame.pack(side=tk.RIGHT, padx=5)
        
        post_break_label = tk.Label(
            post_break_frame, text="任务间休整(分钟):",
            font=("Microsoft YaHei", 9), bg="#f0f0f0", fg="#333333"
        )
        post_break_label.pack(side=tk.LEFT)
        
        self.post_break_var = tk.StringVar(value="5")
        self.post_break_entry = tk.Entry(
            post_break_frame, textvariable=self.post_break_var,
            width=4, font=("Microsoft YaHei", 9), justify=tk.CENTER
        )
        self.post_break_entry.pack(side=tk.LEFT, padx=2)
        self.post_break_entry.bind("<KeyRelease>", lambda e: self._update_total_duration())
    
    def _create_task_queue_area(self):
        """创建任务队列区（支持滚动）"""
        # 外层容器
        self.queue_outer_frame = tk.LabelFrame(self.root, text="任务队列", padx=5, pady=5, bg="#f0f0f0", fg="#333333")
        self.queue_outer_frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        
        # 创建Canvas和Scrollbar实现滚动
        self.canvas_frame = tk.Frame(self.queue_outer_frame, bg="#f0f0f0")
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas 不设固定宽度，自适应窗口
        self.canvas = tk.Canvas(self.canvas_frame, height=210, 
                               bg="#f0f0f0", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        
        # 统一的表格容器（类似 HTML <table>）
        self.table_frame = tk.Frame(self.canvas, bg="#f0f0f0")
        self.table_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.table_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定鼠标滚轮
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # ========== 核心：统一设置列宽（类似 HTML <col width="80px">）==========
        for i, width in enumerate(self.col_widths):
            self.table_frame.grid_columnconfigure(i, minsize=width, weight=0)
        
        # ========== 表头行（直接放在 table_frame 的 row=0）==========
        headers = ["", "任务名称", "预估时长", "休息次数", "休息时长", ""]
        for i, header in enumerate(headers):
            lbl = tk.Label(self.table_frame, text=header,
                          font=("Microsoft YaHei", 9, "bold"), 
                          bg="#e0e0e0", fg="#333333",
                          padx=2, pady=5)
            lbl.grid(row=0, column=i, sticky="ew")
        
        # 数据行从 row=1 开始
        self.next_row = 1
        
        # 任务队列控制按钮
        btn_frame = tk.Frame(self.queue_outer_frame, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.add_task_btn = tk.Button(
            btn_frame, text="添加新任务", width=12, 
            command=self._add_new_task, font=("Microsoft YaHei", 9),
            bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
        )
        self.add_task_btn.pack(side=tk.LEFT, padx=5)
        
        self.import_btn = tk.Button(
            btn_frame, text="批量导入任务", width=12,
            command=self._show_import_dialog, font=("Microsoft YaHei", 9),
            bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
        )
        self.import_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_all_btn = tk.Button(
            btn_frame, text="清空所有任务", width=12,
            command=self._clear_all_tasks, font=("Microsoft YaHei", 9),
            bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
        )
        self.clear_all_btn.pack(side=tk.LEFT, padx=5)
        
        self.jump_to_task_btn = tk.Button(
            btn_frame, text="跳到选定任务", width=12,
            command=self._jump_to_selected_task, font=("Microsoft YaHei", 9),
            bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0", state=tk.DISABLED
        )
        self.jump_to_task_btn.pack(side=tk.LEFT, padx=5)
    
    def _on_mousewheel(self, event):
        """鼠标滚轮事件处理"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def _fix_window_border_cursor(self):
        """修复 Windows 边框光标问题"""
        try:
            # 获取窗口句柄
            self.root.update_idletasks()
            hwnd = windll.user32.GetParent(self.root.winfo_id())
            
            # 获取当前窗口样式
            GWL_STYLE = -16
            style = windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
            
            # 确保窗口有可调整大小的边框样式
            WS_THICKFRAME = 0x00040000
            WS_MAXIMIZEBOX = 0x00010000
            
            # 如果样式没有包含 WS_THICKFRAME，添加它
            if not (style & WS_THICKFRAME):
                style |= WS_THICKFRAME
                windll.user32.SetWindowLongW(hwnd, GWL_STYLE, style)
            
            # 强制窗口重绘
            SWP_FRAMECHANGED = 0x0020
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_NOZORDER = 0x0004
            windll.user32.SetWindowPos(hwnd, None, 0, 0, 0, 0, 
                                       SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER)
        except Exception:
            pass  # 如果 Windows API 调用失败，忽略错误
    
    def _on_window_resize(self, event):
        """窗口大小变化时调整内部控件"""
        # 只处理来自 root 的事件
        if event.widget != self.root:
            return
        
        # 获取窗口宽度，减去边距和滚动条宽度
        window_width = event.width
        table_width = max(self.total_width, window_width - 50)  # 至少保持 total_width
        
        # 更新 Canvas 内的 table_frame 宽度
        self.canvas.itemconfig(self.canvas_window, width=table_width)
    
    def _on_task_select(self, task):
        """点击选中任务（用于跳转功能）"""
        task_index = self.tasks.index(task)
        
        # 不能选中已完成/跳过/正在执行的任务
        if task.completed or task.skipped:
            return
        if task_index == self.current_task_index and self.is_running:
            return
        
        # 切换选中状态
        if self.selected_task_index == task_index:
            self.selected_task_index = None  # 取消选中
        else:
            self.selected_task_index = task_index  # 选中新任务
        
        # 更新高亮显示
        self._highlight_current_task()
    
    def _on_drag_start(self, event, task):
        """开始拖拽"""
        task_index = self.tasks.index(task)
        
        # 已完成/跳过的任务不能拖拽
        if task.completed or task.skipped:
            return
        
        # 计时中：正在执行的任务（第一行）不能拖拽
        if self.is_running and task_index == 0:
            return
        
        self.drag_source_task = task
        self.drag_source_row = task_index
        self.last_highlighted_target = None  # 记录上一个高亮的目标
        
        # 高亮拖拽源行（绿色）
        try:
            task.drag_handle.config(bg="#22C55E", fg="white")
        except:
            pass
    
    def _on_drag_motion(self, event, task):
        """拖拽移动中 - 高亮当前悬停的目标拖拽符号"""
        if self.drag_source_task is None:
            return
        
        # 先清除上一个高亮的目标（如果不是源）
        if self.last_highlighted_target and self.last_highlighted_target != self.drag_source_task:
            try:
                self.last_highlighted_target.drag_handle.config(bg="#f0f0f0", fg="#888888")
            except:
                pass
        
        # 获取鼠标位置下的控件
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        if widget:
            # 找到该控件所在的行
            for i, t in enumerate(self.tasks):
                if hasattr(t, 'drag_handle') and t.drag_handle == widget:
                    # 计时中：不能拖到第一行（正在执行的任务前面）
                    if self.is_running and i == 0:
                        self.last_highlighted_target = None
                        return
                    # 高亮目标行（黄色），但不能是已完成/跳过的
                    if not t.completed and not t.skipped and t != self.drag_source_task:
                        try:
                            t.drag_handle.config(bg="#FEF3C7", fg="#92400E")
                            self.last_highlighted_target = t
                        except:
                            pass
                    else:
                        self.last_highlighted_target = None
                    return
        self.last_highlighted_target = None
    
    def _on_drag_release(self, event, task):
        """拖拽释放"""
        if self.drag_source_task is None:
            return
        
        source_idx = self.drag_source_row
        
        # 获取鼠标释放位置下的控件，找到目标行
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        target_index = None
        
        if widget:
            # 遍历找到目标task
            for i, t in enumerate(self.tasks):
                # 检查鼠标是否在拖拽手柄上
                if hasattr(t, 'drag_handle') and t.drag_handle == widget:
                    target_index = i
                    break
        
        # 如果没找到目标，尝试通过行号查找
        if target_index is None and widget:
            try:
                grid_info = widget.grid_info()
                row = int(grid_info.get('row', -1))
                if row > 0:  # 表头是第0行
                    for i, t in enumerate(self.tasks):
                        if t.row_idx == row:
                            target_index = i
                            break
            except:
                pass
        
        # 清除高亮
        self._clear_drag_state()
        
        # 无效目标或原地释放
        if target_index is None or target_index == source_idx:
            self._highlight_current_task()
            return
        
        # 计时中：不能拖到第一行（正在执行的任务前面）
        if self.is_running and target_index == 0:
            self._highlight_current_task()
            return
        
        # 不能拖到已完成/跳过的任务位置
        target_task = self.tasks[target_index]
        if target_task.completed or target_task.skipped:
            self._highlight_current_task()
            return
        
        # 执行重排
        moved_task = self.tasks.pop(source_idx)
        
        # 计算插入位置（因为已经pop了一个元素）
        if target_index > source_idx:
            target_index -= 1
        self.tasks.insert(target_index, moved_task)
        
        # 更新 current_task_index（仅未计时状态需要）
        # 计时状态下当前任务始终在第一行（index 0），不受拖拽影响
        if not self.is_running:
            if source_idx < self.current_task_index <= target_index + 1:
                self.current_task_index -= 1
            elif target_index <= self.current_task_index < source_idx:
                self.current_task_index += 1
        
        # 更新选中状态
        if self.selected_task_index is not None:
            if self.selected_task_index == source_idx:
                self.selected_task_index = target_index
            elif source_idx < self.selected_task_index <= target_index + 1:
                self.selected_task_index -= 1
            elif target_index <= self.selected_task_index < source_idx:
                self.selected_task_index += 1
        
        # 清除拖拽状态并刷新
        self._clear_drag_state()
        self._refresh_task_table()
        self._update_total_duration()
    
    def _clear_drag_state(self):
        """清除拖拽状态"""
        # 清除源行高亮
        if self.drag_source_task:
            try:
                self.drag_source_task.drag_handle.config(bg="#f0f0f0", fg="#888888")
            except:
                pass
        # 清除目标行高亮
        if self.last_highlighted_target:
            try:
                self.last_highlighted_target.drag_handle.config(bg="#f0f0f0", fg="#888888")
            except:
                pass
        self.drag_source_task = None
        self.drag_source_row = None
        self.last_highlighted_target = None
    
    def _refresh_task_table(self):
        """刷新任务表格显示"""
        # 清除所有任务行的UI
        for task in self.tasks:
            for col in range(6):
                widget = self.table_frame.grid_slaves(row=task.row_idx, column=col)
                if widget:
                    widget[0].destroy()
        
        # 重置行号并重新创建
        self.next_row = 1
        old_tasks = self.tasks.copy()
        self.tasks.clear()
        
        for task in old_tasks:
            row_idx = self.next_row
            self.next_row += 1
            
            # 第0列：拖拽手柄
            task.drag_handle = tk.Label(
                self.table_frame, text="⋮⋮",
                font=("Microsoft YaHei", 9), bg="#f0f0f0", fg="#888888",
                cursor="hand2"
            )
            task.drag_handle.grid(row=row_idx, column=0, padx=2, pady=2, sticky="ew")
            task.drag_handle.bind("<Button-1>", lambda e, t=task: self._on_drag_start(e, t))
            task.drag_handle.bind("<B1-Motion>", lambda e, t=task: self._on_drag_motion(e, t))
            task.drag_handle.bind("<ButtonRelease-1>", lambda e, t=task: self._on_drag_release(e, t))
            
            # 第1列：任务名称
            task.name_var = tk.StringVar(value=task.name)
            name_entry = tk.Entry(self.table_frame, textvariable=task.name_var,
                                 font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
            name_entry.grid(row=row_idx, column=1, padx=2, pady=2, sticky="ew")
            name_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
            
            # 第2列：预估时长
            task.duration_var = tk.StringVar(value=str(task.duration))
            duration_entry = tk.Entry(self.table_frame, textvariable=task.duration_var,
                                     font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
            duration_entry.grid(row=row_idx, column=2, padx=2, pady=2, sticky="ew")
            duration_entry.bind("<KeyRelease>", lambda e: self._update_total_duration())
            duration_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
            
            # 第3列：休息次数
            task.break_count_var = tk.StringVar(value=str(task.break_count))
            break_count_entry = tk.Entry(self.table_frame, textvariable=task.break_count_var,
                                        font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
            break_count_entry.grid(row=row_idx, column=3, padx=2, pady=2, sticky="ew")
            break_count_entry.bind("<KeyRelease>", lambda e: self._update_total_duration())
            break_count_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
            
            # 第4列：休息时长
            task.break_duration_var = tk.StringVar(value=str(task.break_duration))
            break_duration_entry = tk.Entry(self.table_frame, textvariable=task.break_duration_var,
                                           font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
            break_duration_entry.grid(row=row_idx, column=4, padx=2, pady=2, sticky="ew")
            break_duration_entry.bind("<KeyRelease>", lambda e: self._update_total_duration())
            break_duration_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
            
            # 第5列：删除按钮
            task.delete_btn = tk.Button(
                self.table_frame, text="删除",
                command=lambda t=task: self._delete_task(t),
                font=("Microsoft YaHei", 9), bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
            )
            task.delete_btn.grid(row=row_idx, column=5, padx=2, pady=2, sticky="w")
            
            task.row_idx = row_idx
            self.tasks.append(task)
        
        # 更新按钮状态和高亮
        self._update_delete_buttons()
        self._highlight_current_task()
        self._update_selection_highlight()
        self._apply_ui_lock_state()
    
    def _create_timer_display_area(self):
        """创建计时与状态显示区"""
        timer_frame = tk.LabelFrame(self.root, text="计时状态", padx=10, pady=10, bg="#f0f0f0", fg="#333333")
        timer_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 倒计时显示（绿色）
        self.timer_label = tk.Label(
            timer_frame, text="00:00", 
            font=("Microsoft YaHei", 48, "bold"),
            fg="#22C55E", bg="#f0f0f0"
        )
        self.timer_label.pack(pady=10)
        
        # 状态信息
        info_frame = tk.Frame(timer_frame, bg="#f0f0f0")
        info_frame.pack(fill=tk.X)
        
        self.status_label = tk.Label(
            info_frame, text="当前状态：待开始",
            font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#333333"
        )
        self.status_label.pack(anchor=tk.W)
        
        self.current_task_label = tk.Label(
            info_frame, text="当前任务：无",
            font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#333333"
        )
        self.current_task_label.pack(anchor=tk.W)
        
        self.queue_progress_label = tk.Label(
            info_frame, text="队列进度：第 0 条 / 共 0 条 | 已完成 0 条",
            font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#333333"
        )
        self.queue_progress_label.pack(anchor=tk.W)
        
        self.task_progress_label = tk.Label(
            info_frame, text="任务进度：工作段 0/0 | 休息次数 0/0",
            font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#333333"
        )
        self.task_progress_label.pack(anchor=tk.W)
        
        self.total_duration_label = tk.Label(
            info_frame, text="任务总时长：0分钟",
            font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#22C55E"
        )
        self.total_duration_label.pack(anchor=tk.W)
        
        self.post_break_label = tk.Label(
            info_frame, text="任务间休整总时间：0分钟",
            font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#22C55E"
        )
        self.post_break_label.pack(anchor=tk.W)
        
        self.total_schedule_label = tk.Label(
            info_frame, text="总安排时长：0分钟",
            font=("Microsoft YaHei", 10, "bold"), bg="#f0f0f0", fg="#22C55E"
        )
        self.total_schedule_label.pack(anchor=tk.W)
    
    def _create_control_buttons_area(self):
        """创建全局控制按钮区"""
        control_frame = tk.Frame(self.root, pady=10, bg="#f0f0f0")
        control_frame.pack(fill=tk.X)
        
        self.start_btn = tk.Button(
            control_frame, text="新的开始", width=14,
            command=self._start_queue, font=("Microsoft YaHei", 10, "bold"),
            bg="#333333", fg="white", activebackground="#555555"
        )
        self.start_btn.pack(side=tk.LEFT, padx=8)
        
        self.pause_btn = tk.Button(
            control_frame, text="暂停", width=10,
            command=self._toggle_pause, font=("Microsoft YaHei", 10),
            state=tk.DISABLED, bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
        )
        self.pause_btn.pack(side=tk.LEFT, padx=8)
        
        self.edit_btn = tk.Button(
            control_frame, text="修改", width=10,
            command=self._edit_tasks, font=("Microsoft YaHei", 10),
            bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
        )
        self.edit_btn.pack(side=tk.LEFT, padx=8)
        
        self.resume_task_btn = tk.Button(
            control_frame, text="继续任务", width=10,
            command=self._resume_task, font=("Microsoft YaHei", 10),
            state=tk.DISABLED, bg="#e0e0e0", fg="#999999", activebackground="#d0d0d0"
        )
        self.resume_task_btn.pack(side=tk.LEFT, padx=8)
        
        self.skip_btn = tk.Button(
            control_frame, text="跳过当前任务", width=12,
            command=self._skip_current_task, font=("Microsoft YaHei", 10),
            state=tk.DISABLED, bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
        )
        self.skip_btn.pack(side=tk.LEFT, padx=8)
        
        self.early_complete_btn = tk.Button(
            control_frame, text="提前完成", width=10,
            command=self._early_complete_current_task, font=("Microsoft YaHei", 10, "bold"),
            state=tk.DISABLED, bg="#333333", fg="white", activebackground="#555555"
        )
        self.early_complete_btn.pack(side=tk.LEFT, padx=8)
    
    def _create_footer(self):
        """创建底部作者信息"""
        footer_frame = tk.Frame(self.root, bg="#f0f0f0")
        footer_frame.pack(fill=tk.X, pady=(0, 5))
        
        footer_label = tk.Label(
            footer_frame, 
            text="V1.3 Mega_HUGO 2026.03", 
            font=("Microsoft YaHei", 8),
            bg="#f0f0f0", fg="#888888"
        )
        footer_label.pack()
    
    def _init_default_tasks(self):
        """初始化5条默认任务"""
        for _ in range(5):
            self._add_new_task()
    
    def _add_new_task(self, task_name=None):
        """添加新任务"""
        self.task_counter += 1
        task = TaskItem(self.task_counter)
        
        # 如果提供了任务名称，使用它
        if task_name:
            task.name = task_name
        
        # 获取当前行号
        row_idx = self.next_row
        self.next_row += 1
        
        # 第0列：拖拽手柄（⋮⋮）
        task.drag_handle = tk.Label(
            self.table_frame, text="⋮⋮",
            font=("Microsoft YaHei", 9), bg="#f0f0f0", fg="#888888",
            cursor="hand2"
        )
        task.drag_handle.grid(row=row_idx, column=0, padx=2, pady=2, sticky="ew")
        task.drag_handle.bind("<Button-1>", lambda e, t=task: self._on_drag_start(e, t))
        task.drag_handle.bind("<B1-Motion>", lambda e, t=task: self._on_drag_motion(e, t))
        task.drag_handle.bind("<ButtonRelease-1>", lambda e, t=task: self._on_drag_release(e, t))
        
        # 第1列：任务名称输入框
        task.name_var = tk.StringVar(value=task.name)
        name_entry = tk.Entry(self.table_frame, textvariable=task.name_var,
                             font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
        name_entry.grid(row=row_idx, column=1, padx=2, pady=2, sticky="ew")
        name_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
        
        # 第2列：任务预估时长输入框
        task.duration_var = tk.StringVar(value=str(task.duration))
        duration_entry = tk.Entry(self.table_frame, textvariable=task.duration_var,
                                 font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
        duration_entry.grid(row=row_idx, column=2, padx=2, pady=2, sticky="ew")
        duration_entry.bind("<KeyRelease>", lambda e: self._update_total_duration())
        duration_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
        
        # 第3列：休息次数输入框
        task.break_count_var = tk.StringVar(value=str(task.break_count))
        break_count_entry = tk.Entry(self.table_frame, textvariable=task.break_count_var,
                                    font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
        break_count_entry.grid(row=row_idx, column=3, padx=2, pady=2, sticky="ew")
        break_count_entry.bind("<KeyRelease>", lambda e: self._update_total_duration())
        break_count_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
        
        # 第4列：休息时长输入框
        task.break_duration_var = tk.StringVar(value=str(task.break_duration))
        break_duration_entry = tk.Entry(self.table_frame, textvariable=task.break_duration_var,
                                       font=("Microsoft YaHei", 9), bg="white", fg="#333333", insertbackground="#333333")
        break_duration_entry.grid(row=row_idx, column=4, padx=2, pady=2, sticky="ew")
        break_duration_entry.bind("<KeyRelease>", lambda e: self._update_total_duration())
        break_duration_entry.bind("<Button-1>", lambda e, t=task: self._on_task_select(t))
        
        # 第5列：删除按钮
        task.delete_btn = tk.Button(
            self.table_frame, text="删除",
            command=lambda t=task: self._delete_task(t),
            font=("Microsoft YaHei", 9), bg="#e0e0e0", fg="#333333", activebackground="#d0d0d0"
        )
        task.delete_btn.grid(row=row_idx, column=5, padx=2, pady=2, sticky="w")
        
        task.row_idx = row_idx  # 记录行号
        self.tasks.append(task)
        
        # 更新删除按钮状态
        self._update_delete_buttons()
        
        # 更新总时长
        self._update_total_duration()
        
        # 滚动到底部
        self.root.after(100, self._scroll_to_bottom)
        
        return task
    
    def _scroll_to_bottom(self):
        """滚动到任务列表底部"""
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)
    
    def _delete_task(self, task):
        """删除指定任务"""
        if len(self.tasks) <= 1:
            return
        
        # 隐藏该行的所有控件
        for col in range(6):
            widget = self.table_frame.grid_slaves(row=task.row_idx, column=col)
            if widget:
                widget[0].grid_forget()
        
        self.tasks.remove(task)
        self._update_delete_buttons()
        self._update_total_duration()
    
    def _clear_all_tasks(self):
        """清空所有任务（保留1条）"""
        # 删除所有任务UI（保留表头第0行）
        for task in self.tasks:
            for col in range(6):
                widget = self.table_frame.grid_slaves(row=task.row_idx, column=col)
                if widget:
                    widget[0].destroy()
        
        self.tasks.clear()
        self.task_counter = 0
        self.next_row = 1  # 重置行号（表头占第0行）
        
        # 添加1条新任务
        self._add_new_task()
    
    def _update_delete_buttons(self):
        """更新删除按钮状态"""
        can_delete = len(self.tasks) > 1
        for task in self.tasks:
            if can_delete:
                task.delete_btn.config(state=tk.NORMAL)
            else:
                task.delete_btn.config(state=tk.DISABLED)
    
    def _update_total_duration(self):
        """更新任务总时长、任务间休整总时间、总安排时长"""
        task_total_minutes = 0  # 任务总时长（包含任务内休息）
        post_break_total = 0    # 任务间休整总时间（每个任务完成后的5分钟）
        
        # 统计未完成的任务数量（用于计算任务间休整）
        incomplete_task_count = 0
        
        for task in self.tasks:
            # 跳过已完成或已跳过的任务
            if task.completed or task.skipped:
                continue
            
            incomplete_task_count += 1
            
            # 获取任务参数（从输入框读取最新值）
            try:
                duration = int(task.duration_var.get()) if task.duration_var else task.duration
            except:
                duration = task.duration
            
            try:
                break_count = int(task.break_count_var.get()) if task.break_count_var else task.break_count
            except:
                break_count = task.break_count
            
            try:
                break_duration = int(task.break_duration_var.get()) if task.break_duration_var else task.break_duration
            except:
                break_duration = task.break_duration
            
            # 任务总时长 = 预估时长 + 休息次数 * 休息时长
            task_total_minutes += duration + break_count * break_duration
        
        # 任务间休整总时间 = (未完成任务数 - 1) * 用户设置的休整时间
        # 最后一个任务完成后不需要休整
        if incomplete_task_count > 1:
            try:
                post_break_minutes = int(self.post_break_var.get())
            except:
                post_break_minutes = 5
            post_break_total = (incomplete_task_count - 1) * post_break_minutes
        
        # 总安排时长 = 任务总时长 + 任务间休整总时间
        total_schedule_minutes = task_total_minutes + post_break_total
        
        # 格式化显示 - 任务总时长
        if task_total_minutes >= 60:
            hours = task_total_minutes // 60
            mins = task_total_minutes % 60
            self.total_duration_label.config(text=f"任务总时长：{hours}小时{mins}分钟（共{task_total_minutes}分钟）")
        else:
            self.total_duration_label.config(text=f"任务总时长：{task_total_minutes}分钟")
        
        # 格式化显示 - 任务间休整总时间
        self.post_break_label.config(text=f"任务间休整总时间：{post_break_total}分钟")
        
        # 格式化显示 - 总安排时长
        if total_schedule_minutes >= 60:
            hours = total_schedule_minutes // 60
            mins = total_schedule_minutes % 60
            self.total_schedule_label.config(text=f"总安排时长：{hours}小时{mins}分钟（共{total_schedule_minutes}分钟）")
        else:
            self.total_schedule_label.config(text=f"总安排时长：{total_schedule_minutes}分钟")
    
    def _highlight_current_task(self):
        """高亮当前正在计时的任务行"""
        for i, task in enumerate(self.tasks):
            row_idx = task.row_idx
            
            # 获取该行的所有控件（按列顺序）
            widgets_by_col = {}
            for col in range(6):
                widget = self.table_frame.grid_slaves(row=row_idx, column=col)
                if widget:
                    widgets_by_col[col] = widget[0]
            
            # 判断任务状态
            is_current = (i == self.current_task_index and self.is_running and not self.is_paused)
            is_selected = (self.selected_task_index == i and not is_current and not task.completed and not task.skipped)
            
            if is_current:
                # 当前任务：深绿色高亮框，隐藏拖拽手柄
                for col, widget in widgets_by_col.items():
                    if isinstance(widget, tk.Entry):
                        widget.config(bg="#BBF7D0", fg="#166534")
                    elif isinstance(widget, tk.Label):
                        if col == 0:  # 拖拽手柄列
                            widget.config(text="⋮⋮", bg="#BBF7D0", fg="#CCCCCC")  # 浅色表示不可拖拽
                        else:
                            widget.config(bg="#BBF7D0", fg="#166534")
                    elif isinstance(widget, tk.Button):
                        widget.config(bg="#A7F3D0", fg="#166534")
            elif task.completed or task.skipped:
                # 已完成/跳过任务：浅绿色，隐藏拖拽手柄
                for col, widget in widgets_by_col.items():
                    if isinstance(widget, tk.Entry):
                        widget.config(bg="#DCFCE7", fg="#166534")
                    elif isinstance(widget, tk.Label):
                        if col == 0:  # 拖拽手柄列
                            widget.config(text="", bg="#DCFCE7", fg="#CCCCCC")  # 隐藏拖拽手柄
                        else:
                            widget.config(bg="#DCFCE7", fg="#166534")
                    elif isinstance(widget, tk.Button):
                        widget.config(bg="#BBF7D0", fg="#166534")
            elif is_selected:
                # 选中状态：蓝色边框高亮
                for col, widget in widgets_by_col.items():
                    if isinstance(widget, tk.Entry):
                        widget.config(bg="#DBEAFE", fg="#1E40AF")
                    elif isinstance(widget, tk.Label):
                        if col == 0:  # 拖拽手柄列
                            widget.config(text="⋮⋮", bg="#DBEAFE", fg="#1E40AF")
                        else:
                            widget.config(bg="#DBEAFE", fg="#1E40AF")
                    elif isinstance(widget, tk.Button):
                        widget.config(bg="#BFDBFE", fg="#1E40AF")
            else:
                # 普通任务：灰色，显示拖拽手柄
                for col, widget in widgets_by_col.items():
                    if isinstance(widget, tk.Entry):
                        widget.config(bg="white", fg="#333333")
                    elif isinstance(widget, tk.Label):
                        if col == 0:  # 拖拽手柄列
                            widget.config(text="⋮⋮", bg="#f0f0f0", fg="#888888")
                        else:
                            widget.config(bg="#f0f0f0", fg="#333333")
                    elif isinstance(widget, tk.Button):
                        widget.config(bg="#e0e0e0", fg="#333333")
    
    def _validate_tasks(self):
        """验证所有任务参数"""
        for i, task in enumerate(self.tasks):
            # 跳过已完成或已跳过的任务（它们的预估时长可能为0）
            if task.completed or task.skipped:
                continue
            
            # 获取任务名称
            task.name = task.name_var.get().strip() or f"任务{task.task_id}"
            
            # 验证预估时长
            try:
                duration = int(task.duration_var.get())
                if duration < 1:
                    return False, f"任务{i+1}的预估时长必须≥1分钟"
                task.duration = duration
            except ValueError:
                return False, f"任务{i+1}的预估时长必须是整数"
            
            # 验证休息次数
            try:
                break_count = int(task.break_count_var.get())
                if break_count < 0:
                    return False, f"任务{i+1}的休息次数必须≥0"
                task.break_count = break_count
            except ValueError:
                return False, f"任务{i+1}的休息次数必须是整数"
            
            # 验证休息时长
            try:
                break_duration = int(task.break_duration_var.get())
                if break_duration < 1:
                    return False, f"任务{i+1}的休息时长必须≥1分钟"
                task.break_duration = break_duration
            except ValueError:
                return False, f"任务{i+1}的休息时长必须是整数"
        
        # 更新总时长显示
        self._update_total_duration()
        
        return True, ""
    
    def _lock_ui(self, locked):
        """锁定/解锁UI"""
        state = tk.DISABLED if locked else tk.NORMAL
        
        for task in self.tasks:
            task.name_var.set(task.name)
            # 使用 grid_slaves 获取该行的所有控件
            for col in range(6):
                widget = self.table_frame.grid_slaves(row=task.row_idx, column=col)
                if widget:
                    w = widget[0]
                    if isinstance(w, tk.Entry):
                        w.config(state=state)
            task.delete_btn.config(state=state)
        
        self.add_task_btn.config(state=state)
        self.clear_all_btn.config(state=state)
        self.import_btn.config(state=state)
        
        # 跳转按钮：始终可用（任何状态都可以跳转到选定的任务）
        self.jump_to_task_btn.config(state=tk.NORMAL)
    
    def _apply_ui_lock_state(self):
        """应用当前UI锁定状态（用于刷新表格后恢复锁定）"""
        self._lock_ui(self.is_running)
    
    def _update_selection_highlight(self):
        """更新选中任务的高亮显示"""
        for i, task in enumerate(self.tasks):
            if i == self.selected_task_index:
                # 选中状态：蓝色高亮
                try:
                    task.drag_handle.config(bg="#DBEAFE", fg="#1E40AF")
                except:
                    pass
            # 其他状态由 _highlight_current_task 处理
    
    def _show_import_dialog(self):
        """显示批量导入任务弹窗"""
        # 创建置顶窗口
        top = tk.Toplevel(self.root)
        top.title("批量导入任务")
        top.attributes("-topmost", True)
        top.resizable(False, False)
        
        # 居中显示
        top.update_idletasks()
        width = 450
        height = 400
        x = (top.winfo_screenwidth() // 2) - (width // 2)
        y = (top.winfo_screenheight() // 2) - (height // 2)
        top.geometry(f"{width}x{height}+{x}+{y}")
        
        # 提示信息
        tip_label = tk.Label(
            top, text="请输入或粘贴任务文本（支持「- [ ]」前缀格式或纯文本）：", 
            font=("Microsoft YaHei", 10), wraplength=400
        )
        tip_label.pack(pady=10)
        
        # 示例
        example_label = tk.Label(
            top, text="示例：- [ ]  任务名称  或直接输入  任务名称", 
            font=("Microsoft YaHei", 9), fg="#888888"
        )
        example_label.pack()
        
        # 文本输入框
        text_frame = tk.Frame(top)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        text_input = tk.Text(text_frame, font=("Microsoft YaHei", 10), wrap=tk.WORD, height=10)
        text_input.pack(fill=tk.BOTH, expand=True)
        
        # 操作按钮区（两个并列按钮）
        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=10)
        
        def do_replace():
            """替换所有任务"""
            text = text_input.get("1.0", tk.END).strip()
            self._import_tasks_from_text(text, mode='replace')
            top.destroy()
        
        def do_append():
            """新增到现有任务"""
            text = text_input.get("1.0", tk.END).strip()
            self._import_tasks_from_text(text, mode='append')
            top.destroy()
        
        replace_btn = tk.Button(
            btn_frame, text="替换所有任务", width=14,
            command=do_replace, font=("Microsoft YaHei", 9),
            bg="#333333", fg="white", activebackground="#555555"
        )
        replace_btn.pack(side=tk.LEFT, padx=10)
        
        append_btn = tk.Button(
            btn_frame, text="新增到现有任务", width=14,
            command=do_append, font=("Microsoft YaHei", 9),
            bg="#333333", fg="white", activebackground="#555555"
        )
        append_btn.pack(side=tk.LEFT, padx=10)
        
        # 取消按钮
        cancel_btn = tk.Button(
            top, text="取消", width=10,
            command=top.destroy, font=("Microsoft YaHei", 9),
            bg="#e0e0e0", fg="#333333"
        )
        cancel_btn.pack(pady=5)
        
        # 聚焦
        top.focus_set()
        text_input.focus_set()
    
    def _import_tasks_from_text(self, text, mode='append'):
        """从文本批量导入任务（兼容纯文本和带前缀格式）
        
        Args:
            text: 输入的文本内容
            mode: 'replace' 替换所有任务, 'append' 新增到现有任务
        """
        if not text:
            return
        
        # 解析文本获取任务列表
        lines = text.split('\n')
        new_tasks = []
        
        for line in lines:
            # 尝试移除「- [ ]」前缀（无论前缀与文本间是否有空格）
            cleaned = re.sub(r'^\s*-\s*\[\s*\]\s*', '', line).strip()
            
            # 忽略空行
            if cleaned:
                new_tasks.append(cleaned)
        
        if not new_tasks:
            return
        
        # 根据模式处理
        if mode == 'replace':
            # 清除当天的日志记录（旧任务已被替换，不应计入统计）
            self.logger.clear_today_log()
            # 清空现有任务的所有控件
            for task in self.tasks:
                for col in range(6):
                    widget = self.table_frame.grid_slaves(row=task.row_idx, column=col)
                    if widget:
                        widget[0].destroy()
            self.tasks.clear()
            self.task_counter = 0
            # 重置行号计数器
            self.next_row = 1
        
        # 添加新任务
        for task_name in new_tasks:
            self._add_new_task(task_name=task_name)
        
        # 滚动到底部显示新导入的任务
        self.root.after(100, self._scroll_to_bottom)
    
    def _start_queue(self):
        """开始执行队列"""
        # 验证任务参数
        valid, error_msg = self._validate_tasks()
        if not valid:
            self._show_toast_notification("参数错误", error_msg)
            return
        
        # 锁定UI
        self._lock_ui(True)
        
        # 初始化状态
        self.is_running = True
        self.is_paused = False
        # 始终从第一行任务开始（确保拖拽排序后执行正确的任务）
        self.current_task_index = 0
        self.completed_tasks = 0
        # 重置所有任务完成状态
        for task in self.tasks:
            task.completed = False
            task.skipped = False
        
        # 更新按钮状态
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL, text="暂停")
        self.skip_btn.config(state=tk.NORMAL)
        self.early_complete_btn.config(state=tk.NORMAL)
        self.resume_task_btn.config(state=tk.DISABLED, fg="#999999")
        
        # 重置修改模式状态
        self.is_edit_mode = False
        
        # 开始当前任务
        self._start_current_task()
    
    def _resume_task(self):
        """继续任务（使用新的预估时长和休息次数重新计算分割）"""
        if not self.is_edit_mode:
            return
        
        # 验证任务参数
        valid, error_msg = self._validate_tasks()
        if not valid:
            self._show_toast_notification("参数错误", error_msg)
            return
        
        # 锁定UI
        self._lock_ui(True)
        
        # 获取当前任务
        task = self.tasks[self.saved_task_index]
        
        # 从输入框读取新的预估时长和休息次数
        try:
            new_duration = int(task.duration_var.get())
            if new_duration <= 0:
                new_duration = 1
        except (ValueError, AttributeError):
            new_duration = 1
        
        try:
            new_break_count = int(task.break_count_var.get())
            if new_break_count < 0:
                new_break_count = 0
        except (ValueError, AttributeError):
            new_break_count = 0
        
        # 更新任务的预估时长和休息次数
        task.duration = new_duration
        task.break_count = new_break_count
        
        # 如果之前是任务结束休息状态，重置任务完成状态
        if self.is_post_task_break:
            task.completed = False
            task.early_completed = False
            self.completed_tasks -= 1  # 减少已完成计数
        
        # 重置任务结束休息状态
        self.is_post_task_break = False
        
        # 重新开始当前任务（从第一个工作段开始，按照新的参数分割）
        self.is_running = True
        self.is_paused = False
        self.is_edit_mode = False
        self.current_task_index = self.saved_task_index
        self.current_work_segment = 0
        self.current_break_count = 0
        self.is_in_break = False
        
        # 记录任务开始时间
        task.start_time = datetime.now()
        
        # 计算单段工作时长
        total_segments = task.calculate_work_segments()
        single_work_minutes = new_duration / total_segments
        self.remaining_seconds = int(single_work_minutes * 60)
        
        # 更新状态显示
        self.status_label.config(text="当前状态：工作中")
        self.current_task_label.config(text=f"当前任务：【任务{task.task_id}】{task.name}")
        self._update_progress_display()
        
        # 更新按钮状态
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL, text="暂停")
        self.skip_btn.config(state=tk.NORMAL)
        self.early_complete_btn.config(state=tk.NORMAL)
        self.resume_task_btn.config(state=tk.DISABLED, fg="#999999")
        
        # 高亮当前任务
        self._highlight_current_task()
        
        # 开始倒计时（从第一段工作开始）
        self.current_work_segment = 1
        self._run_timer()
    
    def _start_current_task(self):
        """开始当前任务"""
        if self.current_task_index >= len(self.tasks):
            self._queue_completed()
            return
        
        task = self.tasks[self.current_task_index]
        self.current_work_segment = 0
        self.current_break_count = 0
        self.is_in_break = False
        
        # 记录任务开始时间
        task.start_time = datetime.now()
        
        # 高亮当前任务
        self._highlight_current_task()
        
        # 开始第一个工作段
        self._start_work_segment()
    
    def _start_work_segment(self):
        """开始工作段"""
        task = self.tasks[self.current_task_index]
        self.is_in_break = False
        self.current_work_segment += 1
        
        # 计算单段工作时长（秒）
        single_work_minutes = task.calculate_single_work_duration()
        self.remaining_seconds = int(single_work_minutes * 60)
        
        # 更新状态显示
        self.status_label.config(text="当前状态：工作中")
        self.current_task_label.config(text=f"当前任务：【任务{task.task_id}】{task.name}")
        self._update_progress_display()
        
        # 开始倒计时
        self._run_timer()
    
    def _start_break(self):
        """开始休息"""
        task = self.tasks[self.current_task_index]
        self.is_in_break = True
        self.current_break_count += 1
        
        # 休息时长（秒）
        self.remaining_seconds = task.break_duration * 60
        
        # 更新状态显示
        self.status_label.config(text="当前状态：休息中")
        self._update_progress_display()
        
        # 开始倒计时
        self._run_timer()
    
    def _run_timer(self):
        """运行倒计时"""
        if not self.is_running:
            return
        
        if self.is_paused:
            return
        
        # 更新显示
        self._update_timer_display()
        
        # 高亮当前任务
        self._highlight_current_task()
        
        if self.remaining_seconds > 0:
            self.remaining_seconds -= 1
            self.timer_id = self.root.after(1000, self._run_timer)
        else:
            # 时间到
            self._on_time_up()
    
    def _update_timer_display(self):
        """更新倒计时显示"""
        minutes = self.remaining_seconds // 60
        seconds = self.remaining_seconds % 60
        self.timer_label.config(text=f"{minutes:02d}:{seconds:02d}")
    
    def _update_progress_display(self):
        """更新进度显示"""
        task = self.tasks[self.current_task_index]
        total_segments = task.calculate_work_segments()
        
        # 队列进度
        self.queue_progress_label.config(
            text=f"队列进度：第 {self.current_task_index + 1} 条 / 共 {len(self.tasks)} 条 | 已完成 {self.completed_tasks} 条"
        )
        
        # 任务进度
        self.task_progress_label.config(
            text=f"任务进度：工作段 {self.current_work_segment}/{total_segments} | 休息次数 {self.current_break_count}/{task.break_count}"
        )
    
    def _on_time_up(self):
        """时间到处理"""
        task = self.tasks[self.current_task_index]
        total_segments = task.calculate_work_segments()
        
        if self.is_post_task_break:
            # 任务结束休息完毕，检查是否需要跳转到下一个任务
            self._on_post_task_break_end()
            return
        
        if self.is_in_break:
            # 休息结束，继续工作（Toast 自带提示音）
            self._show_toast_notification("休息结束", "休息结束，继续完成任务", silent=False)
            
            if self.current_work_segment < total_segments:
                # 还有工作段
                self._start_work_segment()
            else:
                # 任务完成
                self._task_completed()
        else:
            # 工作段结束
            if self.current_break_count < task.break_count:
                # 需要休息（静音通知）
                self._show_toast_notification("休息时间到", f"第{self.current_break_count + 1}次休息时间到，休息时长{task.break_duration}分钟", silent=True)
                self._start_break()
            else:
                # 任务完成
                self._task_completed()
    
    def _start_post_task_break(self):
        """开始任务结束后的休息（时间由用户设置）"""
        self.is_post_task_break = True
        
        # 获取用户设置的休整时间（分钟）
        try:
            post_break_minutes = int(self.post_break_var.get())
        except:
            post_break_minutes = 5
        
        self.remaining_seconds = post_break_minutes * 60
        
        # 更新状态显示
        self.status_label.config(text="当前状态：任务结束休息中")
        self.task_progress_label.config(text=f"任务进度：任务已完成，{post_break_minutes}分钟休息整顿")
        
        # Toast提示
        task = self.tasks[self.current_task_index]
        self._show_toast_notification("任务已完成", f"【任务{task.task_id}】已完成，进入{post_break_minutes}分钟休息整顿", silent=False)
        
        # 开始倒计时
        self._run_timer()
    
    def _on_post_task_break_end(self):
        """任务结束休息结束后的处理"""
        self.is_post_task_break = False
        
        # 获取用户设置的休整时间（分钟）
        try:
            post_break_minutes = int(self.post_break_var.get())
        except:
            post_break_minutes = 5
        
        self._show_toast_notification("休息结束", f"{post_break_minutes}分钟休息结束", silent=False)
        
        # 检查当前任务的预估时长是否为0
        task = self.tasks[self.current_task_index]
        try:
            duration = int(task.duration_var.get())
        except (ValueError, AttributeError):
            duration = 0
        
        if duration <= 0:
            # 预估时长为0，跳转到下一个任务
            self.current_task_index += 1
            if self.current_task_index < len(self.tasks):
                self._start_current_task()
            else:
                self._queue_completed()
        else:
            # 预估时长不为0，重新执行当前任务
            self.is_edit_mode = False
            self._start_current_task()
    
    def _task_completed(self):
        """当前任务完成（正常完成）- 触发5分钟休息"""
        task = self.tasks[self.current_task_index]
        self.completed_tasks += 1
        task.completed = True
        
        # 记录日志
        self.logger.log_task(task, "normal_complete")
        
        # 更新任务显示为已完成（打勾）
        self._highlight_current_task()
        
        # 将预估时长设为0（因为任务已完成）
        if task.duration_var:
            task.duration_var.set("0")
        
        # 开始5分钟休息（不自动跳转到下一个任务）
        self._start_post_task_break()
    
    def _queue_completed(self):
        """队列全部完成"""
        self.is_running = False
        self.is_edit_mode = False
        
        self._show_toast_notification("队列完成", f"恭喜！所有任务已全部完成！本次共完成{self.completed_tasks}条任务")
        self._play_sound()
        
        # 解锁UI
        self._lock_ui(False)
        
        # 更新按钮状态
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text="暂停")
        self.skip_btn.config(state=tk.DISABLED)
        self.early_complete_btn.config(state=tk.DISABLED)
        self.resume_task_btn.config(state=tk.DISABLED, fg="#999999")
        
        # 更新状态显示
        self.status_label.config(text="当前状态：已完成")
        self.timer_label.config(text="00:00")
    
    def _toggle_pause(self):
        """切换暂停/继续"""
        if not self.is_running:
            return
        
        if self.is_paused:
            # 继续
            self.is_paused = False
            self.pause_btn.config(text="暂停")
            self._run_timer()
        else:
            # 暂停
            self.is_paused = True
            self.pause_btn.config(text="继续")
            if self.timer_id:
                self.root.after_cancel(self.timer_id)
                self.timer_id = None
            # 取消高亮
            self._highlight_current_task()
    
    def _edit_tasks(self):
        """修改任务（保留所有已配置数据，进入编辑模式）"""
        # 保存当前任务进度（用于"继续任务"功能）
        was_running = self.is_running or self.is_paused or self.is_post_task_break
        if was_running:
            self.is_edit_mode = True
            self.saved_task_index = self.current_task_index
            self.saved_work_segment = self.current_work_segment
            self.saved_break_count = self.current_break_count
            self.saved_is_in_break = self.is_in_break
            
            # 计算任务的整体剩余时长并填入预估时长输入框
            task = self.tasks[self.current_task_index]
            
            # 如果是任务结束休息期间，预估时长已经是0了，不需要修改
            if self.is_post_task_break:
                # 任务结束休息期间，预估时长保持为0（已在_task_completed中设置）
                pass
            elif not self.is_in_break:
                # 工作中：计算整体剩余时长
                # 当前段剩余时间
                current_segment_remaining = (self.remaining_seconds + 59) // 60  # 向上取整
                
                # 计算剩余工作段数（包括当前段）
                total_segments = task.calculate_work_segments()
                remaining_segments = total_segments - self.current_work_segment + 1
                
                # 单段原始时长
                original_single_segment = task.duration / total_segments
                
                # 整体剩余时长 = 当前段剩余 + (剩余段数-1) × 单段原始时长
                total_remaining = current_segment_remaining + (remaining_segments - 1) * original_single_segment
                
                if task.duration_var:
                    task.duration_var.set(str(int(total_remaining)))
            else:
                # 休息中：剩余工作段数 × 单段时长
                total_segments = task.calculate_work_segments()
                remaining_segments = total_segments - self.current_work_segment  # 休息结束后进入下一段工作
                original_single_segment = task.duration / total_segments
                total_remaining = remaining_segments * original_single_segment
                
                if task.duration_var:
                    task.duration_var.set(str(int(total_remaining)))
        
        # 停止计时但不重置进度
        self.is_running = False
        self.is_paused = False
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        
        # 解锁UI（保留所有任务数据和当前进度）
        self._lock_ui(False)
        
        # 更新按钮状态
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text="暂停")
        self.skip_btn.config(state=tk.DISABLED)
        self.early_complete_btn.config(state=tk.DISABLED)
        
        # 如果之前正在执行任务，启用"继续任务"按钮
        if was_running:
            self.resume_task_btn.config(state=tk.NORMAL, fg="#333333")
        else:
            self.resume_task_btn.config(state=tk.DISABLED, fg="#999999")
        
        # 状态显示保持当前进度
        self.status_label.config(text="当前状态：待开始")
    
    def _skip_current_task(self):
        """跳过当前任务（记录到日志但不计入时长统计，不触发5分钟休息）"""
        # 停止计时
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        
        # 重置任务结束休息状态
        self.is_post_task_break = False
        self.is_in_break = False
        
        # 标记当前任务为已跳过
        task = self.tasks[self.current_task_index]
        task.skipped = True
        
        # 记录日志（completion_type=skipped，不计入时长统计）
        self.logger.log_task(task, "skipped")
        
        # 更新显示
        self._highlight_current_task()
        
        # 直接跳到下一个任务（不触发5分钟休息）
        self.current_task_index += 1
        self.is_running = True  # 保持运行状态以启动下一个任务
        self.is_paused = False
        self.is_edit_mode = False
        
        if self.current_task_index < len(self.tasks):
            self._start_current_task()
        else:
            self._queue_completed()
    
    def _early_complete_current_task(self):
        """提前完成当前任务 - 在休息整顿期间直接跳到下一个任务，否则触发5分钟休息"""
        # 如果是在任务结束休息整顿期间点击，直接跳到下一个任务
        if self.is_post_task_break:
            # 停止计时
            if self.timer_id:
                self.root.after_cancel(self.timer_id)
                self.timer_id = None
            
            # 重置休息状态
            self.is_post_task_break = False
            self.is_in_break = False
            
            # 跳到下一个任务
            self.current_task_index += 1
            self.is_running = True
            self.is_paused = False
            self.is_edit_mode = False
            
            if self.current_task_index < len(self.tasks):
                self._show_toast_notification("跳过休息", "跳过休息整顿，开始下一个任务")
                self._start_current_task()
            else:
                self._queue_completed()
            return
        
        if not self.is_running:
            return
        
        # 停止计时
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        
        # 计算实际耗时（秒）
        task = self.tasks[self.current_task_index]
        now = datetime.now()
        if task.start_time:
            actual_seconds = int((now - task.start_time).total_seconds())
            task.actual_duration_seconds = actual_seconds
        
        # 标记为提前完成
        task.early_completed = True
        task.completed = True
        self.completed_tasks += 1
        
        # 记录日志
        self.logger.log_task(task, "early_complete")
        
        # 更新显示
        self._highlight_current_task()
        
        # 将预估时长设为0（因为任务已完成）
        if task.duration_var:
            task.duration_var.set("0")
        
        # 显示通知
        actual_min = task.actual_duration_seconds // 60
        actual_sec = task.actual_duration_seconds % 60
        self._show_toast_notification("提前完成", f"【任务{task.task_id}】已提前完成！实际耗时：{actual_min}分{actual_sec}秒")
        
        # 开始5分钟休息（不自动跳转到下一个任务）
        self._start_post_task_break()
    
    def _jump_to_selected_task(self):
        """跳转到选定的任务"""
        # 检查是否选中了目标任务
        if self.selected_task_index is None:
            messagebox.showinfo("提示", "请先在任务队列中选择要跳转到的目标任务")
            return
        
        target_task = self.tasks[self.selected_task_index]
        
        # 检查目标任务是否有效
        if target_task.completed or target_task.skipped:
            messagebox.showinfo("提示", "请选择未开始的任务进行跳转")
            return
        
        if self.selected_task_index == self.current_task_index and self.is_running:
            messagebox.showinfo("提示", "目标任务正在执行中，无需跳转")
            return
        
        # 执行跳转逻辑
        self._execute_jump(self.selected_task_index)
    
    def _execute_jump(self, target_index):
        """执行跳转的核心逻辑"""
        if self.is_running:
            # 正在计时时的跳转：终止当前任务，记录实际时长，把目标任务移到第一位
            if self.timer_id:
                self.root.after_cancel(self.timer_id)
                self.timer_id = None
            
            # 获取当前任务和目标任务
            current_task = self.tasks[self.current_task_index]
            target_task = self.tasks[target_index]
            now = datetime.now()
            
            # 计算当前任务的已消耗时间
            actual_focus_minutes = 0
            if current_task.start_time:
                elapsed_seconds = int((now - current_task.start_time).total_seconds())
                elapsed_minutes = elapsed_seconds / 60
                used_break_minutes = self.current_break_count * current_task.break_duration
                actual_focus_minutes = max(0, elapsed_minutes - used_break_minutes)
            
            # 计算剩余数据
            remaining_duration = current_task.duration - actual_focus_minutes
            remaining_breaks = current_task.break_count - self.current_break_count
            
            # 记录当前任务的跳转终止数据
            current_task.jump_terminated = True
            current_task.actual_focus_minutes = actual_focus_minutes
            current_task.actual_breaks_taken = self.current_break_count
            current_task.remaining_duration = remaining_duration
            current_task.remaining_breaks = remaining_breaks
            current_task.actual_duration_seconds = int(actual_focus_minutes * 60)
            
            # 记录日志
            self.logger.log_task(current_task, "jump_terminated")
            
            # 更新当前任务的参数（剩余部分）
            current_task.duration = max(1, int(remaining_duration))
            current_task.break_count = max(0, remaining_breaks)
            current_task.jump_terminated = False
            
            # 把目标任务移到第一位，当前任务移到第二位
            # 先移除目标任务
            target_task = self.tasks.pop(target_index)
            # 如果当前任务在目标任务前面，索引需要调整
            if target_index < self.current_task_index:
                self.current_task_index -= 1
            # 移除当前任务
            current_task = self.tasks.pop(self.current_task_index)
            
            # 插入到列表开头：目标任务第一，当前任务第二
            self.tasks.insert(0, target_task)
            self.tasks.insert(1, current_task)
            
            # 清除选中状态
            self.selected_task_index = None
            
            # 设置当前任务索引为0（第一位目标任务）
            self.current_task_index = 0
            self.current_work_segment = 0
            self.current_break_count = 0
            self.is_in_break = False
            
            # 刷新表格显示
            self._refresh_task_table()
            
            # 启动目标任务
            self._start_current_task()
            
            # 显示通知
            self._show_toast_notification("跳转成功", f"已将【{target_task.name}】提升到第一位并开始执行", silent=True)
        else:
            # 未计时时的跳转：把目标任务移到第一位，然后开始执行
            target_task = self.tasks[target_index]
            
            # 检查目标任务是否已完成/跳过
            if target_task.completed or target_task.skipped:
                messagebox.showinfo("提示", "目标任务已完成或已跳过，请选择其他任务")
                return
            
            # 把目标任务移到第一位
            target_task = self.tasks.pop(target_index)
            self.tasks.insert(0, target_task)
            
            # 清除选中状态
            self.selected_task_index = None
            
            # 设置当前任务索引为0（第一位）
            self.current_task_index = 0
            
            # 刷新表格显示
            self._refresh_task_table()
            
            # 开始执行目标任务
            self._start_queue()
            
            # 显示通知
            self._show_toast_notification("跳转成功", f"已将【{target_task.name}】提升到第一位并开始执行", silent=True)
    
    def _show_toast_notification(self, title, message, silent=True):
        """显示 Windows 系统通知栏 Toast 消息
        
        Args:
            title: 通知标题
            message: 通知内容
            silent: True=静音, False=使用 Toast 默认提示音
        """
        try:
            audio_tag = '<audio silent="true"/>' if silent else ''
            # 使用 PowerShell 调用 Windows Runtime API 发送 Toast 通知
            ps_script = f'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

$template = @"
<toast duration="short">
    <visual>
        <binding template="ToastText02">
            <text id="1">{title}</text>
            <text id="2">{message}</text>
        </binding>
    </visual>
    {audio_tag}
</toast>
"@

$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml($template)
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("多任务队列计时闹钟").Show($toast)
'''
            subprocess.run(
                ['powershell', '-Command', ps_script],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
        except Exception:
            pass
    
    def _play_sound(self):
        """播放 Windows 默认通知音"""
        try:
            # 使用 Windows 系统默认通知音
            winsound.MessageBeep(winsound.MB_OK)
        except:
            pass
    
    # ==================== 日志和统计功能 ====================
    
    def _show_calendar_window(self):
        """显示日历统计窗口"""
        # 创建置顶窗口
        top = tk.Toplevel(self.root)
        top.title("📅 日志统计")
        top.attributes("-topmost", True)
        top.resizable(False, False)
        
        # 当前显示的年月
        now = datetime.now()
        current_year = now.year
        current_month = now.month
        
        # 主框架
        main_frame = tk.Frame(top, bg="#f0f0f0", padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== 月份导航区 ==========
        nav_frame = tk.Frame(main_frame, bg="#f0f0f0")
        nav_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 月份变量
        year_var = tk.IntVar(value=current_year)
        month_var = tk.IntVar(value=current_month)
        
        def update_calendar():
            year = year_var.get()
            month = month_var.get()
            _render_calendar(year, month)
            _update_statistics(year, month)
        
        def prev_month():
            year = year_var.get()
            month = month_var.get()
            if month == 1:
                year_var.set(year - 1)
                month_var.set(12)
            else:
                month_var.set(month - 1)
            update_calendar()
        
        def next_month():
            year = year_var.get()
            month = month_var.get()
            if month == 12:
                year_var.set(year + 1)
                month_var.set(1)
            else:
                month_var.set(month + 1)
            update_calendar()
        
        # 导航按钮
        prev_btn = tk.Button(
            nav_frame, text="◀", width=3, command=prev_month,
            font=("Microsoft YaHei", 10), bg="#e0e0e0", fg="#333333"
        )
        prev_btn.pack(side=tk.LEFT, padx=5)
        
        # 年月显示
        month_label = tk.Label(
            nav_frame, text=f"{current_year}年{current_month}月",
            font=("Microsoft YaHei", 14, "bold"), bg="#f0f0f0", fg="#333333",
            width=12
        )
        month_label.pack(side=tk.LEFT, padx=10)
        
        next_btn = tk.Button(
            nav_frame, text="▶", width=3, command=next_month,
            font=("Microsoft YaHei", 10), bg="#e0e0e0", fg="#333333"
        )
        next_btn.pack(side=tk.LEFT, padx=5)
        
        # 导出按钮
        export_btn = tk.Button(
            nav_frame, text="导出CSV", width=8,
            command=lambda: _export_csv(year_var.get(), month_var.get()),
            font=("Microsoft YaHei", 9), bg="#333333", fg="white"
        )
        export_btn.pack(side=tk.RIGHT, padx=5)
        
        # ========== 日历网格区 ==========
        calendar_frame = tk.Frame(main_frame, bg="#f0f0f0")
        calendar_frame.pack(fill=tk.BOTH, pady=5)
        
        # 星期表头
        weekdays = ["一", "二", "三", "四", "五", "六", "日"]
        for i, day in enumerate(weekdays):
            lbl = tk.Label(
                calendar_frame, text=day,
                font=("Microsoft YaHei", 10, "bold"),
                bg="#e0e0e0", fg="#333333",
                width=6, height=1
            )
            lbl.grid(row=0, column=i, padx=1, pady=1)
        
        # 日历格子容器
        day_cells = {}  # 存储日期按钮引用
        
        def _render_calendar(year, month):
            """渲染日历"""
            # 清除旧的日期格子
            for widget in calendar_frame.grid_slaves():
                if int(widget.grid_info()["row"]) > 0:
                    widget.destroy()
            
            day_cells.clear()
            
            # 获取该月第一天是星期几（0=周一, 6=周日）
            first_weekday, days_in_month = calendar.monthrange(year, month)
            
            # 获取该月的日志数据
            logs = self.logger.get_month_logs(year, month)
            
            # 渲染日期
            row = 1
            col = first_weekday  # 0=周一, 6=周日
            
            for day in range(1, days_in_month + 1):
                date_str = f"{year:04d}-{month:02d}-{day:02d}"
                
                # 确定颜色
                if date_str in logs:
                    log = logs[date_str]
                    total_tasks = len(log["tasks"])
                    # 统计已完成任务（normal_complete + early_complete）
                    completed_tasks = sum(1 for t in log["tasks"] 
                        if t.get("completion_type") in ("normal_complete", "early_complete") 
                        or t.get("status") == "completed")
                    # 统计跳过任务
                    skipped_tasks = sum(1 for t in log["tasks"] 
                        if t.get("completion_type") == "skipped")
                    # 统计跳转终止任务
                    jump_terminated_tasks = sum(1 for t in log["tasks"] 
                        if t.get("completion_type") == "jump_terminated")
                    
                    # 判断颜色：所有任务都已处理（完成/跳过/跳转终止）则显示绿色
                    if completed_tasks + skipped_tasks + jump_terminated_tasks == total_tasks:
                        # 全部处理：深绿
                        bg_color = "#22C55E"
                        fg_color = "white"
                    else:
                        # 有未完成：黄色
                        bg_color = "#FCD34D"
                        fg_color = "#333333"
                else:
                    # 无任务：灰色
                    bg_color = "#E5E7EB"
                    fg_color = "#6B7280"
                
                # 日期按钮
                btn = tk.Button(
                    calendar_frame, text=str(day),
                    font=("Microsoft YaHei", 9),
                    bg=bg_color, fg=fg_color,
                    width=6, height=2,
                    command=lambda d=date_str: _show_day_detail(d)
                )
                btn.grid(row=row, column=col, padx=1, pady=1)
                day_cells[date_str] = btn
                
                col += 1
                if col > 6:
                    col = 0
                    row += 1
            
            # 更新月份标签
            month_label.config(text=f"{year}年{month}月")
        
        # ========== 统计区 ==========
        stats_frame = tk.LabelFrame(main_frame, text="月度统计", bg="#f0f0f0", fg="#333333", padx=10, pady=5)
        stats_frame.pack(fill=tk.X, pady=10)
        
        stats_labels = {}
        
        def _update_statistics(year, month):
            """更新统计数据"""
            stats = self.logger.get_month_statistics(year, month)
            
            # 清除旧标签
            for widget in stats_frame.winfo_children():
                widget.destroy()
            
            # 第一行：天数统计
            row1 = tk.Frame(stats_frame, bg="#f0f0f0")
            row1.pack(fill=tk.X, pady=2)
            tk.Label(row1, text=f"📊 有任务天数：{stats['total_days']}天", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            tk.Label(row1, text=f"📋 任务总数：{stats['total_tasks']}个", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            tk.Label(row1, text=f"✅ 已完成：{stats['completed_tasks']}个", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            tk.Label(row1, text=f"⏭️ 已跳过：{stats['skipped_tasks']}个", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            tk.Label(row1, text=f"🔄 跳转终止：{stats.get('jump_terminated_tasks', 0)}个", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            
            # 第二行：完成率
            row2 = tk.Frame(stats_frame, bg="#f0f0f0")
            row2.pack(fill=tk.X, pady=2)
            tk.Label(row2, text=f"📈 完成率：{stats['completion_rate']:.1f}%", 
                    font=("Microsoft YaHei", 10, "bold"), bg="#f0f0f0", fg="#22C55E").pack(side=tk.LEFT, padx=5)
            tk.Label(row2, text=f"⏱️ 总专注时长：{stats['completed_focus_minutes']:.1f}分钟", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            tk.Label(row2, text=f"☕ 总休息时长：{stats['total_break_minutes']}分钟", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            
            # 第三行：平均数据
            row3 = tk.Frame(stats_frame, bg="#f0f0f0")
            row3.pack(fill=tk.X, pady=2)
            tk.Label(row3, text=f"📅 日均任务数：{stats['avg_daily_tasks']:.1f}个", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
            tk.Label(row3, text=f"⏰ 平均任务时长：{stats['avg_task_duration']:.1f}分钟", 
                    font=("Microsoft YaHei", 10), bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        
        def _show_day_detail(date_str):
            """显示某天的日志详情"""
            log = self.logger.get_log(date_str)
            
            # 创建详情窗口
            detail_top = tk.Toplevel(top)
            detail_top.title(f"📅 {date_str} 日志详情")
            detail_top.attributes("-topmost", True)
            detail_top.resizable(False, False)
            
            # 主框架
            detail_frame = tk.Frame(detail_top, bg="#f0f0f0", padx=10, pady=10)
            detail_frame.pack(fill=tk.BOTH, expand=True)
            
            # 标题
            title_lbl = tk.Label(
                detail_frame, text=f"📋 {date_str} 任务记录",
                font=("Microsoft YaHei", 12, "bold"), bg="#f0f0f0"
            )
            title_lbl.pack(pady=(0, 10))
            
            if not log["tasks"]:
                # 无任务
                empty_lbl = tk.Label(
                    detail_frame, text="当天暂无任务记录",
                    font=("Microsoft YaHei", 10), bg="#f0f0f0", fg="#888888"
                )
                empty_lbl.pack(pady=20)
            else:
                # 任务列表
                list_frame = tk.Frame(detail_frame, bg="#f0f0f0")
                list_frame.pack(fill=tk.BOTH, expand=True)
                
                # 表头
                headers = ["序号", "任务名称", "时长", "休息", "状态", "完成时间"]
                for i, h in enumerate(headers):
                    lbl = tk.Label(
                        list_frame, text=h,
                        font=("Microsoft YaHei", 9, "bold"),
                        bg="#e0e0e0", fg="#333333",
                        width=12 if i == 1 else 8
                    )
                    lbl.grid(row=0, column=i, padx=2, pady=2)
                
                # 任务数据
                for row_idx, task in enumerate(log["tasks"], start=1):
                    # 序号
                    tk.Label(
                        list_frame, text=str(task["id"]),
                        font=("Microsoft YaHei", 9), bg="#f0f0f0",
                        width=8
                    ).grid(row=row_idx, column=0, padx=2, pady=1)
                    
                    # 任务名称
                    tk.Label(
                        list_frame, text=task["name"],
                        font=("Microsoft YaHei", 9), bg="#f0f0f0",
                        width=12, anchor="w"
                    ).grid(row=row_idx, column=1, padx=2, pady=1)
                    
                    # 时长
                    tk.Label(
                        list_frame, text=f"{task['duration']}分钟",
                        font=("Microsoft YaHei", 9), bg="#f0f0f0",
                        width=8
                    ).grid(row=row_idx, column=2, padx=2, pady=1)
                    
                    # 休息
                    tk.Label(
                        list_frame, text=f"{task['break_count']}次",
                        font=("Microsoft YaHei", 9), bg="#f0f0f0",
                        width=8
                    ).grid(row=row_idx, column=3, padx=2, pady=1)
                    
                    # 状态（兼容新旧格式）
                    comp_type = task.get("completion_type") or task.get("status", "")
                    if comp_type == "normal_complete":
                        status_text = "✅完成"
                        status_color = "#22C55E"
                    elif comp_type == "early_complete":
                        status_text = "⚡提前完成"
                        status_color = "#3B82F6"
                    else:
                        status_text = "⏭️跳过"
                        status_color = "#F59E0B"
                    tk.Label(
                        list_frame, text=status_text,
                        font=("Microsoft YaHei", 9), bg="#f0f0f0",
                        fg=status_color, width=10
                    ).grid(row=row_idx, column=4, padx=2, pady=1)
                    
                    # 完成时间
                    tk.Label(
                        list_frame, text=task["completed_at"].split(" ")[1],
                        font=("Microsoft YaHei", 9), bg="#f0f0f0",
                        width=8
                    ).grid(row=row_idx, column=5, padx=2, pady=1)
            
            # 关闭按钮
            close_btn = tk.Button(
                detail_frame, text="关闭", width=10,
                command=detail_top.destroy,
                font=("Microsoft YaHei", 9), bg="#e0e0e0", fg="#333333"
            )
            close_btn.pack(pady=10)
            
            # 调整窗口大小
            detail_top.update_idletasks()
            width = 550
            height = max(200, 100 + len(log["tasks"]) * 30)
            x = (detail_top.winfo_screenwidth() // 2) - (width // 2)
            y = (detail_top.winfo_screenheight() // 2) - (height // 2)
            detail_top.geometry(f"{width}x{height}+{x}+{y}")
        
        def _export_csv(year, month):
            """导出月度日志为CSV"""
            from tkinter import filedialog
            
            file_path = filedialog.asksaveasfilename(
                parent=top,
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv")],
                initialfile=f"任务日志_{year}年{month}月.csv"
            )
            
            if file_path:
                try:
                    self.logger.export_month_to_csv(year, month, file_path)
                    messagebox.showinfo("导出成功", f"日志已导出至：\n{file_path}", parent=top)
                except Exception as e:
                    messagebox.showerror("导出失败", f"导出失败：{str(e)}", parent=top)
        
        # 初始渲染
        _render_calendar(current_year, current_month)
        _update_statistics(current_year, current_month)
        
        # 图例说明
        legend_frame = tk.Frame(main_frame, bg="#f0f0f0")
        legend_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(legend_frame, text="图例：", font=("Microsoft YaHei", 9), bg="#f0f0f0").pack(side=tk.LEFT)
        tk.Label(legend_frame, text="  ■ 深绿=全部完成  ", font=("Microsoft YaHei", 9), 
                bg="#f0f0f0", fg="#22C55E").pack(side=tk.LEFT)
        tk.Label(legend_frame, text="  ■ 黄色=有未完成  ", font=("Microsoft YaHei", 9), 
                bg="#f0f0f0", fg="#F59E0B").pack(side=tk.LEFT)
        tk.Label(legend_frame, text="  ■ 灰色=无任务", font=("Microsoft YaHei", 9), 
                bg="#f0f0f0", fg="#6B7280").pack(side=tk.LEFT)
        
        # 关闭按钮
        close_btn = tk.Button(
            main_frame, text="关闭", width=10,
            command=top.destroy,
            font=("Microsoft YaHei", 9), bg="#e0e0e0", fg="#333333"
        )
        close_btn.pack(pady=5)
        
        # 调整窗口大小
        top.update_idletasks()
        width = 520
        height = 550
        x = (top.winfo_screenwidth() // 2) - (width // 2)
        y = (top.winfo_screenheight() // 2) - (height // 2)
        top.geometry(f"{width}x{height}+{x}+{y}")
    
    def _on_closing(self):
        """窗口关闭事件"""
        # 终止计时
        self.is_running = False
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
        
        self.root.destroy()


def main():
    """主函数"""
    root = tk.Tk()
    app = TimerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
