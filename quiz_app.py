"""
答题软件主程序
使用Python + tkinter构建Windows本地选择题答题应用
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import sys
import time
import traceback
from datetime import datetime
from typing import Optional, List, Dict
import threading

from database import QuizDatabase, list_saves, delete_save
from excel_handler import import_from_excel, export_to_excel, create_sample_excel

# 设置日志文件路径
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug.log")

def log_error(msg):
    """记录错误到日志文件"""
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"[{datetime.now()}] {msg}\n")
    except:
        pass

def log_exception(e):
    """记录异常详细信息到日志文件"""
    error_detail = traceback.format_exc()
    log_error(f"异常: {e}\n{error_detail}")

# 版本信息
VERSION = "1.0.0"

# 护眼背景色
THEMES = {
    'green': '#C7EDCC',   # 浅绿
    'beige': '#F5F5DC',   # 米黄
    'gray': '#E8E8E8',    # 浅灰
    'white': '#FFFFFF'    # 白色
}

# 题型显示名称
TYPE_NAMES = {
    'single': '单选题',
    'multiple': '多选题',
    'judge': '判断题'
}


class QuizApp:
    """答题软件主类"""
    
    def __init__(self, root: tk.Tk):
        """初始化应用"""
        self.root = root
        self.root.title(f"选择题答题软件 v{VERSION}")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        
        # 初始化变量
        self.db: Optional[QuizDatabase] = None
        self.current_question: Optional[Dict] = None
        self.current_index = 0
        self.questions: List[Dict] = []
        
        # 答题状态
        self.selected_options = []  # 当前选中的选项
        self.showing_answer = False
        self.timer_running = False
        self.timer_hidden = False
        self.timer_start = 0
        self.current_time = 0
        
        # 设置
        self.current_theme = 'green'
        self.save_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saves")
        os.makedirs(self.save_dir, exist_ok=True)
        
        # 创建界面
        self.create_menu()
        self.create_main_ui()
        self.create_bindings()
        
        # 加载默认存档或显示欢迎界面
        self.load_default_or_welcome()
    
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="导入Excel题目", command=self.import_excel)
        file_menu.add_separator()
        file_menu.add_command(label="新建存档", command=self.create_new_save)
        file_menu.add_command(label="加载存档", command=self.load_save_dialog)
        file_menu.add_command(label="另存为...", command=self.save_as_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="导出答题记录", command=self.export_results)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_close)
        
        # 视图菜单
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="视图", menu=view_menu)
        view_menu.add_command(label="题目导航", command=self.show_navigation)
        
        # 背景色子菜单
        theme_menu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="护眼背景", menu=theme_menu)
        theme_menu.add_command(label="浅绿色", command=lambda: self.set_theme('green'))
        theme_menu.add_command(label="米黄色", command=lambda: self.set_theme('beige'))
        theme_menu.add_command(label="浅灰色", command=lambda: self.set_theme('gray'))
        theme_menu.add_command(label="白色", command=lambda: self.set_theme('white'))
        
        view_menu.add_checkbutton(label="隐藏计时器", command=self.toggle_timer_visibility)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="快捷键", command=self.show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="关于", command=self.show_about)
    
    def create_main_ui(self):
        """创建主界面"""
        # 主框架
        self.main_frame = tk.Frame(self.root, bg=THEMES[self.current_theme])
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 上方区域：题目和导航
        top_frame = tk.Frame(self.main_frame, bg=THEMES[self.current_theme])
        top_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左上：题目区域
        left_frame = tk.Frame(top_frame, bg=THEMES[self.current_theme])
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 题目信息
        self.info_label = tk.Label(
            left_frame, text="未加载题目", font=("微软雅黑", 12),
            bg=THEMES[self.current_theme], fg="#333333"
        )
        self.info_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 题目文本
        self.question_text = tk.Text(
            left_frame, font=("微软雅黑", 14), wrap=tk.WORD,
            bg=THEMES[self.current_theme], fg="#333333",
            height=6, padx=10, pady=10, relief=tk.FLAT,
            spacing1=0, spacing2=12  # 1.5倍行距
        )
        self.question_text.pack(fill=tk.X, pady=(0, 15))
        self.question_text.config(state=tk.DISABLED)
        
        # 选项区域
        self.options_frame = tk.Frame(left_frame, bg=THEMES[self.current_theme])
        self.options_frame.pack(fill=tk.BOTH, expand=True)
        
        self.option_vars = []
        self.option_buttons = []
        self.create_option_buttons()
        
        # 右上：导航和统计
        right_frame = tk.Frame(top_frame, bg=THEMES[self.current_theme], width=200)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right_frame.pack_propagate(False)
        
        # 导航按钮
        nav_frame = tk.LabelFrame(
            right_frame, text="题目导航", font=("微软雅黑", 10),
            bg=THEMES[self.current_theme], fg="#333333"
        )
        nav_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Button(
            nav_frame, text="打开导航", font=("微软雅黑", 10),
            command=self.show_navigation, width=15
        ).pack(pady=5)
        
        # 统计信息区域
        stats_frame = tk.LabelFrame(
            right_frame, text="统计信息", font=("微软雅黑", 10),
            bg=THEMES[self.current_theme], fg="#333333"
        )
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.stats_labels = {}
        stats_items = [
            ("总题数", "total"),
            ("已答题", "answered"),
            ("正确", "correct"),
            ("错误", "wrong"),
            ("正确率", "accuracy"),
            ("收藏数", "favorite")
        ]
        
        for text, key in stats_items:
            frame = tk.Frame(stats_frame, bg=THEMES[self.current_theme])
            frame.pack(fill=tk.X, padx=5, pady=2)
            tk.Label(
                frame, text=f"{text}:", font=("微软雅黑", 9),
                bg=THEMES[self.current_theme], width=8, anchor=tk.W
            ).pack(side=tk.LEFT)
            label = tk.Label(
                frame, text="0", font=("微软雅黑", 9, "bold"),
                bg=THEMES[self.current_theme], fg="#0066CC"
            )
            label.pack(side=tk.RIGHT)
            self.stats_labels[key] = label
        
        # 下方区域：控制按钮
        bottom_frame = tk.Frame(self.main_frame, bg=THEMES[self.current_theme])
        bottom_frame.pack(fill=tk.X, pady=(15, 0))
        
        # 左下：收藏和跳转
        left_bottom = tk.Frame(bottom_frame, bg=THEMES[self.current_theme])
        left_bottom.pack(side=tk.LEFT)
        
        self.favorite_btn = tk.Button(
            left_bottom, text="☆ 收藏", font=("微软雅黑", 11),
            command=self.toggle_favorite, width=10
        )
        self.favorite_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.prev_btn = tk.Button(
            left_bottom, text="← 上一题", font=("微软雅黑", 11),
            command=self.prev_question, width=10
        )
        self.prev_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.next_btn = tk.Button(
            left_bottom, text="下一题 →", font=("微软雅黑", 11),
            command=self.next_question, width=10
        )
        self.next_btn.pack(side=tk.LEFT)
        
        # 右下：计时器和控制
        right_bottom = tk.Frame(bottom_frame, bg=THEMES[self.current_theme])
        right_bottom.pack(side=tk.RIGHT)
        
        # 计时器
        self.timer_frame = tk.Frame(right_bottom, bg=THEMES[self.current_theme])
        self.timer_frame.pack(side=tk.LEFT, padx=(0, 15))
        
        self.timer_label = tk.Label(
            self.timer_frame, text="00:00.0", font=("Consolas", 20, "bold"),
            bg=THEMES[self.current_theme], fg="#FF6600"
        )
        self.timer_label.pack()
        
        self.timer_btn = tk.Button(
            self.timer_frame, text="暂停", font=("微软雅黑", 9),
            command=self.toggle_timer, width=8
        )
        self.timer_btn.pack(pady=(5, 0))
        
        # 设置按钮
        tk.Button(
            right_bottom, text="⚙ 设置", font=("微软雅黑", 11),
            command=self.show_settings, width=8
        ).pack(side=tk.LEFT)
        
        # 答案显示标签
        self.answer_label = tk.Label(
            self.main_frame, text="", font=("微软雅黑", 14, "bold"),
            bg=THEMES[self.current_theme], fg="#009900"
        )
        self.answer_label.pack(pady=(10, 0))
        
        # 多选题确认按钮（初始隐藏）
        self.confirm_btn = tk.Button(
            self.main_frame, text="✓ 确认答案", font=("微软雅黑", 12, "bold"),
            command=self.check_answer, bg="#4CAF50", fg="white",
            width=15, height=1
        )
        self.confirm_btn.pack(pady=(10, 0))
        self.confirm_btn.pack_forget()  # 初始隐藏
    
    def create_option_buttons(self):
        """创建选项按钮"""
        # 清除旧按钮
        for widget in self.options_frame.winfo_children():
            widget.destroy()
        
        self.option_vars = []
        self.option_buttons = []
        
        option_labels = ['A', 'B', 'C', 'D', 'E', 'F']
        
        for i, label in enumerate(option_labels):
            var = tk.BooleanVar(value=False)
            self.option_vars.append(var)
            
            btn = tk.Checkbutton(
                self.options_frame,
                text=f"{label}. ",
                variable=var,
                font=("微软雅黑", 13),
                bg=THEMES[self.current_theme],
                fg="#333333",
                anchor=tk.W,
                padx=20, pady=8,
                command=lambda idx=i: self.on_option_click(idx)
            )
            btn.pack(fill=tk.X, pady=3)
            btn.config(state=tk.DISABLED)
            self.option_buttons.append(btn)
    
    def create_bindings(self):
        """创建键盘绑定"""
        self.root.bind("<Left>", lambda e: self.prev_question())
        self.root.bind("<Right>", lambda e: self.next_question())
        self.root.bind("<space>", lambda e: self.show_answer())
        self.root.bind("<Key>", self.on_key_press)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def load_default_or_welcome(self):
        """加载默认存档或显示欢迎界面"""
        saves = list_saves(self.save_dir)
        if saves:
            # 加载最新的存档
            self.load_save(os.path.join(self.save_dir, saves[0]['filename']))
        else:
            self.show_welcome()
    
    def show_welcome(self):
        """显示欢迎界面"""
        self.question_text.config(state=tk.NORMAL)
        self.question_text.delete(1.0, tk.END)
        self.question_text.insert(tk.END, "欢迎使用选择题答题软件！\n\n")
        self.question_text.insert(tk.END, "请通过菜单导入Excel题目文件开始使用。\n")
        self.question_text.insert(tk.END, "支持题型：单选题、多选题、判断题\n\n")
        self.question_text.insert(tk.END, "快捷键：\n")
        self.question_text.insert(tk.END, "  ← →  上一题/下一题\n")
        self.question_text.insert(tk.END, "  空格键  显示答案\n")
        self.question_text.config(state=tk.DISABLED)
    
    # ==================== 题目操作 ====================
    
    def load_question(self, index: int):
        """
        加载指定题目
        
        Args:
            index: 题目索引（从0开始）
        """
        if not self.questions or index < 0 or index >= len(self.questions):
            return
        
        self.current_index = index
        self.current_question = self.questions[index]
        q = self.current_question
        
        # 重置状态
        self.selected_options = []
        self.showing_answer = False
        self.answer_label.config(text="")
        
        # 更新题目信息
        type_name = TYPE_NAMES.get(q['type'], q['type'])
        status = ""
        if q.get('is_answered'):
            status = " ✓" if q.get('is_correct') else " ✗"
        
        self.info_label.config(
            text=f"第 {q['seq_no']} 题 / 共 {len(self.questions)} 题 | {type_name}{status}"
        )
        
        # 更新收藏按钮
        self.update_favorite_button()
        
        # 更新题目文本
        self.question_text.config(state=tk.NORMAL)
        self.question_text.delete(1.0, tk.END)
        self.question_text.insert(tk.END, q['question'])
        self.question_text.config(state=tk.DISABLED)
        
        # 更新选项
        options = [
            q.get('option_a', ''),
            q.get('option_b', ''),
            q.get('option_c', ''),
            q.get('option_d', ''),
            q.get('option_e', ''),
            q.get('option_f', '')
        ]
        
        for i, (btn, opt_text) in enumerate(zip(self.option_buttons, options)):
            if opt_text:
                btn.config(
                    text=f"{'ABCDEF'[i]}. {opt_text}",
                    state=tk.NORMAL,
                    bg=THEMES[self.current_theme]
                )
            else:
                btn.config(
                    text=f"{'ABCDEF'[i]}.",
                    state=tk.DISABLED,
                    bg=THEMES[self.current_theme]
                )
            self.option_vars[i].set(False)
        
        # 如果已经答过，显示之前的选择
        if q.get('is_answered') and q.get('user_answer'):
            user_ans = q['user_answer']
            for char in user_ans:
                idx = ord(char) - ord('A')
                if 0 <= idx < 6:
                    self.option_vars[idx].set(True)
                    self.selected_options.append(char)
            
            # 根据对错显示不同颜色
            if q.get('is_correct'):
                self.highlight_correct()
            else:
                self.highlight_wrong(q['answer'])
            
            # 已答过的题目隐藏确认按钮
            self.confirm_btn.pack_forget()
        else:
            # 未答的题目，多选题显示确认按钮
            if q['type'] == 'multiple':
                self.confirm_btn.pack(pady=(10, 0))
            else:
                self.confirm_btn.pack_forget()
        
        # 重置并启动计时器
        self.reset_timer()
        self.start_timer()
        
        # 更新统计
        self.update_stats()
    
    def on_option_click(self, idx: int):
        """选项点击处理"""
        if not self.current_question:
            return
        
        q = self.current_question
        option_char = chr(ord('A') + idx)
        
        # 如果已经答过，不允许修改
        if q.get('is_answered'):
            return
        
        if q['type'] == 'multiple':
            # 多选题：允许多选，需要确认
            if self.option_vars[idx].get():
                if option_char not in self.selected_options:
                    self.selected_options.append(option_char)
            else:
                if option_char in self.selected_options:
                    self.selected_options.remove(option_char)
            self.selected_options.sort()
        else:
            # 单选题/判断题：立即判断
            for i, var in enumerate(self.option_vars):
                var.set(i == idx)
            
            self.selected_options = [option_char]
            self.check_answer()
    
    def check_answer(self):
        """检查答案"""
        if not self.current_question or not self.selected_options:
            return
        
        q = self.current_question
        user_ans = ''.join(self.selected_options)
        correct_ans = q['answer'].upper()
        
        is_correct = (user_ans == correct_ans)
        
        # 停止计时
        self.stop_timer()
        time_spent = self.current_time
        
        # 更新数据库
        wrong_increment = 1 if not is_correct else 0
        self.db.update_answer(q['seq_no'], user_ans, is_correct, time_spent, wrong_increment)
        
        # 更新本地数据
        q['is_answered'] = 1
        q['is_correct'] = 1 if is_correct else 0
        q['user_answer'] = user_ans
        q['time_spent'] = time_spent
        if not is_correct:
            q['wrong_count'] = q.get('wrong_count', 0) + 1
        
        if is_correct:
            self.highlight_correct()
            self.confirm_btn.pack_forget()  # 隐藏确认按钮
            # 延迟0.2秒自动下一题
            self.root.after(200, self.next_question)
        else:
            self.highlight_wrong(correct_ans)
            self.confirm_btn.pack_forget()  # 隐藏确认按钮
            # 停留当前题，等待用户按→键或按钮
    
    def highlight_correct(self):
        """高亮显示正确答案"""
        self.answer_label.config(text="✓ 回答正确！", fg="#009900")
        for i, btn in enumerate(self.option_buttons):
            if self.option_vars[i].get():
                btn.config(bg="#90EE90")  # 浅绿色
    
    def highlight_wrong(self, correct_answer: str):
        """高亮显示错误答案"""
        self.answer_label.config(text=f"✗ 回答错误！正确答案：{correct_answer}", fg="#CC0000")
        
        # 用户选的标红，正确答案标绿
        for i, btn in enumerate(self.option_buttons):
            char = chr(ord('A') + i)
            if self.option_vars[i].get():
                btn.config(bg="#FFB6C1")  # 浅红色
            elif char in correct_answer:
                btn.config(bg="#90EE90")  # 浅绿色
    
    def show_answer(self):
        """显示正确答案（空格键）"""
        if not self.current_question:
            return
        
        q = self.current_question
        correct_ans = q['answer'].upper()
        
        # 高亮正确答案
        for i, btn in enumerate(self.option_buttons):
            char = chr(ord('A') + i)
            if char in correct_ans:
                btn.config(bg="#90EE90")
        
        self.answer_label.config(
            text=f"正确答案：{correct_ans}", fg="#0066CC"
        )
    
    def prev_question(self):
        """上一题"""
        if self.current_index > 0:
            self.load_question(self.current_index - 1)
    
    def next_question(self):
        """下一题"""
        if self.current_index < len(self.questions) - 1:
            self.load_question(self.current_index + 1)
    
    # ==================== 计时器 ====================
    
    def reset_timer(self):
        """重置计时器"""
        self.current_time = 0
        self.timer_start = time.time()
        self.update_timer_display()
    
    def start_timer(self):
        """启动计时器"""
        self.timer_running = True
        self.timer_start = time.time() - self.current_time
        self.update_timer()
        self.timer_btn.config(text="暂停")
    
    def stop_timer(self):
        """停止计时器"""
        self.timer_running = False
        self.timer_btn.config(text="继续")
    
    def toggle_timer(self):
        """切换计时器状态"""
        if self.timer_running:
            self.stop_timer()
        else:
            self.start_timer()
    
    def update_timer(self):
        """更新计时器"""
        if self.timer_running:
            self.current_time = time.time() - self.timer_start
            self.update_timer_display()
            self.root.after(100, self.update_timer)  # 每100ms更新一次
    
    def update_timer_display(self):
        """更新计时器显示"""
        minutes = int(self.current_time // 60)
        seconds = int(self.current_time % 60)
        tenths = int((self.current_time * 10) % 10)
        self.timer_label.config(text=f"{minutes:02d}:{seconds:02d}.{tenths}")
    
    def toggle_timer_visibility(self):
        """切换计时器可见性"""
        self.timer_hidden = not self.timer_hidden
        if self.timer_hidden:
            self.timer_frame.pack_forget()
        else:
            self.timer_frame.pack(side=tk.LEFT, padx=(0, 15))
    
    # ==================== 收藏 ====================
    
    def toggle_favorite(self):
        """切换收藏状态"""
        if not self.current_question:
            return
        
        q = self.current_question
        is_fav = self.db.toggle_favorite(q['seq_no'])
        q['is_favorite'] = 1 if is_fav else 0
        
        self.update_favorite_button()
        self.update_stats()
    
    def update_favorite_button(self):
        """更新收藏按钮显示"""
        if self.current_question and self.current_question.get('is_favorite'):
            self.favorite_btn.config(text="★ 已收藏", fg="#FF6600")
        else:
            self.favorite_btn.config(text="☆ 收藏", fg="#333333")
    
    # ==================== 统计 ====================
    
    def update_stats(self):
        """更新统计信息显示"""
        if not self.db:
            return
        
        stats = self.db.get_statistics()
        
        self.stats_labels['total'].config(text=str(stats['total']))
        self.stats_labels['answered'].config(text=str(stats['answered']))
        self.stats_labels['correct'].config(text=str(stats['correct']))
        self.stats_labels['wrong'].config(text=str(stats['wrong']))
        self.stats_labels['accuracy'].config(text=f"{stats['accuracy']:.1f}%")
        self.stats_labels['favorite'].config(text=str(stats['favorite']))
    
    # ==================== 导航 ====================
    
    def show_navigation(self):
        """显示题目导航弹窗"""
        if not self.questions:
            messagebox.showinfo("提示", "暂无题目")
            return
        
        nav_window = tk.Toplevel(self.root)
        nav_window.title("题目导航")
        nav_window.geometry("500x400")
        nav_window.transient(self.root)
        nav_window.grab_set()
        
        # 图例说明
        legend_frame = tk.Frame(nav_window)
        legend_frame.pack(fill=tk.X, padx=10, pady=5)
        
        legends = [
            ("#90EE90", "已答对"),
            ("#FFB6C1", "已答错"),
            ("#FFD700", "当前"),
            ("#E8E8E8", "未答")
        ]
        
        for color, text in legends:
            frame = tk.Frame(legend_frame)
            frame.pack(side=tk.LEFT, padx=5)
            tk.Label(frame, bg=color, width=2, height=1).pack(side=tk.LEFT)
            tk.Label(frame, text=text, font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=(2, 0))
        
        # 题号网格
        canvas = tk.Canvas(nav_window)
        scrollbar = ttk.Scrollbar(nav_window, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建题号按钮
        cols = 8
        for i, q in enumerate(self.questions):
            row = i // cols
            col = i % cols
            
            # 确定背景色
            if i == self.current_index:
                bg_color = "#FFD700"  # 当前题目 - 金色
            elif q.get('is_answered'):
                bg_color = "#90EE90" if q.get('is_correct') else "#FFB6C1"
            else:
                bg_color = "#E8E8E8"
            
            btn = tk.Button(
                scroll_frame,
                text=str(q['seq_no']),
                width=5, height=2,
                bg=bg_color,
                command=lambda idx=i: [nav_window.destroy(), self.load_question(idx)]
            )
            btn.grid(row=row, column=col, padx=3, pady=3)
    
    # ==================== 文件操作 ====================
    
    def import_excel(self):
        """导入Excel题目"""
        filepath = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            questions = import_from_excel(filepath)
            if not questions:
                messagebox.showwarning("警告", "未找到有效题目")
                return
            
            # 创建新数据库
            save_name = os.path.splitext(os.path.basename(filepath))[0]
            db_path = os.path.join(self.save_dir, f"{save_name}.db")
            
            # 如果已存在，添加数字后缀
            counter = 1
            original_path = db_path
            while os.path.exists(db_path):
                db_path = original_path.replace('.db', f'_{counter}.db')
                counter += 1
            
            # 关闭旧数据库
            if self.db:
                self.db.close()
            
            # 创建新数据库并导入
            self.db = QuizDatabase(db_path)
            self.db.import_questions(questions)
            self.db.update_save_name(save_name)
            
            self.questions = self.db.get_all_questions()
            self.current_index = 0
            
            if self.questions:
                self.load_question(0)
            
            messagebox.showinfo("成功", f"成功导入 {len(questions)} 道题目")
            
        except ValueError as e:
            # 数据格式错误，显示详细的行号信息
            error_msg = f"Excel格式错误：\n\n{str(e)}\n\n请检查：\n1. B列（题型）填写正确：single/multiple/judge 或 单选/多选/判断\n2. C列（题干）不能为空\n3. J列（答案）不能为空"
            messagebox.showerror("导入失败", error_msg)
            log_error(f"导入格式错误: {e}")
        except Exception as e:
            error_msg = f"导入失败：{str(e)}"
            messagebox.showerror("导入失败", error_msg)
            log_exception(e)
    
    def create_new_save(self):
        """创建新存档"""
        name = simpledialog.askstring("新建存档", "请输入存档名称：")
        if not name:
            return
        
        db_path = os.path.join(self.save_dir, f"{name}.db")
        if os.path.exists(db_path):
            messagebox.showerror("错误", "存档已存在")
            return
        
        if self.db:
            self.db.close()
        
        self.db = QuizDatabase(db_path)
        self.db.update_save_name(name)
        self.questions = []
        self.current_question = None
        
        self.show_welcome()
        messagebox.showinfo("成功", "新存档已创建，请导入题目")
    
    def load_save_dialog(self):
        """加载存档对话框"""
        saves = list_saves(self.save_dir)
        if not saves:
            messagebox.showinfo("提示", "暂无存档")
            return
        
        load_window = tk.Toplevel(self.root)
        load_window.title("加载存档")
        load_window.geometry("500x350")
        load_window.transient(self.root)
        load_window.grab_set()
        
        # 存档列表
        frame = tk.Frame(load_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        headers = ["存档名称", "总题数", "已答", "正确", "更新时间"]
        for col, header in enumerate(headers):
            tk.Label(frame, text=header, font=("微软雅黑", 10, "bold")).grid(
                row=0, column=col, padx=5, pady=5, sticky=tk.W
            )
        
        for row_idx, save in enumerate(saves, 1):
            tk.Label(frame, text=save['name']).grid(row=row_idx, column=0, padx=5, pady=3, sticky=tk.W)
            tk.Label(frame, text=str(save['total'])).grid(row=row_idx, column=1, padx=5, pady=3)
            tk.Label(frame, text=str(save['answered'])).grid(row=row_idx, column=2, padx=5, pady=3)
            tk.Label(frame, text=str(save['correct'])).grid(row=row_idx, column=3, padx=5, pady=3)
            tk.Label(frame, text=save['updated'][:16]).grid(row=row_idx, column=4, padx=5, pady=3)
            
            # 加载按钮
            tk.Button(
                frame, text="加载",
                command=lambda s=save: [
                    load_window.destroy(),
                    self.load_save(os.path.join(self.save_dir, s['filename']))
                ]
            ).grid(row=row_idx, column=5, padx=5, pady=3)
            
            # 删除按钮
            tk.Button(
                frame, text="删除", fg="red",
                command=lambda s=save: self.delete_save_confirm(s, load_window)
            ).grid(row=row_idx, column=6, padx=5, pady=3)
    
    def load_save(self, db_path: str):
        """加载指定存档"""
        try:
            if self.db:
                self.db.close()
            
            self.db = QuizDatabase(db_path)
            self.questions = self.db.get_all_questions()
            
            if self.questions:
                self.load_question(0)
            else:
                self.show_welcome()
                
        except Exception as e:
            messagebox.showerror("错误", f"加载存档失败：{str(e)}")
    
    def delete_save_confirm(self, save: Dict, parent_window: tk.Toplevel):
        """确认删除存档"""
        if messagebox.askyesno("确认", f"确定要删除存档 \"{save['name']}\" 吗？"):
            db_path = os.path.join(self.save_dir, save['filename'])
            if delete_save(db_path):
                messagebox.showinfo("成功", "存档已删除")
                parent_window.destroy()
                self.load_save_dialog()
            else:
                messagebox.showerror("错误", "删除失败")
    
    def save_as_dialog(self):
        """另存为对话框"""
        if not self.db:
            messagebox.showwarning("警告", "当前没有存档")
            return
        
        name = simpledialog.askstring("另存为", "请输入新存档名称：")
        if not name:
            return
        
        new_path = os.path.join(self.save_dir, f"{name}.db")
        if os.path.exists(new_path):
            messagebox.showerror("错误", "存档已存在")
            return
        
        try:
            import shutil
            current_db = self.db.db_path
            self.db.close()
            shutil.copy2(current_db, new_path)
            self.db = QuizDatabase(new_path)
            self.db.update_save_name(name)
            messagebox.showinfo("成功", "存档已复制")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")
    
    def export_results(self):
        """导出答题记录"""
        if not self.db:
            messagebox.showwarning("警告", "当前没有存档")
            return
        
        export_window = tk.Toplevel(self.root)
        export_window.title("导出答题记录")
        export_window.geometry("400x300")
        export_window.transient(self.root)
        export_window.grab_set()
        
        tk.Label(export_window, text="选择导出范围：", font=("微软雅黑", 12)).pack(pady=10)
        
        export_type = tk.StringVar(value="all")
        
        tk.Radiobutton(export_window, text="所有已做题目", variable=export_type, 
                      value="all", font=("微软雅黑", 11)).pack(anchor=tk.W, padx=30, pady=5)
        
        tk.Radiobutton(export_window, text="错题（错题次数>0）", variable=export_type,
                      value="wrong", font=("微软雅黑", 11)).pack(anchor=tk.W, padx=30, pady=5)
        
        tk.Radiobutton(export_window, text="收藏题目", variable=export_type,
                      value="favorite", font=("微软雅黑", 11)).pack(anchor=tk.W, padx=30, pady=5)
        
        # 超时题目
        timeout_frame = tk.Frame(export_window)
        timeout_frame.pack(fill=tk.X, padx=30, pady=5)
        
        tk.Radiobutton(timeout_frame, text="超时题目（超过", variable=export_type,
                      value="timeout", font=("微软雅黑", 11)).pack(side=tk.LEFT)
        timeout_entry = tk.Entry(timeout_frame, width=6, font=("微软雅黑", 11))
        timeout_entry.pack(side=tk.LEFT)
        timeout_entry.insert(0, "60")
        tk.Label(timeout_frame, text="秒）", font=("微软雅黑", 11)).pack(side=tk.LEFT)
        
        def do_export():
            etype = export_type.get()
            
            if etype == "all":
                questions = self.db.get_answered_questions()
            elif etype == "wrong":
                questions = self.db.get_wrong_questions()
            elif etype == "favorite":
                questions = self.db.get_favorite_questions()
            elif etype == "timeout":
                try:
                    timeout = float(timeout_entry.get())
                    questions = self.db.get_timeout_questions(timeout)
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的时间")
                    return
            else:
                questions = []
            
            if not questions:
                messagebox.showinfo("提示", "没有符合条件的题目")
                return
            
            filepath = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not filepath:
                return
            
            try:
                export_to_excel(filepath, questions)
                messagebox.showinfo("成功", f"已导出 {len(questions)} 道题目")
                export_window.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{str(e)}")
        
        tk.Button(export_window, text="导出", font=("微软雅黑", 12),
                 command=do_export, width=15).pack(pady=20)
    
    # ==================== 设置和主题 ====================
    
    def set_theme(self, theme_name: str):
        """设置主题背景色"""
        self.current_theme = theme_name
        color = THEMES[theme_name]
        
        # 更新所有组件背景色
        for widget in [self.main_frame, self.options_frame]:
            widget.config(bg=color)
        
        # 递归更新所有子组件
        self.update_widget_bg(self.main_frame, color)
        
        # 更新选项按钮
        for btn in self.option_buttons:
            if btn['state'] != tk.DISABLED:
                btn.config(bg=color)
    
    def update_widget_bg(self, widget, color):
        """递归更新组件背景色"""
        try:
            if hasattr(widget, 'config') and 'bg' in widget.keys():
                # 跳过特殊组件
                if not isinstance(widget, tk.Button):
                    widget.config(bg=color)
        except:
            pass
        
        for child in widget.winfo_children():
            self.update_widget_bg(child, color)
    
    def show_settings(self):
        """显示设置对话框"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("设置")
        settings_window.geometry("400x300")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        # 主题设置
        theme_frame = tk.LabelFrame(settings_window, text="护眼背景", font=("微软雅黑", 11))
        theme_frame.pack(fill=tk.X, padx=20, pady=10)
        
        themes = [
            ("浅绿色", 'green'),
            ("米黄色", 'beige'),
            ("浅灰色", 'gray'),
            ("白色", 'white')
        ]
        
        for name, key in themes:
            tk.Button(
                theme_frame, text=name, width=10,
                command=lambda k=key: [self.set_theme(k), settings_window.destroy()]
            ).pack(side=tk.LEFT, padx=5, pady=10)
        
        # 统计信息
        stats_frame = tk.LabelFrame(settings_window, text="当前存档统计", font=("微软雅黑", 11))
        stats_frame.pack(fill=tk.X, padx=20, pady=10)
        
        if self.db:
            stats = self.db.get_statistics()
            info_text = f"""
存档名称：{self.db.get_save_name()}
总题数：{stats['total']}
已答题：{stats['answered']}
正确：{stats['correct']}  错误：{stats['wrong']}
正确率：{stats['accuracy']:.1f}%
收藏数：{stats['favorite']}
            """
            tk.Label(stats_frame, text=info_text, font=("微软雅黑", 10),
                    justify=tk.LEFT).pack(padx=10, pady=10)
    
    # ==================== 帮助 ====================
    
    def show_help(self):
        """显示使用说明"""
        help_text = """
【使用说明】

1. 导入题目
   - 通过 文件 → 导入Excel题目 导入题目文件
   - Excel格式：A-序号 B-题型 C-题干 D-I-选项A-F J-答案

2. 答题操作
   - 单选题/判断题：点击选项立即判断
   - 多选题：可多选，按确认后判断（需完全匹配）
   - 正确：自动0.2秒后下一题
   - 错误：停留当前题，按→键或按钮跳转

3. 快捷键
   - ← → ：上一题/下一题
   - 空格键：显示正确答案

4. 计时功能
   - 每题独立计时
   - 可暂停/继续
   - 可在菜单中隐藏计时器

5. 错题与收藏
   - 错题会自动累计次数
   - 可收藏重要题目
   - 支持导出错题和收藏

6. 存档管理
   - 支持多存档
   - 做题进度自动保存
   - 可随时切换存档
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("500x500")
        
        text = tk.Text(help_window, font=("微软雅黑", 11), wrap=tk.WORD, padx=10, pady=10)
        text.pack(fill=tk.BOTH, expand=True)
        text.insert(tk.END, help_text)
        text.config(state=tk.DISABLED)
    
    def show_shortcuts(self):
        """显示快捷键"""
        shortcuts_text = """
【快捷键列表】

题目导航
  ←        上一题
  →        下一题
  空格键    显示正确答案

其他
  可通过菜单进行所有操作
        """
        messagebox.showinfo("快捷键", shortcuts_text)
    
    def show_about(self):
        """显示关于对话框"""
        about_text = f"""
选择题答题软件 v{VERSION}

技术栈：
- Python + tkinter
- openpyxl (Excel处理)
- SQLite (数据存储)

功能特点：
- 支持单选/多选/判断题
- 错题累计与收藏
- 每题独立计时
- 多存档管理
- 护眼背景色
- 完整的导入导出功能
        """
        messagebox.showinfo("关于", about_text)
    
    # ==================== 事件处理 ====================
    
    def on_key_press(self, event):
        """键盘按键处理"""
        # A-F键选择选项
        if event.char and event.char.upper() in 'ABCDEF':
            idx = ord(event.char.upper()) - ord('A')
            if 0 <= idx < 6:
                if self.option_buttons[idx]['state'] != tk.DISABLED:
                    # 切换选中状态
                    current = self.option_vars[idx].get()
                    self.option_vars[idx].set(not current)
                    self.on_option_click(idx)
    
    def on_close(self):
        """关闭窗口处理"""
        if self.db:
            self.db.close()
        self.root.destroy()


def main():
    """主函数"""
    root = tk.Tk()
    
    # 设置DPI感知（Windows高DPI支持）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = QuizApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
