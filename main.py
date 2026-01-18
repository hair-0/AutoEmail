import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, colorchooser
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import sqlite3
import os
import threading
import queue
import time
import datetime

# 尝试导入openpyxl
try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

# ==========================================
# 核心UI组件库 - Liquid Glass 风格
# ==========================================

class ModernTheme:
    """现代配色方案"""
    COLORS = {
        "bg_app": "#f3f4f6",           # 整体应用背景
        "sidebar_bg": "#ffffff",       # 侧边栏背景
        "card_bg": "#ffffff",          # 卡片背景
        "card_border": "#e5e7eb",      # 卡片边框
        "primary": "#6366f1",          # 主色 (Indigo)
        "primary_light": "#e0e7ff",    # 主色浅色背景
        "primary_hover": "#4f46e5",    # 主色悬停
        "text_main": "#1f2937",        # 主文字
        "text_sub": "#6b7280",         # 次要文字
        "danger": "#ef4444",           # 危险色
        "success": "#10b981",          # 成功色
        "input_bg": "#f9fafb",         # 输入框背景
        "shadow": "#9ca3af"            # 阴影颜色
    }
    
    FONTS = {
        "h1": ("Microsoft YaHei", 16, "bold"),
        "h2": ("Microsoft YaHei", 13, "bold"),
        "body": ("Microsoft YaHei", 10),
        "input": ("Microsoft YaHei", 10),
        "icon": ("Segoe UI Emoji", 12)
    }

class ShadowElement(tk.Canvas):
    """用于绘制带有柔和阴影的圆角容器"""
    def __init__(self, parent, radius=15, color="#ffffff", padding=15, **kwargs):
        super().__init__(parent, highlightthickness=0, bg=ModernTheme.COLORS["bg_app"], **kwargs)
        self.radius = radius
        self.color = color
        self.padding = padding
        self.inner_frame = tk.Frame(self, bg=color)
        self.bind("<Configure>", self._draw)

    def _draw(self, event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        
        # 绘制多层阴影
        self._create_rounded_rect(2, 4, w-2, h-2, self.radius, fill="#e5e7eb", outline="")
        self._create_rounded_rect(1, 2, w-3, h-3, self.radius, fill="#d1d5db", outline="")
        
        # 主体
        self._create_rounded_rect(0, 0, w-5, h-5, self.radius, fill=self.color, outline=ModernTheme.COLORS["card_border"])
        
        # 放置内容容器
        self.create_window(self.padding, self.padding, window=self.inner_frame, 
                         anchor="nw", width=w-5-(self.padding*2), height=h-5-(self.padding*2))

    def _create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = (x1+r, y1, x1+r, y1, x2-r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y1+r, x2, y2-r, x2, y2-r, x2, y2, x2-r, y2, x2-r, y2, x1+r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y2-r, x1, y1+r, x1, y1+r, x1, y1)
        return self.create_polygon(points, **kwargs, smooth=True)

class CapsuleButton(tk.Canvas):
    """胶囊形状的按钮"""
    def __init__(self, parent, text, command, bg_color=ModernTheme.COLORS["primary"], text_color="white", width=100, height=35):
        super().__init__(parent, width=width, height=height, bg=parent["bg"], highlightthickness=0)
        self.text = text
        self.command = command
        self.bg_color = bg_color
        self.text_color = text_color
        self.hover_color = bg_color # 简单起见，暂不处理复杂变色
        
        self.bind("<Button-1>", lambda e: self.command())
        self.bind("<Configure>", self._draw)

    def _draw(self, event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        
        self.create_oval(0, 0, h, h, fill=self.bg_color, outline="")
        self.create_oval(w-h, 0, w, h, fill=self.bg_color, outline="")
        self.create_rectangle(h/2, 0, w-h/2, h, fill=self.bg_color, outline="")
        
        self.create_text(w/2, h/2, text=self.text, fill=self.text_color, font=("Microsoft YaHei", 10, "bold"))

class SidebarButton(tk.Canvas):
    """侧边栏导航按钮"""
    def __init__(self, parent, text, icon, command, is_selected=False):
        super().__init__(parent, height=55, bg=ModernTheme.COLORS["sidebar_bg"], highlightthickness=0)
        self.text = text
        self.icon = icon
        self.command = command
        self.is_selected = is_selected
        self.hovering = False
        
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", lambda e: command())
        self.bind("<Configure>", self._draw)

    def set_selected(self, val):
        self.is_selected = val
        self._draw()

    def _on_enter(self, e):
        self.hovering = True
        self._draw()

    def _on_leave(self, e):
        self.hovering = False
        self._draw()

    def _draw(self, event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        
        bg = ModernTheme.COLORS["primary_light"] if self.is_selected else (ModernTheme.COLORS["bg_app"] if self.hovering else ModernTheme.COLORS["sidebar_bg"])
        fg = ModernTheme.COLORS["primary"] if self.is_selected else ModernTheme.COLORS["text_sub"]
        
        if self.is_selected:
            self.create_polygon(10, 5, w-10, 5, w-10, h-5, 10, h-5, smooth=True, fill=bg, outline="")
            self.create_line(10, 15, 10, h-15, width=4, fill=ModernTheme.COLORS["primary"], capstyle=tk.ROUND)
        elif self.hovering:
             self.create_polygon(15, 8, w-15, 8, w-15, h-8, 15, h-8, smooth=True, fill=bg, outline="")

        font_size = 11 if self.is_selected else 10
        icon_size = 14 if self.is_selected else 12
        
        self.create_text(45, h/2, text=self.icon, font=("Segoe UI Emoji", icon_size), fill=fg, anchor="center")
        self.create_text(75, h/2, text=self.text, font=("Microsoft YaHei", font_size, "bold" if self.is_selected else "normal"), fill=fg, anchor="w")

class ModernEntry(tk.Entry):
    """美化输入框"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, relief="flat", bg=ModernTheme.COLORS["input_bg"], 
                        fg=ModernTheme.COLORS["text_main"], insertbackground=ModernTheme.COLORS["primary"],
                        highlightthickness=1, highlightbackground="#e5e7eb", highlightcolor=ModernTheme.COLORS["primary"],
                        **kwargs)

# ==========================================
# 富文本编辑器工具栏 (通用版)
# ==========================================
class EditorToolbar(tk.Frame):
    def __init__(self, parent, text_widget, bg_color="white"):
        super().__init__(parent, bg=bg_color, pady=5)
        self.text_widget = text_widget
        
        # 字体
        self.font_fam = ttk.Combobox(self, values=["Microsoft YaHei", "Arial", "SimSun"], width=12, state="readonly")
        self.font_fam.set("Microsoft YaHei")
        self.font_fam.pack(side=tk.LEFT, padx=5)
        self.font_fam.bind("<<ComboboxSelected>>", self.update_font)
        
        # 字号
        self.font_size = ttk.Combobox(self, values=[str(i) for i in range(10, 36, 2)], width=4, state="readonly")
        self.font_size.set("12")
        self.font_size.pack(side=tk.LEFT, padx=5)
        self.font_size.bind("<<ComboboxSelected>>", self.update_font)
        
        ttk.Separator(self, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        
        # 样式按钮
        self._add_btn("B", "bold", lambda: self.toggle_tag("bold"))
        self._add_btn("I", "italic", lambda: self.toggle_tag("italic"))
        self._add_btn("U", "underline", lambda: self.toggle_tag("underline"))
        
        ttk.Separator(self, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        
        # 颜色
        tk.Button(self, text="🎨", fg=ModernTheme.COLORS["text_main"], font=("Segoe UI Emoji", 10), 
                 command=self.choose_color, relief="flat", bg=bg_color).pack(side=tk.LEFT, padx=2)

    def _add_btn(self, text, style, cmd):
        font_spec = ("serif", 10, style)
        tk.Button(self, text=text, font=font_spec, command=cmd, width=3, relief="flat", bg="#f3f4f6").pack(side=tk.LEFT, padx=2)

    def update_font(self, e=None):
        if not self.text_widget: return
        f = (self.font_fam.get(), int(self.font_size.get()))
        self.text_widget.configure(font=f)
        self.text_widget.tag_configure("bold", font=f + ("bold",))
        self.text_widget.tag_configure("italic", font=f + ("italic",))
        self.text_widget.tag_configure("underline", font=f + ("underline",), underline=True)

    def toggle_tag(self, tag):
        self.update_font()
        try:
            if self.text_widget.tag_ranges("sel"):
                current = self.text_widget.tag_names("sel.first")
                if tag in current:
                    self.text_widget.tag_remove(tag, "sel.first", "sel.last")
                else:
                    self.text_widget.tag_add(tag, "sel.first", "sel.last")
        except: pass

    def choose_color(self):
        c = colorchooser.askcolor()[1]
        if c and self.text_widget.tag_ranges("sel"):
            tag_name = f"col_{c}"
            self.text_widget.tag_configure(tag_name, foreground=c)
            self.text_widget.tag_add(tag_name, "sel.first", "sel.last")

# ==========================================
# 主程序逻辑
# ==========================================

class EmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Aura Mail Sender v4.0")
        self.root.geometry("1300x900")
        self.root.configure(bg=ModernTheme.COLORS["bg_app"])
        
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except: pass

        self.setup_ttk_styles()
        
        self.db_path = os.path.join(os.path.dirname(__file__), 'email_data.db')
        self.attachment_files = []
        self.send_queue = queue.Queue()
        self.pending_emails = {}
        
        self.init_db()
        self.create_layout()
        self.load_config()
        self.start_queue_worker()

    def setup_ttk_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", 
                       background="white", fieldbackground="white", foreground=ModernTheme.COLORS["text_main"],
                       rowheight=32, borderwidth=0, font=("Microsoft YaHei", 10))
        style.configure("Treeview.Heading", 
                       background="#f9fafb", foreground=ModernTheme.COLORS["text_sub"],
                       font=("Microsoft YaHei", 10, "bold"), relief="flat")
        style.map("Treeview", background=[('selected', ModernTheme.COLORS["primary_light"])], foreground=[('selected', ModernTheme.COLORS["primary"])])
        style.configure("Vertical.TScrollbar", troughcolor="#f3f4f6", background="#d1d5db", borderwidth=0, arrowsize=12)

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY, name TEXT, email TEXT, title TEXT, department TEXT)')
        c.execute('CREATE TABLE IF NOT EXISTS history (id INTEGER PRIMARY KEY, recipient_name TEXT, recipient_email TEXT, subject TEXT, sent_at TIMESTAMP, status TEXT)')
        c.execute('CREATE TABLE IF NOT EXISTS config (key TEXT UNIQUE, value TEXT)')
        c.execute('CREATE TABLE IF NOT EXISTS templates (id INTEGER PRIMARY KEY, name TEXT, subject TEXT, content TEXT)')
        conn.commit()
        conn.close()

    def create_layout(self):
        self.sidebar = tk.Frame(self.root, bg=ModernTheme.COLORS["sidebar_bg"], width=240)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)
        
        tk.Label(self.sidebar, text="✨ AuraMail", font=("Microsoft YaHei UI", 20, "bold"), 
                 bg=ModernTheme.COLORS["sidebar_bg"], fg=ModernTheme.COLORS["primary"]).pack(pady=30)
        
        self.nav_btns = {}
        self.pages = {}
        self.content_area = tk.Frame(self.root, bg=ModernTheme.COLORS["bg_app"])
        self.content_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        navs = [("send", "发送邮件", "🚀"), ("queue", "发送队列", "⏳"), 
                ("contacts", "通讯录", "👥"), ("templates", "模板管理", "📄"), ("history", "历史记录", "📜")]
        
        for key, txt, icon in navs:
            btn = SidebarButton(self.sidebar, txt, icon, lambda k=key: self.switch_page(k))
            btn.pack(fill=tk.X, pady=4, padx=8)
            self.nav_btns[key] = btn
            
            frame = tk.Frame(self.content_area, bg=ModernTheme.COLORS["bg_app"])
            self.pages[key] = frame
            getattr(self, f"ui_{key}")(frame)

        tk.Label(self.sidebar, text="v4.0 Ultimate", fg="#9ca3af", bg="white").pack(side=tk.BOTTOM, pady=10)
        self.switch_page("send")

    def switch_page(self, key):
        for k, btn in self.nav_btns.items():
            btn.set_selected(k == key)
        for k, frame in self.pages.items():
            if k == key: frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            else: frame.pack_forget()

    # ================= UI 构建区 =================

    def ui_send(self, parent):
        left = tk.Frame(parent, bg=ModernTheme.COLORS["bg_app"])
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        right = tk.Frame(parent, bg=ModernTheme.COLORS["bg_app"], width=380)
        right.pack(side=tk.RIGHT, fill=tk.Y)

        # 编辑器
        card_edit = ShadowElement(left, radius=15)
        card_edit.pack(fill=tk.BOTH, expand=True)
        inner = card_edit.inner_frame
        
        header = tk.Frame(inner, bg="white")
        header.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(header, text="模板:", bg="white", font=ModernTheme.FONTS["body"]).grid(row=0, column=0, sticky="w")
        self.combo_tmpl = ttk.Combobox(header, state="readonly", width=30); self.combo_tmpl.grid(row=0, column=1, padx=10, sticky="w")
        self.combo_tmpl.bind("<<ComboboxSelected>>", self.on_tmpl_select)
        tk.Button(header, text="🔄", command=self.refresh_tmpl_combo, relief="flat", bg="white").grid(row=0, column=2)
        
        tk.Label(header, text="主题:", bg="white", font=ModernTheme.FONTS["body"]).grid(row=1, column=0, sticky="w", pady=10)
        self.entry_subject = ModernEntry(header, font=("Microsoft YaHei", 12), width=50)
        self.entry_subject.grid(row=1, column=1, columnspan=2, sticky="ew", padx=10)
        
        toolbar = EditorToolbar(inner, None)
        toolbar.pack(fill=tk.X)
        
        frame_txt = tk.Frame(inner, bg="white", highlightthickness=1, highlightbackground="#e5e7eb")
        frame_txt.pack(fill=tk.BOTH, expand=True)
        
        self.txt_content = scrolledtext.ScrolledText(frame_txt, font=("Microsoft YaHei", 12), relief="flat", padx=10, pady=10)
        self.txt_content.pack(fill=tk.BOTH, expand=True)
        toolbar.text_widget = self.txt_content

        # 右侧面板
        card_conf = ShadowElement(right, radius=15)
        card_conf.pack(fill=tk.X, pady=(0, 15))
        c_inner = card_conf.inner_frame
        
        tk.Label(c_inner, text="⚙️ 发件配置", font=ModernTheme.FONTS["h2"], bg="white", fg=ModernTheme.COLORS["primary"]).pack(anchor="w", pady=(0, 10))
        tk.Label(c_inner, text="邮箱:", bg="white").pack(anchor="w")
        self.entry_email = ModernEntry(c_inner); self.entry_email.pack(fill=tk.X, pady=(0, 5))
        tk.Label(c_inner, text="授权码:", bg="white").pack(anchor="w")
        self.entry_pwd = ModernEntry(c_inner, show="*"); self.entry_pwd.pack(fill=tk.X, pady=(0, 5))
        tk.Label(c_inner, text="SMTP:", bg="white").pack(anchor="w")
        self.entry_smtp = ModernEntry(c_inner); self.entry_smtp.pack(fill=tk.X, pady=(0, 10))
        CapsuleButton(c_inner, text="保存配置", width=300, height=35, command=self.save_config).pack()

        # 附件
        card_attach = ShadowElement(right, radius=15)
        card_attach.pack(fill=tk.X, pady=(0, 15))
        a_inner = card_attach.inner_frame
        
        tk.Label(a_inner, text="📎 附件管理", font=ModernTheme.FONTS["h2"], bg="white", fg=ModernTheme.COLORS["primary"]).pack(anchor="w")
        
        list_frame = tk.Frame(a_inner, bg="#f9fafb", borderwidth=1, relief="solid")
        list_frame.configure(highlightbackground="#e5e7eb", highlightthickness=1)
        list_frame.pack(fill=tk.X, pady=5)
        
        self.list_attach = tk.Listbox(list_frame, height=4, relief="flat", bg="#f9fafb", 
                                     font=("Microsoft YaHei", 9), selectbackground="#e0e7ff", selectforeground=ModernTheme.COLORS["primary"])
        self.list_attach.pack(fill=tk.X, padx=2, pady=2)
        
        btn_row = tk.Frame(a_inner, bg="white")
        btn_row.pack(fill=tk.X, pady=5)
        tk.Button(btn_row, text="+ 添加", command=self.add_attachment, bg="#ecfccb", fg="#3f6212", relief="flat", width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="- 移除", command=self.remove_attachment, bg="#fee2e2", fg="#991b1b", relief="flat", width=10).pack(side=tk.RIGHT, padx=2)

        # 收件人
        card_act = ShadowElement(right, radius=15)
        card_act.pack(fill=tk.BOTH, expand=True)
        ac_inner = card_act.inner_frame
        
        tk.Label(ac_inner, text="👥 收件人", font=ModernTheme.FONTS["h2"], bg="white", fg=ModernTheme.COLORS["primary"]).pack(anchor="w")
        
        tools = tk.Frame(ac_inner, bg="white"); tools.pack(fill=tk.X, pady=5)
        tk.Button(tools, text="从通讯录选择 (带搜索)", command=self.open_contact_picker, relief="flat", bg=ModernTheme.COLORS["primary_light"], fg=ModernTheme.COLORS["primary"]).pack(fill=tk.X)
        
        self.list_rcpt = tk.Listbox(ac_inner, height=6, relief="flat", bg="#f9fafb", font=("Microsoft YaHei", 10))
        self.list_rcpt.pack(fill=tk.BOTH, expand=True, pady=5)
        tk.Button(ac_inner, text="清空列表", command=lambda: self.list_rcpt.delete(0, tk.END), relief="flat", fg="red", bg="white").pack(anchor="e")
        
        CapsuleButton(ac_inner, text="🚀 加入发送队列", width=300, height=45, command=self.add_to_queue).pack(side=tk.BOTTOM, pady=10)

    def ui_queue(self, parent):
        card = ShadowElement(parent, radius=15)
        card.pack(fill=tk.BOTH, expand=True)
        inner = card.inner_frame
        
        top = tk.Frame(inner, bg="white"); top.pack(fill=tk.X, pady=10)
        tk.Label(top, text="发送队列 (30s 缓冲)", font=ModernTheme.FONTS["h1"], bg="white").pack(side=tk.LEFT)
        tk.Button(top, text="⚡ 立即发送", command=self.force_send_all, bg=ModernTheme.COLORS["success"], fg="white", relief="flat", padx=15).pack(side=tk.RIGHT, padx=5)
        tk.Button(top, text="↩️ 撤回", command=self.withdraw_email, bg=ModernTheme.COLORS["danger"], fg="white", relief="flat", padx=15).pack(side=tk.RIGHT)
        
        self.tree_queue = ttk.Treeview(inner, columns=("ID", "收件人", "邮箱", "倒计时", "状态"), show="headings")
        for c in ("ID", "收件人", "邮箱", "倒计时", "状态"): self.tree_queue.heading(c, text=c)
        self.tree_queue.pack(fill=tk.BOTH, expand=True)

    def ui_contacts(self, parent):
        card = ShadowElement(parent, radius=15)
        card.pack(fill=tk.BOTH, expand=True)
        inner = card.inner_frame
        
        bar = tk.Frame(inner, bg="white"); bar.pack(fill=tk.X, pady=10)
        self.entry_search = ModernEntry(bar, width=30); self.entry_search.pack(side=tk.LEFT, padx=5)
        tk.Button(bar, text="搜索", command=self.search_contacts, bg=ModernTheme.COLORS["primary"], fg="white", relief="flat").pack(side=tk.LEFT)
        
        tk.Button(bar, text="导入Excel", command=self.import_excel, bg="#0ea5e9", fg="white", relief="flat").pack(side=tk.RIGHT, padx=5)
        tk.Button(bar, text="新建", command=self.add_contact_dialog, bg=ModernTheme.COLORS["success"], fg="white", relief="flat").pack(side=tk.RIGHT)
        
        self.tree_contacts = ttk.Treeview(inner, columns=("ID", "姓名", "邮箱", "职称", "院系"), show="headings")
        for c in ("ID", "姓名", "邮箱", "职称", "院系"): self.tree_contacts.heading(c, text=c)
        self.tree_contacts.pack(fill=tk.BOTH, expand=True)
        self.refresh_contacts()

    def ui_templates(self, parent):
        card = ShadowElement(parent, radius=15)
        card.pack(fill=tk.BOTH, expand=True)
        inner = card.inner_frame
        
        top = tk.Frame(inner, bg="white"); top.pack(fill=tk.X, pady=10)
        # 注意：这里调用的是新的对话框方法
        tk.Button(top, text="新建模板", command=self.new_template_dialog, bg=ModernTheme.COLORS["primary"], fg="white", relief="flat").pack(side=tk.LEFT)
        tk.Button(top, text="删除", command=self.del_template, bg="red", fg="white", relief="flat").pack(side=tk.RIGHT)
        
        self.tree_tmpl = ttk.Treeview(inner, columns=("ID", "名称", "主题"), show="headings")
        self.tree_tmpl.heading("ID", text="ID"); self.tree_tmpl.heading("名称", text="名称"); self.tree_tmpl.heading("主题", text="主题")
        self.tree_tmpl.pack(fill=tk.BOTH, expand=True)
        self.tree_tmpl.bind("<Double-1>", self.load_template_to_editor)
        self.refresh_tmpl_tree()

    def ui_history(self, parent):
        card = ShadowElement(parent, radius=15)
        card.pack(fill=tk.BOTH, expand=True)
        inner = card.inner_frame
        tk.Button(inner, text="清空历史", command=self.clear_history, bg="red", fg="white", relief="flat").pack(anchor="e", pady=10)
        self.tree_hist = ttk.Treeview(inner, columns=("收件人", "邮箱", "主题", "时间", "状态"), show="headings")
        for c in ("收件人", "邮箱", "主题", "时间", "状态"): self.tree_hist.heading(c, text=c)
        self.tree_hist.pack(fill=tk.BOTH, expand=True)
        self.refresh_history()

    # ================= 功能逻辑 =================

    def add_attachment(self):
        files = filedialog.askopenfilenames()
        for f in files:
            if f not in self.attachment_files:
                self.attachment_files.append(f)
                self.list_attach.insert(tk.END, os.path.basename(f))

    def remove_attachment(self):
        sel = self.list_attach.curselection()
        for index in reversed(sel):
            self.list_attach.delete(index)
            self.attachment_files.pop(index)

    def load_config(self):
        conn = sqlite3.connect(self.db_path)
        try:
            cfg = dict(conn.execute("SELECT key, value FROM config").fetchall())
            if "email" in cfg: self.entry_email.insert(0, cfg["email"])
            if "pwd" in cfg: self.entry_pwd.insert(0, cfg["pwd"])
            if "smtp" in cfg: self.entry_smtp.insert(0, cfg["smtp"])
        except: pass
        conn.close()

    def save_config(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("DELETE FROM config")
        conn.execute("INSERT INTO config VALUES (?,?)", ("email", self.entry_email.get()))
        conn.execute("INSERT INTO config VALUES (?,?)", ("pwd", self.entry_pwd.get()))
        conn.execute("INSERT INTO config VALUES (?,?)", ("smtp", self.entry_smtp.get()))
        conn.commit(); conn.close()
        messagebox.showinfo("成功", "配置已保存")

    def add_to_queue(self):
        rcpts = self.list_rcpt.get(0, tk.END)
        if not rcpts: return messagebox.showwarning("提示", "收件人列表为空")
        
        subject = self.entry_subject.get()
        body = self.txt_content.get("1.0", tk.END)
        sender = self.entry_email.get(); pwd = self.entry_pwd.get(); server = self.entry_smtp.get()
        
        conn = sqlite3.connect(self.db_path)
        send_time = time.time() + 30
        count = 0
        for r_str in rcpts:
            try:
                name = r_str.split('<')[0].strip()
                email = r_str.split('<')[1].strip('>')
                res = conn.execute("SELECT title, department FROM contacts WHERE email=?", (email,)).fetchone()
                title = res[0] if res else ""; dept = res[1] if res else ""
                final_body = body.replace("{姓名}", name).replace("{职称}", title).replace("{院系}", dept)
                eid = f"{int(time.time()*1000)}_{count}"
                self.pending_emails[eid] = {
                    "id": eid, "name": name, "email": email, "subject": subject, "content": final_body,
                    "sender": sender, "pwd": pwd, "server": server, 
                    "attachments": self.attachment_files.copy(),
                    "send_at": send_time, "status": "等待中"
                }
                count += 1
            except: pass
        conn.close()
        self.refresh_queue_ui(); self.switch_page("queue")
        messagebox.showinfo("成功", f"已添加 {count} 封邮件到队列")

    def start_queue_worker(self):
        def worker():
            while True:
                time.sleep(1)
                now = time.time()
                to_send = []
                for eid, data in self.pending_emails.items():
                    if data["status"] == "等待中":
                        rem = int(data["send_at"] - now)
                        if rem <= 0: to_send.append(eid)
                for eid in to_send:
                    self.pending_emails[eid]["status"] = "发送中"
                    self.send_queue.put(self.pending_emails[eid])
                self.root.after(0, self.refresh_queue_ui)

        def sender():
            while True:
                data = self.send_queue.get()
                self._send_mail(data)
                self.send_queue.task_done()

        threading.Thread(target=worker, daemon=True).start()
        threading.Thread(target=sender, daemon=True).start()

    def _send_mail(self, data):
        try:
            msg = MIMEMultipart()
            msg['From'] = data['sender']
            msg['To'] = data['email']
            msg['Subject'] = data['subject']
            msg.attach(MIMEText(data['content'], 'plain', 'utf-8'))
            for fpath in data['attachments']:
                with open(fpath, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream'); part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(fpath)}"')
                msg.attach(part)
            
            if "qq.com" in data['server']: s = smtplib.SMTP_SSL(data['server'], 465)
            elif "office365" in data['server']: s = smtplib.SMTP(data['server'], 587); s.starttls()
            else: s = smtplib.SMTP_SSL(data['server'], 465)
            s.login(data['sender'], data['pwd']); s.sendmail(data['sender'], data['email'], msg.as_string()); s.quit()
            del self.pending_emails[data['id']]
            self._log_history(data, "成功")
        except Exception as e:
            self.pending_emails[data['id']]["status"] = "失败"
            self._log_history(data, f"失败: {e}")

    def _log_history(self, data, status):
        conn = sqlite3.connect(self.db_path)
        conn.execute("INSERT INTO history (recipient_name, recipient_email, subject, sent_at, status) VALUES (?,?,?,?,?)",
                    (data['name'], data['email'], data['subject'], datetime.datetime.now(), status))
        conn.commit(); conn.close()

    def refresh_queue_ui(self):
        for i in self.tree_queue.get_children(): self.tree_queue.delete(i)
        now = time.time()
        for eid, d in self.pending_emails.items():
            rem = max(0, int(d["send_at"] - now)) if d["status"] == "等待中" else "-"
            self.tree_queue.insert("", tk.END, values=(eid, d["name"], d["email"], f"{rem}s", d["status"]))

    def force_send_all(self):
        for eid in self.pending_emails:
            if self.pending_emails[eid]["status"] == "等待中": self.pending_emails[eid]["send_at"] = time.time()

    def withdraw_email(self):
        sel = self.tree_queue.selection()
        if sel:
            eid = self.tree_queue.item(sel[0])['values'][0]
            if str(eid) in self.pending_emails: del self.pending_emails[str(eid)]; self.refresh_queue_ui()

    # =================== 关键更新：带搜索的联系人选择器 ===================
    def open_contact_picker(self):
        """打开带搜索功能的联系人选择器"""
        top = tk.Toplevel(self.root)
        top.title("选择联系人")
        top.geometry("500x600")
        top.configure(bg="white")
        
        # 1. 顶部搜索区
        search_frame = tk.Frame(top, bg="white", pady=10)
        search_frame.pack(fill=tk.X, padx=10)
        
        tk.Label(search_frame, text="🔍", bg="white", font=ModernTheme.FONTS["icon"]).pack(side=tk.LEFT)
        search_entry = ModernEntry(search_frame, width=30)
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 2. 列表区
        tree = ttk.Treeview(top, columns=("n", "e", "t"), show="headings", selectmode="extended")
        tree.heading("n", text="姓名"); tree.column("n", width=100)
        tree.heading("e", text="邮箱"); tree.column("e", width=200)
        tree.heading("t", text="职称"); tree.column("t", width=100)
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 3. 数据加载与搜索逻辑
        def load_data(query=""):
            for i in tree.get_children(): tree.delete(i)
            conn = sqlite3.connect(self.db_path)
            if query:
                sql = "SELECT name, email, title FROM contacts WHERE name LIKE ? OR email LIKE ?"
                params = (f"%{query}%", f"%{query}%")
                rows = conn.execute(sql, params).fetchall()
            else:
                rows = conn.execute("SELECT name, email, title FROM contacts").fetchall()
            conn.close()
            for r in rows: tree.insert("", tk.END, values=r)

        load_data() # 初始加载
        
        # 绑定搜索事件
        search_btn = tk.Button(search_frame, text="搜索", command=lambda: load_data(search_entry.get()), 
                             bg=ModernTheme.COLORS["primary"], fg="white", relief="flat")
        search_btn.pack(side=tk.LEFT, padx=5)
        search_entry.bind("<Return>", lambda e: load_data(search_entry.get()))

        # 4. 底部确认按钮
        def add_selected():
            for i in tree.selection():
                v = tree.item(i)['values']
                self.list_rcpt.insert(tk.END, f"{v[0]} <{v[1]}>")
            top.destroy()
            
        btn_frame = tk.Frame(top, bg="white", pady=10)
        btn_frame.pack(fill=tk.X)
        CapsuleButton(btn_frame, text="添加选中联系人", width=200, command=add_selected).pack()

    # =================== 关键更新：带富文本编辑器的模板弹窗 ===================
    def new_template_dialog(self):
        """新建模板对话框 - 包含文本编辑器"""
        top = tk.Toplevel(self.root)
        top.title("新建/编辑模板")
        top.geometry("700x600")
        top.configure(bg=ModernTheme.COLORS["bg_app"])
        
        # 使用卡片容器
        card = ShadowElement(top, radius=15)
        card.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        inner = card.inner_frame
        
        # 标题栏
        tk.Label(inner, text="模板名称:", bg="white").pack(anchor="w")
        entry_name = ModernEntry(inner, width=40); entry_name.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(inner, text="邮件主题:", bg="white").pack(anchor="w")
        entry_subject = ModernEntry(inner, width=40); entry_subject.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(inner, text="正文内容:", bg="white").pack(anchor="w")
        
        # 引入工具栏
        editor_frame = tk.Frame(inner, bg="white", highlightthickness=1, highlightbackground="#e5e7eb")
        editor_frame.pack(fill=tk.BOTH, expand=True)
        
        # 先创建文本框
        txt = scrolledtext.ScrolledText(editor_frame, font=("Microsoft YaHei", 12), relief="flat", padx=10, pady=10, height=10)
        
        # 再创建工具栏并绑定
        toolbar = EditorToolbar(editor_frame, txt)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        txt.pack(fill=tk.BOTH, expand=True)
        
        # 保存逻辑
        def save():
            name = entry_name.get()
            subj = entry_subject.get()
            content = txt.get("1.0", tk.END)
            
            if not name: return messagebox.showwarning("提示", "请输入模板名称")
            
            conn = sqlite3.connect(self.db_path)
            conn.execute("INSERT INTO templates (name, subject, content) VALUES (?,?,?)", (name, subj, content))
            conn.commit()
            conn.close()
            self.refresh_tmpl_tree()
            messagebox.showinfo("成功", "模板已保存")
            top.destroy()
            
        tk.Button(inner, text="💾 保存模板", command=save, bg=ModernTheme.COLORS["primary"], fg="white", 
                 font=("Microsoft YaHei", 11, "bold"), relief="flat", pady=8).pack(fill=tk.X, pady=15)

    # 辅助功能
    def refresh_contacts(self):
        for i in self.tree_contacts.get_children(): self.tree_contacts.delete(i)
        conn = sqlite3.connect(self.db_path)
        for r in conn.execute("SELECT id, name, email, title, department FROM contacts"): self.tree_contacts.insert("", tk.END, values=r)
        conn.close()

    def add_contact_dialog(self):
        d = tk.Toplevel(self.root); d.geometry("300x250")
        f = {}
        for k in ["姓名","邮箱","职称","院系"]:
            tk.Label(d, text=k).pack(); e = tk.Entry(d); e.pack(); f[k] = e
        def s():
            conn = sqlite3.connect(self.db_path)
            conn.execute("INSERT INTO contacts (name,email,title,department) VALUES (?,?,?,?)", 
                        (f["姓名"].get(), f["邮箱"].get(), f["职称"].get(), f["院系"].get()))
            conn.commit(); conn.close(); self.refresh_contacts(); d.destroy()
        tk.Button(d, text="保存", command=s).pack(pady=10)

    def import_excel(self):
        if not EXCEL_SUPPORT: return
        fn = filedialog.askopenfilename()
        if fn:
            wb = openpyxl.load_workbook(fn); ws = wb.active
            conn = sqlite3.connect(self.db_path)
            for r in ws.iter_rows(min_row=2, values_only=True):
                if r[0] and r[1]: 
                    conn.execute("INSERT INTO contacts (name,email,title,department) VALUES (?,?,?,?)", 
                                (r[0], r[1], r[2] if len(r)>2 else "", r[3] if len(r)>3 else ""))
            conn.commit(); conn.close(); self.refresh_contacts()

    def search_contacts(self):
        q = self.entry_search.get()
        for i in self.tree_contacts.get_children(): self.tree_contacts.delete(i)
        conn = sqlite3.connect(self.db_path)
        for r in conn.execute("SELECT * FROM contacts WHERE name LIKE ? OR email LIKE ?", (f"%{q}%", f"%{q}%")):
            self.tree_contacts.insert("", tk.END, values=r)
        conn.close()

    def refresh_tmpl_tree(self):
        for i in self.tree_tmpl.get_children(): self.tree_tmpl.delete(i)
        conn = sqlite3.connect(self.db_path)
        for r in conn.execute("SELECT id, name, subject FROM templates"): self.tree_tmpl.insert("", tk.END, values=r)
        conn.close()
    
    def refresh_tmpl_combo(self):
        conn = sqlite3.connect(self.db_path)
        self.combo_tmpl['values'] = [r[0] for r in conn.execute("SELECT name FROM templates")]
        conn.close()

    def on_tmpl_select(self, e):
        n = self.combo_tmpl.get()
        conn = sqlite3.connect(self.db_path)
        r = conn.execute("SELECT subject, content FROM templates WHERE name=?", (n,)).fetchone()
        conn.close()
        if r:
            self.entry_subject.delete(0, tk.END); self.entry_subject.insert(0, r[0])
            self.txt_content.delete("1.0", tk.END); self.txt_content.insert(tk.END, r[1])

    def del_template(self):
        sel = self.tree_tmpl.selection()
        if sel:
            tid = self.tree_tmpl.item(sel[0])['values'][0]
            conn = sqlite3.connect(self.db_path)
            conn.execute("DELETE FROM templates WHERE id=?", (tid,))
            conn.commit(); conn.close(); self.refresh_tmpl_tree()

    def load_template_to_editor(self, e):
        sel = self.tree_tmpl.selection()
        if sel:
            n = self.tree_tmpl.item(sel[0])['values'][1]
            self.switch_page("send")
            self.combo_tmpl.set(n)
            self.on_tmpl_select(None)

    def refresh_history(self):
        for i in self.tree_hist.get_children(): self.tree_hist.delete(i)
        conn = sqlite3.connect(self.db_path)
        for r in conn.execute("SELECT recipient_name, recipient_email, subject, sent_at, status FROM history ORDER BY id DESC"):
            self.tree_hist.insert("", tk.END, values=r)
        conn.close()
    
    def clear_history(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("DELETE FROM history"); conn.commit(); conn.close()
        self.refresh_history()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSender(root)
    root.mainloop()