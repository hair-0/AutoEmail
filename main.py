import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import sqlite3
from datetime import datetime
import os
import csv
import threading
import queue
import time

# 尝试导入openpyxl用于Excel支持
try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

class EmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("邮件发送工具 v2.3")
        self.root.geometry("1280x880")
        
        # 设置窗口图标
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'icon.png')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            # 如果icon.png不存在，使用默认图标
            pass
        
        # 启用高DPI支持，解决字体模糊问题
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
        
        # 设置字体
        self.setup_fonts()
        
        # 检查Excel支持
        if not EXCEL_SUPPORT:
            print("警告: 未安装openpyxl库，Excel导入功能不可用")
        
        # 发送队列和撤回功能
        self.send_queue = queue.Queue()
        self.sending = False
        self.sending_threads = []
        self.pending_emails = {}  # 存储待发送的邮件，支持撤回
        
        # 初始化数据库
        self.init_database()
        
        # 创建界面
        self.create_widgets()
        
        # 加载配置
        self.load_config()
        
        # 启动队列处理线程
        self.start_queue_processor()
    
    def setup_fonts(self):
        """设置字体"""
        # 使用宋体，符合中文显示习惯
        import platform
        system = platform.system()
        
        if system == "Windows":
            # 使用宋体，增加字号确保清晰
            self.title_font = ("SimSun", 14, "bold")
            self.normal_font = ("SimSun", 12)
            self.small_font = ("SimSun", 11)
            self.header_font = ("SimSun", 12, "bold")
        else:
            self.title_font = ("Arial", 14, "bold")
            self.normal_font = ("Arial", 12)
            self.small_font = ("Arial", 11)
            self.header_font = ("Arial", 12, "bold")
        
        # 配置ttk样式，优化视觉效果
        style = ttk.Style()
        
        # 使用更现代的主题
        try:
            style.theme_use('clam')
        except:
            pass
        
        # 配置标签样式 - 现代白色主题
        style.configure("Title.TLabel", font=self.title_font, foreground="#2c3e50")
        style.configure("Normal.TLabel", font=self.normal_font, foreground="#34495e")
        style.configure("Header.TLabel", font=self.header_font, foreground="#3498db")
        
        # 配置按钮样式 - 现代白色主题
        style.configure("TButton", 
                       font=self.normal_font, 
                       padding=(12, 6),
                       relief="flat",
                       borderwidth=1)
        style.map("TButton", 
                  background=[("active", "#e8f4fd"), ("pressed", "#d6eaf8")])
        
        # 配置LabelFrame样式 - 现代白色主题
        style.configure("TLabelframe", 
                        padding=15,
                        borderwidth=1,
                        relief="solid")
        style.configure("TLabelframe.Label", 
                        font=self.title_font, 
                        foreground="#2c3e50",
                        background="#f8f9fa")
        
        # 配置Notebook样式 - 现代白色主题
        style.configure("TNotebook", 
                        padding=[10, 5], 
                        borderwidth=1)
        style.configure("TNotebook.Tab", 
                        padding=[14, 10], 
                        font=self.normal_font)
        style.map("TNotebook.Tab", 
                  background=[("selected", "#3498db")],
                  foreground=[("selected", "#ffffff")])
        
        # 配置Treeview样式 - 现代白色主题
        style.configure("Treeview", 
                        font=self.normal_font, 
                        rowheight=30,
                        borderwidth=1,
                        relief="solid")
        style.configure("Treeview.Heading", 
                        font=self.header_font, 
                        foreground="#2c3e50",
                        relief="flat")
        style.map("Treeview", 
                  background=[("selected", "#3498db")],
                  foreground=[("selected", "#ffffff")])
        
        # 配置Entry样式 - 现代白色主题
        style.configure("TEntry", 
                        font=self.normal_font,
                        fieldbackground="white",
                        borderwidth=1,
                        relief="solid",
                        padding=5)
        style.map("TEntry",
                  fieldbackground=[("focus", "white")],
                  bordercolor=[("focus", "#3498db")])
        
        # 配置Combobox样式 - 现代白色主题
        style.configure("TCombobox", 
                        font=self.normal_font,
                        fieldbackground="white",
                        borderwidth=1,
                        relief="solid",
                        padding=5)
        style.map("TCombobox",
                  fieldbackground=[("readonly", "white")],
                  selectbackground=[("readonly", "#3498db")],
                  selectforeground=[("readonly", "#ffffff")])
    
    def init_database(self):
        """初始化数据库"""
        self.db_path = os.path.join(os.path.dirname(__file__), 'email_data.db')
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建通讯录表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT NOT NULL,
                title TEXT,
                department TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 创建历史记录表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                recipient_name TEXT,
                recipient_email TEXT,
                subject TEXT,
                content TEXT,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT
            )
        ''')
        
        # 创建配置表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS config (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                key TEXT UNIQUE,
                value TEXT
            )
        ''')
        
        # 创建邮件模板表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                subject TEXT,
                content TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def create_widgets(self):
        """创建界面组件"""
        # 创建Notebook（标签页）
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # 设置窗口背景色
        self.root.configure(bg="#f8f9fa")
        
        # 发送邮件标签页
        self.send_frame = ttk.Frame(notebook)
        notebook.add(self.send_frame, text="发送邮件")
        self.create_send_tab()
        
        # 待发送队列标签页（撤回功能）
        self.queue_frame = ttk.Frame(notebook)
        notebook.add(self.queue_frame, text="发送队列（可撤回）")
        self.create_queue_tab()
        
        # 模板管理标签页
        self.templates_frame = ttk.Frame(notebook)
        notebook.add(self.templates_frame, text="邮件模板")
        self.create_templates_tab()
        
        # 通讯录标签页
        self.contacts_frame = ttk.Frame(notebook)
        notebook.add(self.contacts_frame, text="通讯录")
        self.create_contacts_tab()
        
        # 历史记录标签页
        self.history_frame = ttk.Frame(notebook)
        notebook.add(self.history_frame, text="历史记录")
        self.create_history_tab()
    
    def create_send_tab(self):
        """创建发送邮件标签页"""
        # 邮箱设置 - 使用更大的padding避免标题重叠
        settings_frame = ttk.LabelFrame(self.send_frame, text="邮箱设置")
        settings_frame.pack(fill=tk.X, padx=12, pady=10)
        
        ttk.Label(settings_frame, text="发件邮箱:", style="Normal.TLabel").grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
        self.sender_email = ttk.Entry(settings_frame, width=34, font=self.normal_font)
        self.sender_email.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W)
        
        ttk.Label(settings_frame, text="授权码:", style="Normal.TLabel").grid(row=0, column=2, padx=10, pady=8, sticky=tk.W)
        self.sender_password = ttk.Entry(settings_frame, width=30, show="*", font=self.normal_font)
        self.sender_password.grid(row=0, column=3, padx=10, pady=8, sticky=tk.W)
        
        ttk.Label(settings_frame, text="SMTP服务器:", style="Normal.TLabel").grid(row=1, column=0, padx=10, pady=8, sticky=tk.W)
        self.smtp_server = ttk.Entry(settings_frame, width=20, font=self.normal_font)
        self.smtp_server.grid(row=1, column=1, padx=10, pady=8, sticky=tk.W)
        self.smtp_server.insert(0, "smtp.163.com")
        
        ttk.Label(settings_frame, text="端口:", style="Normal.TLabel").grid(row=1, column=2, padx=10, pady=8, sticky=tk.W)
        self.smtp_port = ttk.Entry(settings_frame, width=14, font=self.normal_font)
        self.smtp_port.grid(row=1, column=3, padx=10, pady=8, sticky=tk.W)
        self.smtp_port.insert(0, "465")
        
        ttk.Button(settings_frame, text="保存配置", command=self.save_config, width=14).grid(row=0, column=4, rowspan=2, padx=12, pady=8)
        
        # 邮件设置
        email_frame = ttk.LabelFrame(self.send_frame, text="邮件内容")
        email_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)
        
        ttk.Label(email_frame, text="邮件模板:", style="Normal.TLabel").grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
        self.template_combobox = ttk.Combobox(email_frame, width=34, state="readonly", font=self.normal_font)
        self.template_combobox.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W)
        self.template_combobox.bind('<<ComboboxSelected>>', self.on_template_select)
        
        ttk.Button(email_frame, text="刷新模板", command=self.refresh_templates, width=10).grid(row=0, column=2, padx=10, pady=8)
        
        ttk.Label(email_frame, text="主题:", style="Normal.TLabel").grid(row=1, column=0, padx=10, pady=8, sticky=tk.W)
        self.subject = ttk.Entry(email_frame, width=90, font=self.normal_font)
        self.subject.grid(row=1, column=1, columnspan=2, padx=10, pady=8, sticky=tk.W)
        
        # 收件人选择
        ttk.Label(email_frame, text="收件人:", style="Normal.TLabel").grid(row=2, column=0, padx=10, pady=8, sticky=tk.NW)
        
        # 创建收件人列表框容器，添加滚动条
        recipient_container = ttk.Frame(email_frame)
        recipient_container.grid(row=2, column=1, padx=10, pady=8, sticky=tk.W)
        
        recipient_scroll = ttk.Scrollbar(recipient_container, orient=tk.VERTICAL)
        self.recipient_listbox = tk.Listbox(recipient_container, height=5, selectmode=tk.MULTIPLE, 
                                           font=self.normal_font, yscrollcommand=recipient_scroll.set,
                                           borderwidth=1, relief="solid", highlightthickness=0)
        recipient_scroll.config(command=self.recipient_listbox.yview)
        self.recipient_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        recipient_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        recipient_btn_frame = ttk.Frame(email_frame)
        recipient_btn_frame.grid(row=2, column=2, padx=10, pady=8, sticky=tk.N)
        ttk.Button(recipient_btn_frame, text="从通讯录添加", command=self.load_contacts_to_listbox, width=14).pack(pady=4)
        ttk.Button(recipient_btn_frame, text="移除选中", command=self.remove_selected_recipients, width=14).pack(pady=4)
        ttk.Button(recipient_btn_frame, text="清空列表", command=self.clear_recipients, width=14).pack(pady=4)
        
        # 附件选择
        ttk.Label(email_frame, text="附件:", style="Normal.TLabel").grid(row=3, column=0, padx=10, pady=8, sticky=tk.NW)
        
        # 创建附件列表框容器，添加滚动条
        attachment_container = ttk.Frame(email_frame)
        attachment_container.grid(row=3, column=1, padx=10, pady=8, sticky=tk.W)
        
        attachment_scroll = ttk.Scrollbar(attachment_container, orient=tk.VERTICAL)
        self.attachments_listbox = tk.Listbox(attachment_container, height=3, selectmode=tk.SINGLE, 
                                             font=self.small_font, yscrollcommand=attachment_scroll.set,
                                             borderwidth=1, relief="solid", highlightthickness=0)
        attachment_scroll.config(command=self.attachments_listbox.yview)
        self.attachments_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        attachment_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        attachment_btn_frame = ttk.Frame(email_frame)
        attachment_btn_frame.grid(row=3, column=2, padx=10, pady=8, sticky=tk.N)
        ttk.Button(attachment_btn_frame, text="添加附件", command=self.add_attachment, width=14).pack(pady=4)
        ttk.Button(attachment_btn_frame, text="移除附件", command=self.remove_attachment, width=14).pack(pady=4)
        
        self.attachment_files = []  # 存储附件文件路径
        
        # 邮件模板
        ttk.Label(email_frame, text="邮件内容 (可使用占位符: {姓名}, {职称}, {院系}):", 
                 style="Normal.TLabel").grid(row=4, column=0, padx=10, pady=8, sticky=tk.NW)
        self.email_content = scrolledtext.ScrolledText(email_frame, width=90, height=15, font=self.normal_font,
                                                       wrap=tk.WORD, padx=10, pady=10,
                                                       borderwidth=1, relief="solid")
        self.email_content.grid(row=4, column=1, columnspan=2, padx=10, pady=8, sticky=tk.NSEW)
        
        # 默认模板
        default_template = """尊敬的{姓名}{职称}老师：

您好！

我们诚挚地邀请您担任本次[比赛名称]的评委。本次比赛旨在[比赛目的]，将于[时间]在[地点]举行。

作为[专业领域]的专家，您丰富的经验和深厚的学识将为本次比赛提供宝贵的指导和支持。您的参与将确保比赛的公平性和专业性。

如果您能拨冗担任评委，请您回复本邮件确认。我们将随后与您详细沟通具体事宜。

感谢您的时间和支持！

此致
敬礼！

[您的姓名]
[您的联系方式]
[日期]"""
        self.email_content.insert(tk.END, default_template)
        
        # 刷新模板列表
        self.refresh_templates()
        
        # 发送按钮
        send_btn_frame = ttk.Frame(self.send_frame)
        send_btn_frame.pack(fill=tk.X, padx=12, pady=14)
        ttk.Button(send_btn_frame, text="加入发送队列（可撤回）", command=self.add_to_queue, width=22).pack(side=tk.RIGHT, padx=10)
        ttk.Button(send_btn_frame, text="立即发送", command=self.send_emails, width=14).pack(side=tk.RIGHT, padx=10)
        ttk.Button(send_btn_frame, text="预览邮件", command=self.preview_email, width=14).pack(side=tk.RIGHT, padx=10)
        
        # 进度条
        self.progress = ttk.Progressbar(self.send_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, padx=12, pady=8)
        self.progress['value'] = 0
        self.progress.pack_forget()  # 初始隐藏
    
    def create_queue_tab(self):
        """创建发送队列标签页（撤回功能）"""
        # 说明标签
        info_label = ttk.Label(self.queue_frame, 
            text="提示：邮件加入队列后，会有30秒的延迟，期间可以点击'撤回'取消发送。\n30秒后会自动开始发送。", 
            style="Normal.TLabel", wraplength=1180)
        info_label.pack(fill=tk.X, padx=12, pady=8)
        
        # 操作按钮
        btn_frame = ttk.Frame(self.queue_frame)
        btn_frame.pack(fill=tk.X, padx=12, pady=8)
        
        ttk.Button(btn_frame, text="立即开始发送", command=self.start_queue_sending, width=14).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="撤回选中", command=self.withdraw_email, width=12).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="清空队列", command=self.clear_queue, width=12).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="刷新队列", command=self.refresh_queue, width=12).pack(side=tk.LEFT, padx=8)
        
        # 队列列表
        list_frame = ttk.LabelFrame(self.queue_frame, text="待发送队列")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        
        columns = ("id", "recipient_name", "recipient_email", "subject", "delay_time", "status")
        self.queue_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=22)
        
        self.queue_tree.heading("id", text="ID")
        self.queue_tree.heading("recipient_name", text="收件人")
        self.queue_tree.heading("recipient_email", text="邮箱")
        self.queue_tree.heading("subject", text="主题")
        self.queue_tree.heading("delay_time", text="剩余时间（秒）")
        self.queue_tree.heading("status", text="状态")
        
        self.queue_tree.column("id", width=60)
        self.queue_tree.column("recipient_name", width=130)
        self.queue_tree.column("recipient_email", width=220)
        self.queue_tree.column("subject", width=280)
        self.queue_tree.column("delay_time", width=130)
        self.queue_tree.column("status", width=110)
        
        self.queue_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.queue_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.queue_tree.configure(yscrollcommand=scrollbar.set)
        
        # 刷新队列
        self.refresh_queue()
    
    def start_queue_processor(self):
        """启动队列处理器线程"""
        def processor():
            while True:
                try:
                    time.sleep(1)  # 每秒检查一次
                    
                    # 更新倒计时
                    now = time.time()
                    to_send = []
                    to_remove = []
                    
                    for email_id, email_data in self.pending_emails.items():
                        if email_data['status'] == '等待中':
                            remaining = email_data['send_time'] - now
                            if remaining <= 0:
                                to_send.append(email_id)
                            else:
                                # 更新界面显示
                                self.root.after(0, lambda eid=email_id, r=int(remaining): self.update_queue_delay(eid, r))
                    
                    # 发送到期的邮件
                    for email_id in to_send:
                        if email_id in self.pending_emails:
                            email_data = self.pending_emails[email_id]
                            email_data['status'] = '发送中'
                            self.root.after(0, lambda eid=email_id: self.update_queue_status(eid, '发送中'))
                            self.send_queue.put(email_data)
                    
                except Exception as e:
                    print(f"队列处理器错误: {e}")
        
        # 启动处理器线程
        processor_thread = threading.Thread(target=processor, daemon=True)
        processor_thread.start()
        
        # 启动发送线程
        sender_thread = threading.Thread(target=self.queue_sender, daemon=True)
        sender_thread.start()
    
    def queue_sender(self):
        """队列发送线程"""
        while True:
            try:
                email_data = self.send_queue.get(timeout=1)
                self.send_single_email(email_data)
                self.send_queue.task_done()
            except queue.Empty:
                continue
            except Exception as e:
                print(f"发送错误: {e}")
    
    def add_to_queue(self):
        """添加邮件到发送队列"""
        sender_email = self.sender_email.get().strip()
        sender_password = self.sender_password.get().strip()
        smtp_server = self.smtp_server.get().strip()
        smtp_port = self.smtp_port.get().strip()
        subject = self.subject.get().strip()
        content = self.email_content.get("1.0", tk.END).strip()
        
        if not sender_email or not sender_password:
            messagebox.showwarning("警告", "请先设置发件邮箱和授权码")
            return
        
        if not subject or not content:
            messagebox.showwarning("警告", "请填写邮件主题和内容")
            return
        
        recipients = self.recipient_listbox.get(0, tk.END)
        if not recipients:
            messagebox.showwarning("警告", "请至少添加一个收件人")
            return
        
        # 解析收件人
        recipient_list = []
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        for recipient in recipients:
            try:
                name = recipient.split('<')[0].strip()
                email = recipient.split('<')[1].rstrip('>')
                
                # 获取联系人的职称和院系
                cursor.execute("SELECT title, department FROM contacts WHERE email=?", (email,))
                result = cursor.fetchone()
                title = result[0] if result and result[0] else ""
                department = result[1] if result and result[1] else ""
                
                # 替换模板中的占位符
                personalized_content = content.replace("{姓名}", name)
                personalized_content = personalized_content.replace("{职称}", f"{title}" if title else "")
                personalized_content = personalized_content.replace("{院系}", f"{department}" if department else "")
                
                recipient_list.append({
                    'name': name,
                    'email': email,
                    'subject': subject,
                    'content': personalized_content,
                    'sender_email': sender_email,
                    'sender_password': sender_password,
                    'smtp_server': smtp_server,
                    'smtp_port': smtp_port,
                    'attachments': self.attachment_files.copy()
                })
            except:
                continue
        
        conn.close()
        
        # 添加到队列（30秒延迟）
        send_time = time.time() + 30  # 30秒后发送
        for i, recipient in enumerate(recipient_list):
            email_id = f"email_{int(time.time() * 1000)}_{i}"
            recipient['email_id'] = email_id
            recipient['send_time'] = send_time
            recipient['status'] = '等待中'
            self.pending_emails[email_id] = recipient
        
        self.refresh_queue()
        messagebox.showinfo("成功", f"已将 {len(recipient_list)} 封邮件加入发送队列\n将在30秒后自动发送，期间可以撤回")
    
    def send_single_email(self, email_data):
        """发送单封邮件"""
        email_id = email_data['email_id']
        
        if email_id not in self.pending_emails:
            return
        
        try:
            try:
                port = int(email_data['smtp_port'])
            except:
                port = 465
            
            server = smtplib.SMTP_SSL(email_data['smtp_server'], port)
            server.login(email_data['sender_email'], email_data['sender_password'])
            
            # 创建邮件
            msg = MIMEMultipart()
            msg['From'] = email_data['sender_email']
            msg['To'] = email_data['email']
            msg['Subject'] = email_data['subject']
            
            msg.attach(MIMEText(email_data['content'], 'plain', 'utf-8'))
            
            # 添加附件
            for attachment_file in email_data['attachments']:
                if os.path.exists(attachment_file):
                    with open(attachment_file, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename="{os.path.basename(attachment_file)}"'
                    )
                    msg.attach(part)
            
            # 发送邮件
            server.sendmail(email_data['sender_email'], email_data['email'], msg.as_string())
            
            # 更新状态
            email_data['status'] = '已发送'
            self.root.after(0, lambda eid=email_id: self.update_queue_status(eid, '已发送'))
            
            # 记录到历史
            self.save_history(
                email_data['name'], 
                email_data['email'], 
                email_data['subject'], 
                email_data['content'], 
                "成功"
            )
            
            server.quit()
            
            # 3秒后从队列中移除
            time.sleep(3)
            if email_id in self.pending_emails:
                del self.pending_emails[email_id]
                self.root.after(0, self.refresh_queue)
            
        except Exception as e:
            email_data['status'] = f'失败: {str(e)}'
            self.root.after(0, lambda eid=email_id, err=str(e): self.update_queue_status(eid, f'失败: {err}'))
            
            self.save_history(
                email_data['name'], 
                email_data['email'], 
                email_data['subject'], 
                email_data['content'], 
                f"失败: {str(e)}"
            )
    
    def withdraw_email(self):
        """撤回邮件"""
        selected = self.queue_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一封邮件")
            return
        
        item = self.queue_tree.item(selected[0])
        values = item['values']
        email_id = values[0]
        
        if email_id in self.pending_emails:
            status = self.pending_emails[email_id]['status']
            if status == '等待中':
                del self.pending_emails[email_id]
                self.refresh_queue()
                messagebox.showinfo("成功", "邮件已撤回")
            else:
                messagebox.showwarning("警告", f"邮件状态为'{status}'，无法撤回")
        else:
            messagebox.showwarning("警告", "邮件不存在或已发送")
    
    def start_queue_sending(self):
        """立即开始发送队列中的邮件"""
        for email_id, email_data in self.pending_emails.items():
            if email_data['status'] == '等待中':
                email_data['send_time'] = time.time()  # 立即发送
                email_data['status'] = '发送中'
        
        self.refresh_queue()
        messagebox.showinfo("提示", "队列中的邮件将立即开始发送")
    
    def clear_queue(self):
        """清空队列"""
        if not messagebox.askyesno("确认", "确定要清空所有待发送邮件吗？"):
            return
        
        self.pending_emails.clear()
        self.refresh_queue()
    
    def refresh_queue(self):
        """刷新队列显示"""
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        
        for email_id, email_data in self.pending_emails.items():
            remaining = int(email_data['send_time'] - time.time())
            if remaining < 0:
                remaining = 0
            
            self.queue_tree.insert("", tk.END, values=(
                email_id,
                email_data['name'],
                email_data['email'],
                email_data['subject'],
                remaining if email_data['status'] == '等待中' else '-',
                email_data['status']
            ))
    
    def update_queue_delay(self, email_id, remaining):
        """更新队列中的倒计时"""
        for item in self.queue_tree.get_children():
            values = self.queue_tree.item(item)['values']
            if values[0] == email_id and values[5] == '等待中':
                self.queue_tree.item(item, values=(values[0], values[1], values[2], values[3], remaining, values[5]))
                break
    
    def update_queue_status(self, email_id, status):
        """更新队列中的状态"""
        for item in self.queue_tree.get_children():
            values = self.queue_tree.item(item)['values']
            if values[0] == email_id:
                self.queue_tree.item(item, values=(values[0], values[1], values[2], values[3], '-', status))
                break
    
    def create_templates_tab(self):
        """创建邮件模板标签页"""
        # 添加/编辑区域
        add_frame = ttk.LabelFrame(self.templates_frame, text="添加/编辑模板")
        add_frame.pack(fill=tk.X, padx=12, pady=8)
        
        ttk.Label(add_frame, text="模板名称:", style="Normal.TLabel").grid(row=0, column=0, padx=8, pady=6, sticky=tk.W)
        self.template_name = ttk.Entry(add_frame, width=32, font=self.normal_font)
        self.template_name.grid(row=0, column=1, padx=8, pady=6, sticky=tk.W)
        
        ttk.Label(add_frame, text="邮件主题:", style="Normal.TLabel").grid(row=1, column=0, padx=8, pady=6, sticky=tk.W)
        self.template_subject = ttk.Entry(add_frame, width=65, font=self.normal_font)
        self.template_subject.grid(row=1, column=1, padx=8, pady=6, sticky=tk.W)
        
        ttk.Label(add_frame, text="邮件内容 (占位符: {姓名}, {职称}, {院系}):", 
                 style="Normal.TLabel").grid(row=2, column=0, padx=8, pady=6, sticky=tk.NW)
        self.template_content = scrolledtext.ScrolledText(add_frame, width=85, height=12, font=self.normal_font,
                                                         wrap=tk.WORD, padx=8, pady=8)
        self.template_content.grid(row=2, column=1, padx=8, pady=6, sticky=tk.W)
        
        btn_frame = ttk.Frame(add_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=8)
        ttk.Button(btn_frame, text="保存模板", command=self.save_template, width=12).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="更新模板", command=self.update_template, width=12).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="清空", command=self.clear_template_form, width=12).pack(side=tk.LEFT, padx=8)
        
        # 模板列表
        list_frame = ttk.LabelFrame(self.templates_frame, text="模板列表")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        
        columns = ("id", "name", "subject", "created_at")
        self.templates_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        self.templates_tree.heading("id", text="ID")
        self.templates_tree.heading("name", text="模板名称")
        self.templates_tree.heading("subject", text="主题")
        self.templates_tree.heading("created_at", text="创建时间")
        
        self.templates_tree.column("id", width=60)
        self.templates_tree.column("name", width=180)
        self.templates_tree.column("subject", width=280)
        self.templates_tree.column("created_at", width=170)
        
        self.templates_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.templates_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.templates_tree.configure(yscrollcommand=scrollbar.set)
        
        # 绑定选择事件
        self.templates_tree.bind('<<TreeviewSelect>>', self.on_template_tree_select)
        
        # 操作按钮
        template_btn_frame = ttk.Frame(self.templates_frame)
        template_btn_frame.pack(fill=tk.X, padx=12, pady=8)
        ttk.Button(template_btn_frame, text="删除选中", command=self.delete_template, width=12).pack(side=tk.LEFT, padx=8)
        ttk.Button(template_btn_frame, text="使用此模板", command=self.use_template, width=14).pack(side=tk.LEFT, padx=8)
        
        # 刷新模板列表
        self.refresh_templates_tree()
    
    def create_contacts_tab(self):
        """创建通讯录标签页"""
        # 添加/编辑区域
        add_frame = ttk.LabelFrame(self.contacts_frame, text="添加/编辑联系人")
        add_frame.pack(fill=tk.X, padx=12, pady=10)
        
        ttk.Label(add_frame, text="姓名:", style="Normal.TLabel").grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
        self.contact_name = ttk.Entry(add_frame, width=24, font=self.normal_font)
        self.contact_name.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W)
        
        ttk.Label(add_frame, text="邮箱:", style="Normal.TLabel").grid(row=0, column=2, padx=10, pady=8, sticky=tk.W)
        self.contact_email = ttk.Entry(add_frame, width=34, font=self.normal_font)
        self.contact_email.grid(row=0, column=3, padx=10, pady=8, sticky=tk.W)
        
        ttk.Label(add_frame, text="职称:", style="Normal.TLabel").grid(row=1, column=0, padx=10, pady=8, sticky=tk.W)
        self.contact_title = ttk.Entry(add_frame, width=24, font=self.normal_font)
        self.contact_title.grid(row=1, column=1, padx=10, pady=8, sticky=tk.W)
        
        ttk.Label(add_frame, text="院系:", style="Normal.TLabel").grid(row=1, column=2, padx=10, pady=8, sticky=tk.W)
        self.contact_department = ttk.Entry(add_frame, width=34, font=self.normal_font)
        self.contact_department.grid(row=1, column=3, padx=10, pady=8, sticky=tk.W)
        
        btn_frame = ttk.Frame(add_frame)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=10)
        ttk.Button(btn_frame, text="添加", command=self.add_contact, width=14).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="更新", command=self.update_contact, width=14).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="清空", command=self.clear_contact_form, width=14).pack(side=tk.LEFT, padx=10)
        
        # 搜索区域
        search_frame = ttk.Frame(self.contacts_frame)
        search_frame.pack(fill=tk.X, padx=12, pady=10)
        ttk.Label(search_frame, text="搜索:", style="Normal.TLabel").pack(side=tk.LEFT, padx=10)
        self.search_entry = ttk.Entry(search_frame, width=34, font=self.normal_font)
        self.search_entry.pack(side=tk.LEFT, padx=10)
        ttk.Button(search_frame, text="搜索", command=self.search_contacts, width=12).pack(side=tk.LEFT, padx=10)
        ttk.Button(search_frame, text="显示全部", command=self.refresh_contacts, width=12).pack(side=tk.LEFT, padx=10)
        
        # 联系人列表
        list_frame = ttk.LabelFrame(self.contacts_frame, text="联系人列表")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)
        
        # 创建Treeview
        columns = ("id", "name", "email", "title", "department")
        self.contacts_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=18)
        
        self.contacts_tree.heading("id", text="ID")
        self.contacts_tree.heading("name", text="姓名")
        self.contacts_tree.heading("email", text="邮箱")
        self.contacts_tree.heading("title", text="职称")
        self.contacts_tree.heading("department", text="院系")
        
        self.contacts_tree.column("id", width=60)
        self.contacts_tree.column("name", width=120)
        self.contacts_tree.column("email", width=220)
        self.contacts_tree.column("title", width=120)
        self.contacts_tree.column("department", width=180)
        
        self.contacts_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.contacts_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.contacts_tree.configure(yscrollcommand=scrollbar.set)
        
        # 绑定选择事件
        self.contacts_tree.bind('<<TreeviewSelect>>', self.on_contact_select)
        
        # 操作按钮
        contact_btn_frame = ttk.Frame(self.contacts_frame)
        contact_btn_frame.pack(fill=tk.X, padx=12, pady=10)
        ttk.Button(contact_btn_frame, text="删除选中", command=self.delete_contact, width=14).pack(side=tk.LEFT, padx=10)
        ttk.Button(contact_btn_frame, text="导出为CSV", command=self.export_contacts_csv, width=16).pack(side=tk.LEFT, padx=10)
        ttk.Button(contact_btn_frame, text="批量导入CSV", command=self.import_contacts_csv, width=16).pack(side=tk.LEFT, padx=10)
        
        # 添加Excel导入按钮
        if EXCEL_SUPPORT:
            ttk.Button(contact_btn_frame, text="批量导入Excel", command=self.import_contacts_excel, width=16).pack(side=tk.LEFT, padx=10)
        
        # 刷新联系人列表
        self.refresh_contacts()
    
    def create_history_tab(self):
        """创建历史记录标签页"""
        # 搜索区域
        search_frame = ttk.Frame(self.history_frame)
        search_frame.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(search_frame, text="搜索:", style="Normal.TLabel").pack(side=tk.LEFT, padx=8)
        self.history_search_entry = ttk.Entry(search_frame, width=32, font=self.normal_font)
        self.history_search_entry.pack(side=tk.LEFT, padx=8)
        ttk.Button(search_frame, text="搜索", command=self.search_history, width=10).pack(side=tk.LEFT, padx=8)
        ttk.Button(search_frame, text="显示全部", command=self.refresh_history, width=10).pack(side=tk.LEFT, padx=8)
        ttk.Button(search_frame, text="清空历史", command=self.clear_history, width=10).pack(side=tk.LEFT, padx=8)
        ttk.Button(search_frame, text="导出CSV", command=self.export_history_csv, width=12).pack(side=tk.LEFT, padx=8)
        
        # 历史记录列表
        list_frame = ttk.LabelFrame(self.history_frame, text="发送历史")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        
        columns = ("id", "recipient_name", "recipient_email", "subject", "sent_at", "status")
        self.history_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=22)
        
        self.history_tree.heading("id", text="ID")
        self.history_tree.heading("recipient_name", text="收件人")
        self.history_tree.heading("recipient_email", text="邮箱")
        self.history_tree.heading("subject", text="主题")
        self.history_tree.heading("sent_at", text="发送时间")
        self.history_tree.heading("status", text="状态")
        
        self.history_tree.column("id", width=60)
        self.history_tree.column("recipient_name", width=120)
        self.history_tree.column("recipient_email", width=220)
        self.history_tree.column("subject", width=250)
        self.history_tree.column("sent_at", width=170)
        self.history_tree.column("status", width=90)
        
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        
        # 双击查看详情
        self.history_tree.bind('<Double-1>', self.on_history_double_click)
        
        # 刷新历史记录
        self.refresh_history()
    
    def save_config(self):
        """保存配置"""
        email = self.sender_email.get().strip()
        password = self.sender_password.get().strip()
        smtp_server = self.smtp_server.get().strip()
        smtp_port = self.smtp_port.get().strip()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 删除旧配置
        cursor.execute("DELETE FROM config")
        
        # 保存新配置
        if email:
            cursor.execute("INSERT INTO config (key, value) VALUES (?, ?)", ("email", email))
        if password:
            cursor.execute("INSERT INTO config (key, value) VALUES (?, ?)", ("password", password))
        if smtp_server:
            cursor.execute("INSERT INTO config (key, value) VALUES (?, ?)", ("smtp_server", smtp_server))
        if smtp_port:
            cursor.execute("INSERT INTO config (key, value) VALUES (?, ?)", ("smtp_port", smtp_port))
        
        conn.commit()
        conn.close()
        
        messagebox.showinfo("成功", "配置已保存")
    
    def load_config(self):
        """加载配置"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT key, value FROM config")
        configs = cursor.fetchall()
        
        for key, value in configs:
            if key == "email":
                self.sender_email.insert(0, value)
            elif key == "password":
                self.sender_password.insert(0, value)
            elif key == "smtp_server":
                self.smtp_server.delete(0, tk.END)
                self.smtp_server.insert(0, value)
            elif key == "smtp_port":
                self.smtp_port.delete(0, tk.END)
                self.smtp_port.insert(0, value)
        
        conn.close()
    
    def add_attachment(self):
        """添加附件"""
        filename = filedialog.askopenfilename(
            title="选择附件",
            filetypes=[("所有文件", "*.*")]
        )
        if filename:
            self.attachment_files.append(filename)
            self.attachments_listbox.insert(tk.END, os.path.basename(filename))
    
    def remove_attachment(self):
        """移除附件"""
        selected = self.attachments_listbox.curselection()
        if selected:
            index = selected[0]
            self.attachments_listbox.delete(index)
            self.attachment_files.pop(index)
    
    def load_contacts_to_listbox(self):
        """从通讯录加载到收件人列表框"""
        self.recipient_listbox.delete(0, tk.END)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, name, email, title, department FROM contacts ORDER BY name")
        contacts = cursor.fetchall()
        
        for contact in contacts:
            # 格式: "姓名 <邮箱>" (ID在内部使用)
            self.recipient_listbox.insert(tk.END, f"{contact[1]} <{contact[2]}>")
        
        conn.close()
    
    def remove_selected_recipients(self):
        """移除选中的收件人"""
        selected = self.recipient_listbox.curselection()
        for index in reversed(selected):
            self.recipient_listbox.delete(index)
    
    def clear_recipients(self):
        """清空收件人列表"""
        self.recipient_listbox.delete(0, tk.END)
    
    def add_contact(self):
        """添加联系人"""
        name = self.contact_name.get().strip()
        email = self.contact_email.get().strip()
        title = self.contact_title.get().strip()
        department = self.contact_department.get().strip()
        
        if not name or not email:
            messagebox.showwarning("警告", "姓名和邮箱不能为空")
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute(
                "INSERT INTO contacts (name, email, title, department) VALUES (?, ?, ?, ?)",
                (name, email, title, department)
            )
            conn.commit()
            messagebox.showinfo("成功", "联系人已添加")
            self.refresh_contacts()
            self.clear_contact_form()
        except sqlite3.IntegrityError:
            messagebox.showerror("错误", "邮箱已存在")
        finally:
            conn.close()
    
    def update_contact(self):
        """更新联系人"""
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个联系人")
            return
        
        item = self.contacts_tree.item(selected[0])
        contact_id = item['values'][0]
        
        name = self.contact_name.get().strip()
        email = self.contact_email.get().strip()
        title = self.contact_title.get().strip()
        department = self.contact_department.get().strip()
        
        if not name or not email:
            messagebox.showwarning("警告", "姓名和邮箱不能为空")
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute(
                "UPDATE contacts SET name=?, email=?, title=?, department=? WHERE id=?",
                (name, email, title, department, contact_id)
            )
            conn.commit()
            messagebox.showinfo("成功", "联系人已更新")
            self.refresh_contacts()
            self.clear_contact_form()
        except sqlite3.IntegrityError:
            messagebox.showerror("错误", "邮箱已存在")
        finally:
            conn.close()
    
    def delete_contact(self):
        """删除联系人"""
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个联系人")
            return
        
        if not messagebox.askyesno("确认", "确定要删除选中的联系人吗？"):
            return
        
        item = self.contacts_tree.item(selected[0])
        contact_id = item['values'][0]
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM contacts WHERE id=?", (contact_id,))
        conn.commit()
        conn.close()
        
        self.refresh_contacts()
        self.clear_contact_form()
    
    def clear_contact_form(self):
        """清空联系人表单"""
        self.contact_name.delete(0, tk.END)
        self.contact_email.delete(0, tk.END)
        self.contact_title.delete(0, tk.END)
        self.contact_department.delete(0, tk.END)
    
    def on_contact_select(self, event):
        """联系人选择事件"""
        selected = self.contacts_tree.selection()
        if not selected:
            return
        
        item = self.contacts_tree.item(selected[0])
        values = item['values']
        
        if len(values) >= 5:
            self.contact_name.delete(0, tk.END)
            self.contact_name.insert(0, values[1])
            
            self.contact_email.delete(0, tk.END)
            self.contact_email.insert(0, values[2])
            
            self.contact_title.delete(0, tk.END)
            self.contact_title.insert(0, values[3])
            
            self.contact_department.delete(0, tk.END)
            self.contact_department.insert(0, values[4])
    
    def refresh_contacts(self):
        """刷新联系人列表"""
        for item in self.contacts_tree.get_children():
            self.contacts_tree.delete(item)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, name, email, title, department FROM contacts ORDER BY name")
        contacts = cursor.fetchall()
        
        for contact in contacts:
            self.contacts_tree.insert("", tk.END, values=contact)
        
        conn.close()
    
    def search_contacts(self):
        """搜索联系人"""
        keyword = self.search_entry.get().strip()
        
        for item in self.contacts_tree.get_children():
            self.contacts_tree.delete(item)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            "SELECT id, name, email, title, department FROM contacts WHERE name LIKE ? OR email LIKE ? OR title LIKE ? OR department LIKE ? ORDER BY name",
            (f"%{keyword}%", f"%{keyword}%", f"%{keyword}%", f"%{keyword}%")
        )
        contacts = cursor.fetchall()
        
        for contact in contacts:
            self.contacts_tree.insert("", tk.END, values=contact)
        
        conn.close()
    
    def export_contacts_csv(self):
        """导出联系人为CSV"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")],
            title="导出联系人"
        )
        
        if not filename:
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT name, email, title, department FROM contacts ORDER BY name")
        contacts = cursor.fetchall()
        
        with open(filename, 'w', encoding='utf-8-sig') as f:
            f.write("姓名,邮箱,职称,院系\n")
            for contact in contacts:
                f.write(f"{contact[0]},{contact[1]},{contact[2]},{contact[3]}\n")
        
        conn.close()
        
        messagebox.showinfo("成功", "联系人已导出")
    
    def import_contacts_csv(self):
        """从CSV批量导入联系人"""
        filename = filedialog.askopenfilename(
            title="选择CSV文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if not filename:
            return
        
        try:
            with open(filename, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                contacts = list(reader)
                
                if not contacts:
                    messagebox.showwarning("警告", "CSV文件为空")
                    return
                
                # 检查列名
                required_columns = ['姓名', '邮箱']
                if not all(col in contacts[0] for col in required_columns):
                    messagebox.showerror("错误", "CSV文件必须包含'姓名'和'邮箱'列")
                    return
                
                # 导入联系人
                self._import_contacts_list(contacts)
                
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {str(e)}")
    
    def import_contacts_excel(self):
        """从Excel批量导入联系人"""
        if not EXCEL_SUPPORT:
            messagebox.showerror("错误", "未安装openpyxl库，无法导入Excel文件\n请运行: pip install openpyxl")
            return
        
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not filename:
            return
        
        try:
            workbook = openpyxl.load_workbook(filename)
            sheet = workbook.active
            
            # 读取表头
            headers = [cell.value for cell in sheet[1]]
            
            # 检查必需列
            required_columns = ['姓名', '邮箱']
            if not all(col in headers for col in required_columns):
                messagebox.showerror("错误", f"Excel文件必须包含'姓名'和'邮箱'列\n当前列: {', '.join(headers)}")
                return
            
            # 读取数据
            contacts = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[0]:  # 跳过空行
                    continue
                
                contact = {}
                for i, header in enumerate(headers):
                    if i < len(row) and row[i]:
                        contact[header] = str(row[i]).strip()
                    else:
                        contact[header] = ''
                
                contacts.append(contact)
            
            if not contacts:
                messagebox.showwarning("警告", "Excel文件为空或无数据")
                return
            
            # 导入联系人
            self._import_contacts_list(contacts)
            workbook.close()
            
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {str(e)}")
    
    def _import_contacts_list(self, contacts):
        """导入联系人列表（内部方法）"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        success_count = 0
        error_count = 0
        
        for contact in contacts:
            name = contact.get('姓名', '').strip()
            email = contact.get('邮箱', '').strip()
            title = contact.get('职称', '').strip()
            department = contact.get('院系', '').strip()
            
            if not name or not email:
                error_count += 1
                continue
            
            try:
                cursor.execute(
                    "INSERT INTO contacts (name, email, title, department) VALUES (?, ?, ?, ?)",
                    (name, email, title, department)
                )
                success_count += 1
            except sqlite3.IntegrityError:
                error_count += 1
        
        conn.commit()
        conn.close()
        
        self.refresh_contacts()
        
        message = f"导入完成！\n成功: {success_count}\n失败: {error_count}"
        messagebox.showinfo("完成", message)
    
    def send_emails(self):
        """立即发送邮件（不经过队列）"""
        sender_email = self.sender_email.get().strip()
        sender_password = self.sender_password.get().strip()
        smtp_server = self.smtp_server.get().strip()
        smtp_port = self.smtp_port.get().strip()
        subject = self.subject.get().strip()
        content = self.email_content.get("1.0", tk.END).strip()
        
        if not sender_email or not sender_password:
            messagebox.showwarning("警告", "请先设置发件邮箱和授权码")
            return
        
        if not subject or not content:
            messagebox.showwarning("警告", "请填写邮件主题和内容")
            return
        
        recipients = self.recipient_listbox.get(0, tk.END)
        if not recipients:
            messagebox.showwarning("警告", "请至少添加一个收件人")
            return
        
        if not messagebox.askyesno("确认", f"确定要立即发送邮件给 {len(recipients)} 个收件人吗？\n此操作无法撤回！"):
            return
        
        # 在新线程中发送邮件，避免界面卡顿
        threading.Thread(target=self._send_emails_thread, args=(sender_email, sender_password, smtp_server, smtp_port, subject, content, recipients), daemon=True).start()
    
    def _send_emails_thread(self, sender_email, sender_password, smtp_server, smtp_port, subject, content, recipients):
        """发送邮件的线程函数"""
        # 显示进度条
        self.root.after(0, lambda: self.progress.pack(fill=tk.X, padx=10, pady=5))
        
        # 解析收件人
        recipient_list = []
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        for recipient in recipients:
            try:
                name = recipient.split('<')[0].strip()
                email = recipient.split('<')[1].rstrip('>')
                
                # 获取联系人的职称和院系
                cursor.execute("SELECT title, department FROM contacts WHERE email=?", (email,))
                result = cursor.fetchone()
                title = result[0] if result and result[0] else ""
                department = result[1] if result and result[1] else ""
                
                recipient_list.append((name, email, title, department))
            except:
                continue
        
        conn.close()
        
        if not recipient_list:
            self.root.after(0, lambda: messagebox.showerror("错误", "收件人格式错误"))
            return
        
        # 连接SMTP服务器
        try:
            try:
                port = int(smtp_port)
            except:
                port = 465
            
            server = smtplib.SMTP_SSL(smtp_server, port)
            server.login(sender_email, sender_password)
            
            success_count = 0
            fail_count = 0
            total = len(recipient_list)
            
            for i, (name, email, title, department) in enumerate(recipient_list):
                try:
                    # 替换模板中的占位符
                    personalized_content = content.replace("{姓名}", name)
                    personalized_content = personalized_content.replace("{职称}", f"{title}" if title else "")
                    personalized_content = personalized_content.replace("{院系}", f"{department}" if department else "")
                    
                    # 创建邮件
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = email
                    msg['Subject'] = subject
                    
                    msg.attach(MIMEText(personalized_content, 'plain', 'utf-8'))
                    
                    # 添加附件
                    for attachment_file in self.attachment_files:
                        if os.path.exists(attachment_file):
                            with open(attachment_file, 'rb') as f:
                                part = MIMEBase('application', 'octet-stream')
                                part.set_payload(f.read())
                            encoders.encode_base64(part)
                            part.add_header(
                                'Content-Disposition',
                                f'attachment; filename="{os.path.basename(attachment_file)}"'
                            )
                            msg.attach(part)
                    
                    # 发送邮件
                    server.sendmail(sender_email, email, msg.as_string())
                    
                    # 记录到历史
                    self.save_history(name, email, subject, personalized_content, "成功")
                    success_count += 1
                    
                except Exception as e:
                    self.save_history(name, email, subject, content, f"失败: {str(e)}")
                    fail_count += 1
                
                # 更新进度条
                progress = (i + 1) / total * 100
                self.root.after(0, lambda p=progress: self.progress.configure(value=p))
                self.root.after(0, lambda: self.root.update())
            
            server.quit()
            
            self.root.after(0, self.refresh_history)
            self.root.after(0, lambda: self.progress.pack_forget())
            
            message = f"发送完成！\n成功: {success_count}\n失败: {fail_count}"
            self.root.after(0, lambda: messagebox.showinfo("完成", message))
            
        except Exception as e:
            self.root.after(0, lambda: self.progress.pack_forget())
            self.root.after(0, lambda: messagebox.showerror("错误", f"发送失败: {str(e)}"))
    
    def preview_email(self):
        """预览邮件"""
        selected = self.recipient_listbox.curselection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个收件人")
            return
        
        recipient = self.recipient_listbox.get(selected[0])
        try:
            name = recipient.split('<')[0].strip()
            email = recipient.split('<')[1].rstrip('>')
            
            # 获取联系人的职称和院系
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT title, department FROM contacts WHERE email=?", (email,))
            result = cursor.fetchone()
            title = result[0] if result and result[0] else ""
            department = result[1] if result and result[1] else ""
            conn.close()
            
            content = self.email_content.get("1.0", tk.END).strip()
            personalized_content = content.replace("{姓名}", name)
            personalized_content = personalized_content.replace("{职称}", f"{title}" if title else "")
            personalized_content = personalized_content.replace("{院系}", f"{department}" if department else "")
            
            # 创建预览窗口
            preview_window = tk.Toplevel(self.root)
            preview_window.title("邮件预览")
            preview_window.geometry("650x550")
            
            preview_text = scrolledtext.ScrolledText(preview_window, width=75, height=28, font=self.normal_font,
                                                     wrap=tk.WORD, padx=10, pady=10)
            preview_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
            preview_text.insert(tk.END, personalized_content)
            preview_text.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {str(e)}")
    
    def save_history(self, recipient_name, recipient_email, subject, content, status):
        """保存发送历史"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT INTO history (recipient_name, recipient_email, subject, content, status) VALUES (?, ?, ?, ?, ?)",
            (recipient_name, recipient_email, subject, content, status)
        )
        
        conn.commit()
        conn.close()
    
    def refresh_history(self):
        """刷新历史记录"""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            "SELECT id, recipient_name, recipient_email, subject, sent_at, status FROM history ORDER BY sent_at DESC"
        )
        records = cursor.fetchall()
        
        for record in records:
            self.history_tree.insert("", tk.END, values=record)
        
        conn.close()
    
    def search_history(self):
        """搜索历史记录"""
        keyword = self.history_search_entry.get().strip()
        
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            "SELECT id, recipient_name, recipient_email, subject, sent_at, status FROM history WHERE recipient_name LIKE ? OR recipient_email LIKE ? OR subject LIKE ? ORDER BY sent_at DESC",
            (f"%{keyword}%", f"%{keyword}%", f"%{keyword}%")
        )
        records = cursor.fetchall()
        
        for record in records:
            self.history_tree.insert("", tk.END, values=record)
        
        conn.close()
    
    def export_history_csv(self):
        """导出历史记录为CSV"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")],
            title="导出历史记录"
        )
        
        if not filename:
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT recipient_name, recipient_email, subject, sent_at, status FROM history ORDER BY sent_at DESC")
        records = cursor.fetchall()
        
        with open(filename, 'w', encoding='utf-8-sig') as f:
            f.write("收件人,邮箱,主题,发送时间,状态\n")
            for record in records:
                f.write(f"{record[0]},{record[1]},{record[2]},{record[3]},{record[4]}\n")
        
        conn.close()
        
        messagebox.showinfo("成功", "历史记录已导出")
    
    def clear_history(self):
        """清空历史记录"""
        if not messagebox.askyesno("确认", "确定要清空所有历史记录吗？"):
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM history")
        conn.commit()
        conn.close()
        
        self.refresh_history()
    
    def on_history_double_click(self, event):
        """历史记录双击事件"""
        selected = self.history_tree.selection()
        if not selected:
            return
        
        item = self.history_tree.item(selected[0])
        values = item['values']
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT content FROM history WHERE id=?", (values[0],))
        result = cursor.fetchone()
        
        conn.close()
        
        if result:
            # 创建详情窗口
            detail_window = tk.Toplevel(self.root)
            detail_window.title("邮件详情")
            detail_window.geometry("650x550")
            
            detail_text = scrolledtext.ScrolledText(detail_window, width=75, height=28, font=self.normal_font,
                                                  wrap=tk.WORD, padx=10, pady=10)
            detail_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
            detail_text.insert(tk.END, result[0])
            detail_text.config(state=tk.DISABLED)
    
    # ========== 模板管理相关方法 ==========
    
    def save_template(self):
        """保存模板"""
        name = self.template_name.get().strip()
        subject = self.template_subject.get().strip()
        content = self.template_content.get("1.0", tk.END).strip()
        
        if not name:
            messagebox.showwarning("警告", "模板名称不能为空")
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute(
                "INSERT INTO templates (name, subject, content) VALUES (?, ?, ?)",
                (name, subject, content)
            )
            conn.commit()
            messagebox.showinfo("成功", "模板已保存")
            self.refresh_templates_tree()
            self.clear_template_form()
        except sqlite3.IntegrityError:
            messagebox.showerror("错误", "模板名称已存在")
        finally:
            conn.close()
    
    def update_template(self):
        """更新模板"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个模板")
            return
        
        item = self.templates_tree.item(selected[0])
        template_id = item['values'][0]
        
        name = self.template_name.get().strip()
        subject = self.template_subject.get().strip()
        content = self.template_content.get("1.0", tk.END).strip()
        
        if not name:
            messagebox.showwarning("警告", "模板名称不能为空")
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute(
                "UPDATE templates SET name=?, subject=?, content=? WHERE id=?",
                (name, subject, content, template_id)
            )
            conn.commit()
            messagebox.showinfo("成功", "模板已更新")
            self.refresh_templates_tree()
            self.clear_template_form()
        except sqlite3.IntegrityError:
            messagebox.showerror("错误", "模板名称已存在")
        finally:
            conn.close()
    
    def delete_template(self):
        """删除模板"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个模板")
            return
        
        if not messagebox.askyesno("确认", "确定要删除选中的模板吗？"):
            return
        
        item = self.templates_tree.item(selected[0])
        template_id = item['values'][0]
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM templates WHERE id=?", (template_id,))
        conn.commit()
        conn.close()
        
        self.refresh_templates_tree()
        self.clear_template_form()
    
    def clear_template_form(self):
        """清空模板表单"""
        self.template_name.delete(0, tk.END)
        self.template_subject.delete(0, tk.END)
        self.template_content.delete("1.0", tk.END)
    
    def on_template_tree_select(self, event):
        """模板选择事件"""
        selected = self.templates_tree.selection()
        if not selected:
            return
        
        item = self.templates_tree.item(selected[0])
        values = item['values']
        
        if len(values) >= 4:
            self.template_name.delete(0, tk.END)
            self.template_name.insert(0, values[1])
            
            self.template_subject.delete(0, tk.END)
            self.template_subject.insert(0, values[2])
            
            # 获取完整内容
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT content FROM templates WHERE id=?", (values[0],))
            result = cursor.fetchone()
            conn.close()
            
            if result:
                self.template_content.delete("1.0", tk.END)
                self.template_content.insert(tk.END, result[0])
    
    def refresh_templates_tree(self):
        """刷新模板列表"""
        for item in self.templates_tree.get_children():
            self.templates_tree.delete(item)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, name, subject, created_at FROM templates ORDER BY created_at DESC")
        templates = cursor.fetchall()
        
        for template in templates:
            self.templates_tree.insert("", tk.END, values=template)
        
        conn.close()
    
    def refresh_templates(self):
        """刷新模板下拉框"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT name FROM templates ORDER BY created_at DESC")
        templates = cursor.fetchall()
        
        template_names = [t[0] for t in templates]
        self.template_combobox['values'] = template_names
        
        conn.close()
    
    def on_template_select(self, event):
        """模板选择事件"""
        template_name = self.template_combobox.get()
        if not template_name:
            return
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT subject, content FROM templates WHERE name=?", (template_name,))
        result = cursor.fetchone()
        conn.close()
        
        if result:
            self.subject.delete(0, tk.END)
            self.subject.insert(0, result[0])
            
            self.email_content.delete("1.0", tk.END)
            self.email_content.insert(tk.END, result[1])
    
    def use_template(self):
        """使用选中的模板"""
        selected = self.templates_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个模板")
            return
        
        item = self.templates_tree.item(selected[0])
        template_name = item['values'][1]
        
        # 切换到发送邮件标签页
        # 设置下拉框值
        self.template_combobox.set(template_name)
        # 触发选择事件
        self.on_template_select(None)

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSender(root)
    root.mainloop()
