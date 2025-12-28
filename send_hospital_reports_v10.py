
# -*- coding: utf-8 -*-
"""
Outlook 批量邮件（医院模板版）— v10
新增/改进：
1) 保存“.msg”为【已发送副本】（Sent Items 中的那一封），针对大附件：支持“保存失败自动重试（含复制后再保存）”、可调最大等待秒数。
2) 发送失败可重试：记录失败序号（CSV 第一列“序号/Seq/ID/Index/No/#”等），提供“仅重发失败项”按钮；状态区显示失败清单。
3) 默认正文字体为“微软雅黑”（Microsoft YaHei），并在界面提供常用字体“代号”下拉，一键填充 CSS font-family。
4) 新增：可在文本框输入“接收失败的收件人邮箱”，自动匹配配置文件所在行，并按行合并重发（保留该行 CC/BCC），发送后保存已发送 .msg。
   仍支持自定义 CSS（可编辑输入框）。
其余：保留 v7/v8 的紧凑布局、DPI 适配、浏览文件夹、段落格式标记等。
"""
import csv
import re
import sys
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Set

# ---- Windows DPI 适配 ----
try:
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor v2
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass

APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) if hasattr(sys, "_MEIPASS") else Path(__file__).resolve().parent
CONFIG_NAME = "mail_config.csv"

try:
    import win32com.client as win32
except Exception:
    win32 = None

# Outlook 常量
OL_MSG = 3      # olMSG
OL_MSG_UNI = 9  # olMSGUnicode
OL_FOLDER_SENT = 5  # olFolderSentMail


def cn_date(d: datetime) -> str:
    """中文日期：2025年12月20日"""
    return f"{d.year}年{d.month}月{d.day}日"



def ensure_sample_config(path: Path):
    if path.exists(): return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "Hospital", "To", "Cc", "Bcc", "SubjectTemplate", "BodyTemplate"])
        w.writerow(["1", "示例医院A", "a@x.com; b@y.com", "", "", "【{hospital}】报告（{date_range_cn}）",
                    "[indent]尊敬的{hospital}同事：\n\n[indent]附件为 {date_range_cn} 的相关报告，请查收。\n[right]此致\n[right]敬礼"])
        w.writerow(["2", "示例医院B", "exampleB@hospital.com", "", "", "{hospital}—每周报告（{start_date} 至 {end_date}）",
                    "Dear team at {hospital},\n\nPlease find attached the reports covering {start_date_cn} 至 {end_date_cn}.\n[right]Best regards."])


SEQ_CANDIDATES = ["序号", "Seq", "ID", "Index", "No", "#", "Number"]


def read_config(path: Path) -> List[Dict[str, str]]:
    if not path.exists():
        raise FileNotFoundError(f"未找到配置文件：{path}")
    items = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("配置文件没有表头。")
        # 找序号列
        seq_col = None
        for cand in SEQ_CANDIDATES:
            if cand in reader.fieldnames:
                seq_col = cand
                break
        # 必要列
        required = {"Hospital", "To", "SubjectTemplate", "BodyTemplate"}
        missing = [c for c in required if c not in reader.fieldnames]
        if missing:
            raise ValueError(f"配置文件缺少列：{', '.join(missing)}")

        for idx, row in enumerate(reader, start=1):
            row = {k: (row.get(k, "") or "").strip() for k in reader.fieldnames}
            if not row.get("Hospital") or not row.get("To"):
                continue
            seq = row.get(seq_col, str(idx)) if seq_col else str(idx)
            row["Seq"] = seq
            items.append(row)

    if not items:
        raise ValueError("配置文件里没有有效行。")
    return items


def find_files(folder: Path):
    return sorted([p for p in folder.iterdir() if p.is_file()])


def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", (s or "")).strip()[:150] or "email"



EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

def extract_emails(s: str) -> List[str]:
    """从字符串中提取邮箱（统一为小写），支持 'Name <a@b.com>'、分号/逗号/换行等。"""
    if not s:
        return []
    s = str(s)
    # 常见中文分隔符归一化
    s = s.replace("；", ";").replace("，", ",").replace("、", ";")
    found = EMAIL_RE.findall(s)
    # 去重但保持顺序
    out = []
    seen = set()
    for e in found:
        e2 = e.strip().lower()
        if e2 and e2 not in seen:
            seen.add(e2)
            out.append(e2)
    return out

def render_html(template: str, mapping: dict, font_family: str, font_pt: int) -> str:
    filled = template.format(**mapping)
    low = filled.lower()
    if "<html" in low and "<body" in low:
        return filled

    body_style = f"font-family:{font_family}; font-size:{font_pt}pt; line-height:1.6;"
    out = []
    for raw in filled.splitlines():
        line = raw.rstrip("\r")
        if not line.strip():
            out.append("<p>&nbsp;</p>")
            continue
        align = "left"
        indent = "0"
        m = re.match(r'^\s*\[([a-zA-Z]+)\]\s*(.*)$', line)
        content = line
        if m:
            tag, content = m.group(1).lower(), m.group(2)
            if tag == "indent":
                indent = "2em"
            elif tag == "right":
                align = "right"
            elif tag == "center":
                align = "center"
            elif tag == "left":
                align = "left"
        content = (content
                   .replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace("  ", "&nbsp;&nbsp;"))
        out.append(f'<p style="text-align:{align}; text-indent:{indent};">{content}</p>')
    return f'<html><body style="{body_style}">' + "".join(out) + "</body></html>"


def outlook_restrict_datetime(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")


def find_sent_item(outlook, subject: str, sent_after: datetime, expect_attach_count: Optional[int] = None):
    ns = outlook.GetNamespace("MAPI")
    sent_folder = ns.GetDefaultFolder(OL_FOLDER_SENT)
    items = sent_folder.Items
    subj = subject.replace("'", "''")
    date_str = outlook_restrict_datetime(sent_after)
    restriction = f"[Subject] = '{subj}' AND [SentOn] >= '{date_str}'"
    try:
        subset = items.Restrict(restriction)
        subset.Sort("[SentOn]", True)
    except Exception:
        subset = items
        subset.Sort("[SentOn]", True)

    count = min(60, getattr(subset, "Count", 0) or 0)
    for i in range(1, count + 1):
        try:
            it = subset.Item(i)
        except Exception:
            continue
        try:
            if getattr(it, "Class", None) != 43:  # olMail
                continue
            if getattr(it, "Subject", "") != subject:
                continue
            if expect_attach_count is not None:
                try:
                    if it.Attachments.Count != expect_attach_count:
                        continue
                except Exception:
                    pass
            return it
        except Exception:
            continue
    return None


def unique_path_by_subject(msg_dir: Path, subject: str) -> Path:
    base = safe_name(subject)
    p = msg_dir / f"{base}.msg"
    if not p.exists():
        return p
    i = 1
    while True:
        p2 = msg_dir / f"{base} ({i}).msg"
        if not p2.exists():
            return p2
        i += 1


def save_sent_copy_as_msg_with_retries(sent_item, msg_dir: Path, subject: str,
                                       retries: int = 3, wait_each: float = 3.0) -> Optional[Path]:
    msg_dir.mkdir(parents=True, exist_ok=True)
    for attempt in range(1, max(1, retries) + 1):
        target = unique_path_by_subject(msg_dir, subject)
        try:
            obj = sent_item if attempt == 1 else sent_item.Copy()
            try:
                obj.SaveAs(str(target), OL_MSG)
            except Exception:
                obj.SaveAs(str(target), OL_MSG_UNI)
            if attempt > 1:
                try:
                    obj.Delete()
                except Exception:
                    pass
            return target
        except Exception:
            time.sleep(max(0.2, wait_each))
    return None


def send_mail_and_archive(outlook, to_addr, cc, bcc, subject, html_body, attachments, save_msg, msg_dir: Path,
                          poll_seconds: float = 120.0, poll_interval: float = 1.2,
                          save_retries: int = 4, save_wait_each: float = 4.0) -> Tuple[str, Optional[Path]]:
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = to_addr
    if cc: mail.CC = cc
    if bcc: mail.BCC = bcc
    mail.Subject = subject
    mail.HTMLBody = html_body
    for a in attachments:
        mail.Attachments.Add(str(a))

    send_mark = datetime.now() - timedelta(minutes=2)
    expected_attach = len(list(attachments))
    mail.Send()

    saved = None
    if save_msg and msg_dir:
        deadline = time.time() + float(max(10.0, poll_seconds))
        while time.time() < deadline:
            try:
                sent_item = find_sent_item(outlook, subject, send_mark, expect_attach_count=expected_attach)
                if sent_item is not None:
                    saved = save_sent_copy_as_msg_with_retries(
                        sent_item, msg_dir, subject,
                        retries=max(1, save_retries),
                        wait_each=max(0.5, save_wait_each),
                    )
                    break
            except Exception:
                pass
            time.sleep(max(0.2, poll_interval))

    return "Sent", saved


# ---------- 可滚动容器 ----------
class ScrollableFrame(ttk.Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook 批量发送")
        self.geometry("1000x780")
        self.minsize(900, 720)
        self.resizable(True, True)

        # 样式/缩放/密度
        self.base_font_family = "Microsoft YaHei UI"
        self.base_font_size = 10
        self.scale_var = tk.DoubleVar(value=1.05)
        self.density_var = tk.StringVar(value="紧凑")

        self.style = ttk.Style(self)
        try: self.style.theme_use("clam")
        except Exception: pass

        self.font_base = tkfont.Font(family=self.base_font_family, size=self.base_font_size)
        self.font_big = tkfont.Font(family=self.base_font_family, size=self.base_font_size + 1, weight="bold")
        self.font_entry = tkfont.Font(family=self.base_font_family, size=self.base_font_size + 2)
        self.font_list = tkfont.Font(family=self.base_font_family, size=self.base_font_size + 2)

        self.apply_style()

        toolbar = ttk.Frame(self, padding=(10, 6))
        toolbar.pack(side="top", fill="x")
        ttk.Label(toolbar, text="布局密度：").pack(side="left")
        self.density_box = ttk.Combobox(toolbar, values=["紧凑", "标准", "宽松"], width=6, state="readonly",
                                        textvariable=self.density_var)
        self.density_box.pack(side="left", padx=(4, 10))
        self.density_box.bind("<<ComboboxSelected>>", self.on_density_change)

        ttk.Label(toolbar, text="界面缩放：").pack(side="left")
        self.scale_ctl = ttk.Scale(toolbar, from_=0.9, to=1.4, variable=self.scale_var,
                                   command=self.on_scale_change, length=180)
        self.scale_ctl.pack(side="left", padx=8)

        container = ScrollableFrame(self)
        container.pack(fill="both", expand=True, padx=8, pady=8)
        self.root = container.inner

        # 业务变量
        self.attach_mode = tk.StringVar(value="folder")
        self.attach_dir = tk.StringVar(value="")
        self.selected_files = []

        self.config_path = tk.StringVar(value=str((APP_DIR / CONFIG_NAME)))
        self.start = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        self.end = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        self.mode = tk.StringVar(value="draft")  # draft / send

        # 字体：默认微软雅黑（Microsoft YaHei）
        self.font_presets = {
            "微软雅黑 (推荐)": "'Microsoft YaHei', '微软雅黑', 'Microsoft YaHei UI', SimSun, Arial",
            "宋体 (SimSun)": "SimSun, '宋体', 'Microsoft YaHei', Arial",
            "黑体 (SimHei)": "SimHei, '黑体', 'Microsoft YaHei', Arial",
            "仿宋 (FangSong)": "FangSong, '仿宋', 'Microsoft YaHei', SimSun",
            "Arial": "Arial, Helvetica, sans-serif",
            "Times New Roman": "'Times New Roman', Times, serif",
        }
        self.font_family = tk.StringVar(value=self.font_presets["微软雅黑 (推荐)"])
        self.font_pt = tk.StringVar(value="10")

        # 已发送 .msg 存档控制（大附件容错）
        self.save_msg = tk.BooleanVar(value=True)
        self.msg_dir = tk.StringVar(value=str(APP_DIR / "sent_msgs"))
        self.retry_save = tk.BooleanVar(value=True)
        self.save_retries = tk.IntVar(value=4)
        self.poll_seconds = tk.IntVar(value=120)

        # 失败重试
        self.failed_seqs: Set[str] = set()

        self.build_ui()
        self.apply_density_layout()
        self.apply_runtime_widget_fonts()

    # ---------- 样式 ----------
    def apply_style(self):
        s = float(self.scale_var.get() or 1.0)
        try: self.tk.call('tk', 'scaling', s)
        except Exception: pass

        base_sz = max(9, int(self.base_font_size * s))
        self.font_base.configure(size=base_sz)
        self.font_big.configure(size=base_sz + 1)
        self.font_entry.configure(size=base_sz + 2)
        self.font_list.configure(size=base_sz + 2)

        density = getattr(self, "density_var", None)
        density = density.get() if density else "紧凑"
        if density == "紧凑":
            pad_btn = (8, 5); pad_entry = (5, 4); pad_lbl = (3, 2); pad_group = (8, 6)
        elif density == "标准":
            pad_btn = (10, 7); pad_entry = (6, 5); pad_lbl = (4, 3); pad_group = (10, 8)
        else:
            pad_btn = (12, 9); pad_entry = (8, 6); pad_lbl = (6, 4); pad_group = (12, 10)

        self.style.configure(".", font=self.font_base)
        self.style.configure("TButton", padding=pad_btn)
        self.style.configure("TEntry", padding=pad_entry, font=self.font_entry)
        self.style.configure("TLabel", padding=pad_lbl)
        self.style.configure("TLabelframe", padding=pad_group)
        self.style.configure("TLabelframe.Label", font=self.font_big)
        self.style.configure("TCheckbutton", padding=pad_entry)
        self.style.configure("TRadiobutton", padding=pad_entry)

    def apply_runtime_widget_fonts(self):
        if hasattr(self, "file_list"):
            try: self.file_list.configure(font=self.font_list)
            except Exception: pass
        if hasattr(self, "failed_list"):
            try: self.failed_list.configure(font=self.font_list)
            except Exception: pass
        if hasattr(self, "resend_text"):
            try: self.resend_text.configure(font=self.font_list)
            except Exception: pass


    def on_scale_change(self, _):
        self.apply_style(); self.apply_runtime_widget_fonts()

    def on_density_change(self, _):
        self.apply_style(); self.apply_density_layout(); self.apply_runtime_widget_fonts()

    # ---------- UI ----------
    def build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # ① 配置与日期
        self.sec1 = ttk.LabelFrame(self.root, text="① 配置与日期")
        self.sec1.grid(row=0, column=0, sticky="w", **pad)
        self.sec1.grid_columnconfigure(1, weight=0)

        ttk.Label(self.sec1, text="配置文件：").grid(row=0, column=0, sticky="e", **pad)
        self.ent_config_path = ttk.Entry(self.sec1, textvariable=self.config_path, width=58)
        self.ent_config_path.grid(row=0, column=1, sticky="w", **pad)
        cfg_btns = ttk.Frame(self.sec1); cfg_btns.grid(row=0, column=2, sticky="w", **pad)
        ttk.Button(cfg_btns, text="选择文件…", command=self.pick_config_file, width=12).pack(fill="x", pady=2)
        ttk.Button(cfg_btns, text="浏览文件夹…", command=self.pick_config_folder, width=12).pack(fill="x", pady=2)

        ttk.Label(self.sec1, text="开始日期 (YYYY-MM-DD)：").grid(row=1, column=0, sticky="e", **pad)
        self.ent_start = ttk.Entry(self.sec1, textvariable=self.start, width=20)
        self.ent_start.grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(self.sec1, text="结束日期 (YYYY-MM-DD)：").grid(row=2, column=0, sticky="e", **pad)
        self.ent_end = ttk.Entry(self.sec1, textvariable=self.end, width=20)
        self.ent_end.grid(row=2, column=1, sticky="w", **pad)

        # ② 附件选择
        self.sec2 = ttk.LabelFrame(self.root, text="② 附件选择")
        self.sec2.grid(row=1, column=0, sticky="w", **pad)
        self.sec2.grid_columnconfigure(1, weight=0)

        ttk.Label(self.sec2, text="附件模式：").grid(row=0, column=0, sticky="e", **pad)
        am = ttk.Frame(self.sec2); am.grid(row=0, column=1, sticky="w", **pad)
        ttk.Radiobutton(am, text="目录全部文件", variable=self.attach_mode, value="folder", command=self.toggle_attach).pack(side="left", padx=(0,10))
        ttk.Radiobutton(am, text="手动选择文件", variable=self.attach_mode, value="files", command=self.toggle_attach).pack(side="left")

        self.folder_frame = ttk.Frame(self.sec2)
        self.folder_frame.grid(row=1, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(self.folder_frame, text="附件目录：").grid(row=0, column=0, sticky="e", **pad)
        self.ent_attach_dir = ttk.Entry(self.folder_frame, textvariable=self.attach_dir, width=58)
        self.ent_attach_dir.grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(self.folder_frame, text="浏览文件夹…", command=self.pick_attach_folder, width=12).grid(row=0, column=2, **pad)

        self.files_frame = ttk.Frame(self.sec2)
        self.files_frame.grid(row=2, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(self.files_frame, text="已选文件：").grid(row=0, column=0, sticky="ne", **pad)
        right = ttk.Frame(self.files_frame); right.grid(row=0, column=1, sticky="w", **pad)
        self.file_list = tk.Listbox(right, height=7, width=70)
        self.file_list.pack(side="left")
        sb = ttk.Scrollbar(right, orient="vertical", command=self.file_list.yview)
        sb.pack(side="left", fill="y")
        self.file_list.configure(yscrollcommand=sb.set)
        btns = ttk.Frame(self.files_frame); btns.grid(row=0, column=2, sticky="n", **pad)
        ttk.Button(btns, text="添加文件…", command=self.add_files, width=12).pack(fill="x", pady=2)
        ttk.Button(btns, text="清空列表", command=self.clear_files, width=12).pack(fill="x", pady=2)

        # ③ 正文样式 / 存档
        self.sec3 = ttk.LabelFrame(self.root, text="③ 正文样式与存档设置")
        self.sec3.grid(row=2, column=0, sticky="w", **pad)
        self.sec3.grid_columnconfigure(1, weight=0)

        ttk.Label(self.sec3, text="字体预设：").grid(row=0, column=0, sticky="e", **pad)
        self.font_combo = ttk.Combobox(self.sec3, values=list(self.font_presets.keys()), width=22, state="readonly")
        self.font_combo.grid(row=0, column=1, sticky="w", **pad)
        self.font_combo.set("微软雅黑 (推荐)")
        self.font_combo.bind("<<ComboboxSelected>>", self.on_font_preset_changed)

        ttk.Label(self.sec3, text="font-family（可编辑）：").grid(row=1, column=0, sticky="e", **pad)
        self.ent_font_family = ttk.Entry(self.sec3, textvariable=self.font_family, width=52)
        self.ent_font_family.grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(self.sec3, text="字号（pt）：").grid(row=1, column=2, sticky="e", padx=(6, 0))
        self.ent_font_pt = ttk.Entry(self.sec3, textvariable=self.font_pt, width=10)
        self.ent_font_pt.grid(row=1, column=3, sticky="w", padx=(0, 10), pady=5)

        ttk.Label(self.sec3, text="保存 .msg：").grid(row=2, column=0, sticky="e", **pad)
        save_box = ttk.Frame(self.sec3); save_box.grid(row=2, column=1, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(save_box, text="（仅‘直接发送’时）保存已发送 .msg 到下方目录", variable=self.save_msg).pack(side="left", padx=(0,10))
        ttk.Checkbutton(save_box, text="保存失败自动重试（适合大附件）", variable=self.retry_save).pack(side="left")

        ttk.Label(self.sec3, text="保存目录：").grid(row=3, column=0, sticky="e", **pad)
        self.ent_msg_dir = ttk.Entry(self.sec3, textvariable=self.msg_dir, width=52)
        self.ent_msg_dir.grid(row=3, column=1, sticky="w", **pad)
        ttk.Button(self.sec3, text="浏览文件夹…", command=self.pick_msg_folder, width=12).grid(row=3, column=2, sticky="w", **pad)

        ttk.Label(self.sec3, text="最大等待秒：").grid(row=4, column=0, sticky="e", **pad)
        self.ent_poll_seconds = ttk.Entry(self.sec3, textvariable=self.poll_seconds, width=10)
        self.ent_poll_seconds.grid(row=4, column=1, sticky="w", **pad)
        ttk.Label(self.sec3, text="保存重试次数：").grid(row=4, column=2, sticky="e", padx=(6, 0))
        self.ent_save_retries = ttk.Entry(self.sec3, textvariable=self.save_retries, width=10)
        self.ent_save_retries.grid(row=4, column=3, sticky="w", padx=(0, 10), pady=5)

        # ④ 执行 / 重试
        self.sec4 = ttk.LabelFrame(self.root, text="④ 执行与重试")
        self.sec4.grid(row=3, column=0, sticky="w", **pad)
        mode_box = ttk.Frame(self.sec4); mode_box.grid(row=0, column=0, sticky="w", **pad)
        ttk.Radiobutton(mode_box, text="保存到草稿（不保存 .msg）", variable=self.mode, value="draft").pack(side="left", padx=(0,10))
        ttk.Radiobutton(mode_box, text="直接发送（并存档 .msg）", variable=self.mode, value="send").pack(side="left")
        ttk.Button(self.sec4, text="开始批量发送", command=self.run_all, width=16).grid(row=0, column=1, sticky="e", **pad)
        ttk.Button(self.sec4, text="仅重发失败项", command=self.run_failed, width=16).grid(row=0, column=2, sticky="e", **pad)

        # 失败清单
        fail_box = ttk.Frame(self.sec4); fail_box.grid(row=1, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(fail_box, text="失败清单（序号 - 医院 - 错误）：").pack(anchor="w")
        self.failed_list = tk.Listbox(fail_box, height=6, width=110)
        self.failed_list.pack(fill="x", expand=True)

        # 指定收件人重发（按配置行自动归类）
        resend_box = ttk.Frame(self.sec4)
        resend_box.grid(row=2, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(
            resend_box,
            text="指定收件人重发（输入接收失败的邮箱，自动匹配配置行；支持 ; , 空格/换行分隔）："
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        self.resend_text = tk.Text(resend_box, height=3, width=90)
        self.resend_text.grid(row=1, column=0, sticky="w", pady=(4, 0))
        resend_btns = ttk.Frame(resend_box)
        resend_btns.grid(row=1, column=1, sticky="n", padx=(8, 0), pady=(4, 0))
        ttk.Button(resend_btns, text="按收件人重发", command=self.run_resend_by_recipients, width=16).pack(fill="x")
        ttk.Button(resend_btns, text="清空", command=lambda: self.resend_text.delete("1.0", tk.END), width=16).pack(fill="x", pady=(6, 0))

        # 状态栏
        self.status = tk.StringVar(value="准备就绪。")
        status_bar = ttk.Frame(self.root); status_bar.grid(row=4, column=0, sticky="w", padx=10, pady=(0,6))
        ttk.Label(status_bar, textvariable=self.status).pack(side="left")

        self.toggle_attach()

    # ---------- 字体预设 ----------
    def on_font_preset_changed(self, _):
        name = self.font_combo.get()
        css = self.font_presets.get(name, self.font_presets["微软雅黑 (推荐)"])
        self.font_family.set(css)

    # ---------- 密度 ----------
    def apply_density_layout(self):
        density = self.density_var.get()
        if density == "紧凑":
            for sec in (self.sec1, self.sec2, self.sec3):
                sec.grid_columnconfigure(1, weight=0)
            self.ent_config_path.configure(width=58)
            self.ent_attach_dir.configure(width=58)
            self.ent_msg_dir.configure(width=52)
            self.file_list.configure(height=7, width=70)
            self.geometry("980x780")
        elif density == "标准":
            for sec in (self.sec1, self.sec2, self.sec3):
                sec.grid_columnconfigure(1, weight=1)
            self.ent_config_path.configure(width=68)
            self.ent_attach_dir.configure(width=68)
            self.ent_msg_dir.configure(width=62)
            self.file_list.configure(height=8, width=78)
            self.geometry("1100x860")
        else:
            for sec in (self.sec1, self.sec2, self.sec3):
                sec.grid_columnconfigure(1, weight=1)
            self.ent_config_path.configure(width=80)
            self.ent_attach_dir.configure(width=80)
            self.ent_msg_dir.configure(width=74)
            self.file_list.configure(height=9, width=86)
            self.geometry("1260x920")

    # ---------- 路径选择 ----------
    def pick_config_file(self):
        p = filedialog.askopenfilename(title="选择配置文件", filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if p: self.config_path.set(p)

    def pick_config_folder(self):
        d = filedialog.askdirectory(title="选择配置文件所在文件夹")
        if not d: return
        cfg = Path(d) / CONFIG_NAME
        ensure_sample_config(cfg)
        self.config_path.set(str(cfg))
        messagebox.showinfo("配置文件", f"已指向：{cfg}")

    def pick_attach_folder(self):
        d = filedialog.askdirectory(title="选择附件目录")
        if d: self.attach_dir.set(d)

    def pick_msg_folder(self):
        d = filedialog.askdirectory(title="选择 .msg 保存目录")
        if d: self.msg_dir.set(d)

    # ---------- 附件选择 ----------
    def add_files(self):
        paths = filedialog.askopenfilenames(title="选择附件文件",
                                            filetypes=[("所有文件", "*.*"), ("PDF", "*.pdf"),
                                                       ("Word", "*.doc *.docx"), ("Excel", "*.xls *.xlsx")])
        if not paths: return
        for p in paths:
            p = Path(p)
            if p.is_file():
                self.selected_files.append(p)
                self.file_list.insert(tk.END, str(p))

    def clear_files(self):
        self.selected_files.clear()
        self.file_list.delete(0, tk.END)

    def toggle_attach(self):
        if self.attach_mode.get() == "folder":
            self.folder_frame.grid()
            self.files_frame.grid_remove()
        else:
            self.files_frame.grid()
            self.folder_frame.grid_remove()

    # ---------- 执行（全部/失败） ----------
    def run_all(self):
        self.failed_seqs.clear()
        self.failed_list.delete(0, tk.END)
        self._run(batch="all")

    def run_failed(self):
        if not self.failed_seqs:
            messagebox.showinfo("提示", "当前没有失败的序号需要重发。")
            return
        self._run(batch="failed")



    def run_resend_by_recipients(self):
        """输入若干‘接收失败’邮箱，自动匹配配置文件行，并按行合并重发（保留该行 CC/BCC），且保存已发送 .msg。"""
        try:
            raw = self.resend_text.get("1.0", tk.END).strip()
        except Exception:
            raw = ""
        emails = extract_emails(raw)
        if not emails:
            messagebox.showinfo("提示", "请在“指定收件人重发”框中输入至少一个邮箱。")
            return

        # 仅在发送模式下支持“保存已发送 .msg”
        if self.mode.get() != "send":
            messagebox.showinfo("提示", "该功能建议在“直接发送（并存档 .msg）”模式下使用。请切换到‘直接发送’后再执行。")
            return
        if not bool(self.save_msg.get()) or not (self.msg_dir.get().strip()):
            messagebox.showinfo("提示", "该功能需要保存 .msg：请勾选“保存 .msg”并设置保存目录后再执行。")
            return

        self._run_resend_by_emails(emails)

    def _run_resend_by_emails(self, emails: List[str]):
        # 扫描公用输入（日期/配置/附件/字体/存档参数）
        try:
            s, e, items, attachments, font_family, font_pt, save_msg, msg_dir, poll_seconds, save_retries = self._collect_common()
        except Exception as ex:
            messagebox.showerror("输入有误", str(ex))
            return

        if win32 is None:
            messagebox.showerror("缺少依赖", "未检测到 pywin32。请先安装 pywin32。")
            return

        # 构造公共占位符
        mapping_common = {
            "start_date": s.strftime("%Y-%m-%d"),
            "end_date": e.strftime("%Y-%m-%d"),
            "start_date_cn": cn_date(s),
            "end_date_cn": cn_date(e),
            "date_range_cn": f"{cn_date(s)}至{cn_date(e)}",
        }

        # email -> Seq(s) 映射（只按 To 匹配）
        email_to_seqs: Dict[str, List[str]] = {}
        seq_to_item: Dict[str, Dict[str, str]] = {}
        for it in items:
            seq = str(it.get("Seq", "")).strip()
            if not seq:
                continue
            seq_to_item[seq] = it
            for em in extract_emails(it.get("To", "")):
                email_to_seqs.setdefault(em, []).append(seq)

        # 输入邮箱按“配置行”归类；同一行合并一起发
        groups: Dict[str, Set[str]] = {}
        group_order: List[str] = []
        unmatched: List[str] = []
        ambiguous: List[str] = []
        for em in emails:
            seqs = email_to_seqs.get(em, [])
            if not seqs:
                unmatched.append(em)
                continue
            if len(seqs) > 1:
                ambiguous.append(f"{em} -> {', '.join(seqs)}（已默认取第一行）")
            seq = seqs[0]
            if seq not in groups:
                groups[seq] = set()
                group_order.append(seq)
            groups[seq].add(em)

        if not groups:
            messagebox.showinfo("完成", "未找到可重发的配置行（输入邮箱均未匹配到配置文件的 To 列）。")
            return

        outlook = win32.Dispatch("Outlook.Application")

        # 结果统计
        total = len(group_order)
        ok = 0
        archived = 0
        failed_count = 0

        self.failed_list.insert(tk.END, "---- 指定收件人重发 ----")

        for i, seq in enumerate(group_order, start=1):
            it = seq_to_item.get(seq)
            if not it:
                continue
            hospital = it.get("Hospital", "")
            cc = it.get("Cc", "")
            bcc = it.get("Bcc", "")
            subject_t = it.get("SubjectTemplate", "")
            body_t = it.get("BodyTemplate", "")

            to_list = sorted(groups.get(seq, set()))
            to_addr = "; ".join(to_list)

            m = dict(mapping_common)
            m["hospital"] = hospital

            try:
                subject = subject_t.format(**m)
                html = render_html(body_t, m, font_family, font_pt)

                res, saved = send_mail_and_archive(
                    outlook, to_addr, cc, bcc, subject, html, attachments,
                    save_msg=True, msg_dir=msg_dir,
                    poll_seconds=poll_seconds,
                    poll_interval=1.2,
                    save_retries=(save_retries if self.retry_save.get() else 1),
                    save_wait_each=3.0
                )
                if res == "Sent":
                    ok += 1
                    if saved:
                        archived += 1
                else:
                    raise RuntimeError(f"未知状态：{res}")
            except Exception as ex:
                failed_count += 1
                self.failed_list.insert(tk.END, f"{seq} - {hospital} - To: {to_addr} - {ex}")

            self.status.set(f"重发进度：{i}/{total}")

        # 汇总提示
        lines = [f"按收件人重发完成：共 {total} 组（按配置行合并），成功：{ok}；成功存档 .msg：{archived}；失败：{failed_count}。"]
        if unmatched:
            lines.append("")
            lines.append("未匹配到配置行（配置文件 To 列未包含这些邮箱）：")
            lines.append("；".join(unmatched))
        if ambiguous:
            lines.append("")
            lines.append("配置文件中重复出现的邮箱（已默认取第一行）：")
            lines.extend(ambiguous)

        messagebox.showinfo("完成", "\n".join(lines))
    def _collect_common(self):
        # 日期
        try:
            s = datetime.strptime(self.start.get().strip(), "%Y-%m-%d")
            e = datetime.strptime(self.end.get().strip(), "%Y-%m-%d")
        except Exception:
            raise ValueError("日期格式需为 YYYY-MM-DD。")
        if e < s:
            raise ValueError("结束日期不能早于开始日期。")

        # 配置
        cfg = Path(self.config_path.get().strip())
        items = read_config(cfg)

        # 附件
        if self.attach_mode.get() == "folder":
            d = Path(self.attach_dir.get().strip())
            if not d.exists() or not d.is_dir():
                raise ValueError("附件目录不存在或无效。")
            attachments = find_files(d)
            if not attachments:
                raise ValueError("目录下没有可发送的文件。")
        else:
            attachments = [p for p in self.selected_files if Path(p).exists() and Path(p).is_file()]
            if not attachments:
                raise ValueError("未选择任何有效文件。")

        # 存档参数
        font_family = self.font_family.get().strip() or self.font_presets["微软雅黑 (推荐)"]
        try:
            font_pt = int(self.font_pt.get().strip())
        except Exception:
            font_pt = 12

        save_msg = bool(self.save_msg.get())
        poll_seconds = int(self.poll_seconds.get() or 120)
        save_retries = int(self.save_retries.get() or 4)
        msg_dir = Path(self.msg_dir.get().strip()) if (save_msg and self.mode.get() == "send") else None
        if self.mode.get() == "send" and save_msg and not self.msg_dir.get().strip():
            raise ValueError("选择了‘直接发送并存档’，但未指定 .msg 保存目录。")

        return s, e, items, attachments, font_family, font_pt, save_msg, msg_dir, poll_seconds, save_retries

    def _run(self, batch: str):
        # 扫描公用输入
        try:
            s, e, items, attachments, font_family, font_pt, save_msg, msg_dir, poll_seconds, save_retries = self._collect_common()
        except Exception as ex:
            messagebox.showerror("输入有误", str(ex))
            return

        # 过滤失败项
        if batch == "failed":
            items = [it for it in items if it.get("Seq") in self.failed_seqs]
            if not items:
                messagebox.showinfo("提示", "失败队列已清空或配置文件已变更。")
                return

        # 构造公共占位符
        mapping_common = {
            "start_date": s.strftime("%Y-%m-%d"),
            "end_date": e.strftime("%Y-%m-%d"),
            "start_date_cn": cn_date(s),
            "end_date_cn": cn_date(e),
            "date_range_cn": f"{cn_date(s)}至{cn_date(e)}",
        }

        if win32 is None:
            messagebox.showerror("缺少依赖", "未检测到 pywin32。请先安装 pywin32。")
            return
        outlook = win32.Dispatch("Outlook.Application")

        total = len(items); ok = 0; archived = 0; failed_count = 0
        for i, it in enumerate(items, start=1):
            seq = str(it.get("Seq", i))
            hospital = it.get("Hospital", "")
            to_addr = it.get("To", "")
            cc = it.get("Cc", "")
            bcc = it.get("Bcc", "")
            subject_t = it.get("SubjectTemplate", "")
            body_t = it.get("BodyTemplate", "")

            m = dict(mapping_common); m["hospital"] = hospital
            try:
                subject = subject_t.format(**m)
                html = render_html(body_t, m, font_family, font_pt)

                if self.mode.get() == "draft":
                    mail = outlook.CreateItem(0)
                    mail.To = to_addr
                    if cc: mail.CC = cc
                    if bcc: mail.BCC = bcc
                    mail.Subject = subject
                    mail.HTMLBody = html
                    for a in attachments: mail.Attachments.Add(str(a))
                    mail.Save()
                    saved = None
                    res = "Saved to Drafts"
                else:
                    res, saved = send_mail_and_archive(
                        outlook, to_addr, cc, bcc, subject, html, attachments,
                        save_msg=save_msg, msg_dir=msg_dir,
                        poll_seconds=poll_seconds,
                        poll_interval=1.2,
                        save_retries=(save_retries if self.retry_save.get() else 1),
                        save_wait_each=3.0
                    )

                if res in ("Sent", "Saved to Drafts"):
                    ok += 1
                    if saved: archived += 1
                    if seq in self.failed_seqs:
                        self.failed_seqs.remove(seq)
                else:
                    raise RuntimeError(f"未知状态：{res}")
            except Exception as ex:
                failed_count += 1
                self.failed_seqs.add(seq)
                self.failed_list.insert(tk.END, f"{seq} - {hospital} - {ex}")

            self.status.set(f"进度：{i}/{total}")

        if self.mode.get() == "send" and save_msg:
            messagebox.showinfo("完成", f"共处理 {total} 封，成功：{ok}；成功存档 .msg：{archived}；失败：{failed_count}。")
        else:
            messagebox.showinfo("完成", f"共处理 {total} 封，成功：{ok}；失败：{failed_count}。")

    # ---------- 事件 ----------
    def on_scale_change(self, _):
        self.apply_style(); self.apply_runtime_widget_fonts()

    def on_density_change(self, _):
        self.apply_style(); self.apply_density_layout(); self.apply_runtime_widget_fonts()


if __name__ == "__main__":
    App().mainloop()
