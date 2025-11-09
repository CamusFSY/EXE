
# -*- coding: utf-8 -*-
"""
批量发送 Outlook 邮件（医院多模板、统一附件目录、可输入日期范围）
- 双击 EXE 运行（用 PyInstaller 打包后）
- 每次运行时输入：附件目录、开始日期、结束日期
- 读取同目录下 mail_config.csv 的模板，逐医院发送
作者：你的小助手
"""
import os
import sys
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path

# 可选：如果担心中文路径，推荐将项目放到只有英文路径的目录
APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) if hasattr(sys, "_MEIPASS") else Path(__file__).resolve().parent

CONFIG_NAME = "mail_config.csv"   # 配置文件名（放在与exe同一目录）
LOG_DIR = APP_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

try:
    import win32com.client as win32
except Exception as e:
    win32 = None


def format_cn_date(d: datetime) -> str:
    return f"{d.year}年{d.month}月{d.day}日"


def read_config_csv(path: Path):
    """
    读取 CSV 配置，返回列表：[{Hospital, To, Cc, Bcc, SubjectTemplate, BodyTemplate}, ...]
    必填列：Hospital, To, SubjectTemplate, BodyTemplate
    可选列：Cc, Bcc
    """
    if not path.exists():
        raise FileNotFoundError(f"未找到配置文件：{path}")
    items = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"Hospital", "To", "SubjectTemplate", "BodyTemplate"}
        missing = [c for c in required if c not in reader.fieldnames]
        if missing:
            raise ValueError(f"配置文件缺少必需列：{', '.join(missing)}\n当前列：{reader.fieldnames}")
        for row in reader:
            # 清理空白
            entry = {k: (row.get(k, "") or "").strip() for k in reader.fieldnames}
            if not entry["Hospital"] or not entry["To"]:
                # 跳过不完整行
                continue
            items.append(entry)
    if not items:
        raise ValueError("配置文件没有有效行。请至少填写一行医院配置。")
    return items


def discover_files(folder: Path):
    """遍历文件夹内所有文件（不含子文件夹），返回绝对路径列表"""
    if not folder.exists() or not folder.is_dir():
        raise NotADirectoryError(f"附件目录无效：{folder}")
    files = [p for p in folder.iterdir() if p.is_file()]
    files.sort()
    if not files:
        raise FileNotFoundError(f"目录下没有可发送的文件：{folder}")
    return files


def make_html_body_from_template(template: str, mapping: dict) -> str:
    """
    将模板渲染为 HTML 正文。
    - 如果模板本身包含 <html> 或 <body> 就直接使用 format 后的内容；
    - 否则自动替换换行 -> <br> 并包裹基本 HTML。
    """
    filled = template.format(**mapping)
    lower = filled.lower()
    if ("<html" in lower) or ("<body" in lower):
        return filled
    # 简单包裹
    filled = filled.replace("\n", "<br>")
    return f"<html><body><div>{filled}</div></body></html>"


def send_one_mail(outlook, to_addr, cc, bcc, subject, html_body, attachments, send_mode: str):
    """
    send_mode: 'draft' => 存草稿； 'send' => 直接发送
    """
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = to_addr
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Subject = subject
    mail.HTMLBody = html_body

    for file_path in attachments:
        mail.Attachments.Add(str(file_path))

    if send_mode == "draft":
        mail.Save()
        return "Saved to Drafts"
    else:
        mail.Send()
        return "Sent"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook 批量发送（医院模板）")
        self.geometry("720x460")
        self.resizable(False, False)

        # Vars
        self.attach_dir_var = tk.StringVar(value="")
        self.start_date_var = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        self.end_date_var = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        self.mode_var = tk.StringVar(value="draft")  # 'draft' or 'send'
        self.config_path = APP_DIR / CONFIG_NAME

        # UI
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True)

        row = 0
        ttk.Label(frm, text="配置文件：").grid(row=row, column=0, sticky="e", **pad)
        self.config_entry = ttk.Entry(frm, width=60)
        self.config_entry.insert(0, str(self.config_path))
        self.config_entry.grid(row=row, column=1, sticky="w", **pad)
        ttk.Button(frm, text="更换…", command=self.pick_config).grid(row=row, column=2, **pad)

        row += 1
        ttk.Label(frm, text="附件目录：").grid(row=row, column=0, sticky="e", **pad)
        self.attach_entry = ttk.Entry(frm, textvariable=self.attach_dir_var, width=60)
        self.attach_entry.grid(row=row, column=1, sticky="w", **pad)
        ttk.Button(frm, text="浏览…", command=self.pick_folder).grid(row=row, column=2, **pad)

        row += 1
        ttk.Label(frm, text="报告开始日期 (YYYY-MM-DD)：").grid(row=row, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.start_date_var, width=20).grid(row=row, column=1, sticky="w", **pad)

        row += 1
        ttk.Label(frm, text="报告结束日期 (YYYY-MM-DD)：").grid(row=row, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.end_date_var, width=20).grid(row=row, column=1, sticky="w", **pad)

        row += 1
        ttk.Label(frm, text="发送方式：").grid(row=row, column=0, sticky="e", **pad)
        mode_box = ttk.Frame(frm)
        mode_box.grid(row=row, column=1, sticky="w", **pad)
        ttk.Radiobutton(mode_box, text="保存到草稿（建议先检查）", variable=self.mode_var, value="draft").pack(side="left")
        ttk.Radiobutton(mode_box, text="直接发送", variable=self.mode_var, value="send").pack(side="left")

        row += 1
        ttk.Separator(frm).grid(row=row, column=0, columnspan=3, sticky="ew", **pad)

        row += 1
        self.run_btn = ttk.Button(frm, text="开始批量发送", command=self.run_sending)
        self.run_btn.grid(row=row, column=1, **pad)

        row += 1
        ttk.Label(frm, text="进度：").grid(row=row, column=0, sticky="e", **pad)
        self.progress = ttk.Progressbar(frm, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=row, column=1, sticky="w", **pad)

        row += 1
        self.status_var = tk.StringVar(value="准备就绪。")
        ttk.Label(frm, textvariable=self.status_var).grid(row=row, column=1, sticky="w", **pad)

        row += 1
        tips = (
            "提示：\n"
            "1) 请先在同目录准备 mail_config.csv（见示例），支持每家医院不同标题/正文模板；\n"
            "2) 模板可使用占位符：{start_date_cn}、{end_date_cn}、{date_range_cn}、{start_date}、{end_date}、{hospital}；\n"
            "3) 建议先选择“保存到草稿”，检查无误后再改为“直接发送”。"
        )
        ttk.Label(frm, text=tips, foreground="#555").grid(row=row, column=1, sticky="w", **pad)

    def pick_folder(self):
        folder = filedialog.askdirectory(title="选择附件目录")
        if folder:
            self.attach_dir_var.set(folder)

    def pick_config(self):
        path = filedialog.askopenfilename(title="选择配置文件", filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")])
        if path:
            self.config_path = Path(path)
            self.config_entry.delete(0, tk.END)
            self.config_entry.insert(0, str(self.config_path))

    def run_sending(self):
        # 基础校验
        if win32 is None:
            messagebox.showerror("缺少依赖", "未检测到 pywin32。请先安装 Python 的 pywin32 包后再打包 EXE。")
            return
        attach_dir = Path(self.attach_dir_var.get().strip())
        if not attach_dir.exists():
            messagebox.showerror("输入有误", "附件目录不存在。")
            return
        # 解析日期
        try:
            start_dt = datetime.strptime(self.start_date_var.get().strip(), "%Y-%m-%d")
            end_dt = datetime.strptime(self.end_date_var.get().strip(), "%Y-%m-%d")
        except Exception:
            messagebox.showerror("输入有误", "日期格式需为 YYYY-MM-DD。")
            return
        if end_dt < start_dt:
            messagebox.showerror("输入有误", "结束日期不能早于开始日期。")
            return
        # 读取配置
        try:
            cfg = read_config_csv(self.config_path)
        except Exception as e:
            messagebox.showerror("配置错误", str(e))
            return

        # 附件列表
        try:
            files = discover_files(attach_dir)
        except Exception as e:
            messagebox.showerror("附件错误", str(e))
            return

        # 占位符映射
        mapping_common = {
            "start_date": start_dt.strftime("%Y-%m-%d"),
            "end_date": end_dt.strftime("%Y-%m-%d"),
            "start_date_cn": format_cn_date(start_dt),
            "end_date_cn": format_cn_date(end_dt),
            "date_range_cn": f"{format_cn_date(start_dt)}至{format_cn_date(end_dt)}",
        }

        outlook = win32.Dispatch("Outlook.Application")

        total = len(cfg)
        self.progress["value"] = 0
        self.progress["maximum"] = total

        # 日志
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        log_path = LOG_DIR / f"send_log_{ts}.csv"
        with log_path.open("w", encoding="utf-8-sig", newline="") as lf:
            log_writer = csv.writer(lf)
            log_writer.writerow(["Time", "Hospital", "To", "Subject", "FileCount", "Mode", "Result"])

            sent_ok = 0
            for i, entry in enumerate(cfg, start=1):
                hospital = entry.get("Hospital", "")
                to_addr = entry.get("To", "")
                cc = entry.get("Cc", "")
                bcc = entry.get("Bcc", "")
                subject_t = entry.get("SubjectTemplate", "")
                body_t = entry.get("BodyTemplate", "")

                mapping = dict(mapping_common)
                mapping["hospital"] = hospital

                try:
                    subject = subject_t.format(**mapping)
                    html_body = make_html_body_from_template(body_t, mapping)
                    result = send_one_mail(outlook, to_addr, cc, bcc, subject, html_body, files, self.mode_var.get())
                    sent_ok += 1 if "Sent" in result or "Saved" in result else 0
                except Exception as e:
                    result = f"ERROR: {e}"

                log_writer.writerow([datetime.now().isoformat(timespec="seconds"), hospital, to_addr, subject, len(files), self.mode_var.get(), result])

                self.progress["value"] = i
                self.status_var.set(f"正在处理：{i}/{total} —— {hospital} ……")
                self.update_idletasks()

        messagebox.showinfo("完成", f"已处理 {total} 家医院，成功：{sent_ok}。\n日志：{log_path}")
        self.status_var.set("完成。")

if __name__ == "__main__":
    app = App()
    app.mainloop()
