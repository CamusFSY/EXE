import os

import shutil

import tkinter as tk

from tkinter import filedialog, messagebox, scrolledtext


class FileCopyFilterApp:

    def __init__(self, master):

        self.master = master

        master.title("文件批量筛选复制工具（按 Excel 名称）")

        master.geometry("780x540")

        # 源目录 & 目标目录

        self.src_dir = tk.StringVar(value="（未选择源目录）")

        self.dst_dir = tk.StringVar(value="（未选择目标目录）")

        # 模式：keep = 只复制这些名字；exclude = 排除这些名字复制其他

        self.mode = tk.StringVar(value="keep")

        # ============== 源目录选择 ==============

        frame_src = tk.Frame(master)

        frame_src.pack(fill=tk.X, padx=10, pady=(10, 5))

        tk.Label(frame_src, text="源目录（只读，不会被修改）：").pack(side=tk.LEFT)

        tk.Label(frame_src, textvariable=self.src_dir, fg="blue").pack(

            side=tk.LEFT, expand=True, fill=tk.X

        )

        tk.Button(frame_src, text="选择源目录", command=self.choose_src_folder).pack(

            side=tk.RIGHT, padx=5

        )

        # ============== 目标目录选择 ==============

        frame_dst = tk.Frame(master)

        frame_dst.pack(fill=tk.X, padx=10, pady=(0, 10))

        tk.Label(frame_dst, text="目标目录（保存筛选后的文件）：").pack(side=tk.LEFT)

        tk.Label(frame_dst, textvariable=self.dst_dir, fg="green").pack(

            side=tk.LEFT, expand=True, fill=tk.X

        )

        tk.Button(frame_dst, text="选择目标目录", command=self.choose_dst_folder).pack(

            side=tk.RIGHT, padx=5

        )

        # ============== 模式选择 ==============

        frame_mode = tk.Frame(master)

        frame_mode.pack(fill=tk.X, padx=10)

        tk.Label(frame_mode, text="复制模式：").pack(side=tk.LEFT)

        tk.Radiobutton(

            frame_mode,

            text="只复制这些名字的文件（其他全部忽略）",

            variable=self.mode,

            value="keep",

        ).pack(side=tk.LEFT, padx=5)

        tk.Radiobutton(

            frame_mode,

            text="复制除这些名字外的所有文件",

            variable=self.mode,

            value="exclude",

        ).pack(side=tk.LEFT, padx=5)

        # ============== 文件名输入区域 ==============

        frame_mid = tk.Frame(master)

        frame_mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(

            frame_mid,

            text="请从 Excel 中复制“文件名（不含后缀）”这一列，粘贴到下面，每行一个：",

        ).pack(anchor="w")

        self.text_names = scrolledtext.ScrolledText(frame_mid, height=14)

        self.text_names.pack(fill=tk.BOTH, expand=True)

        # ============== 执行按钮 + 状态 ==============

        frame_bottom = tk.Frame(master)

        frame_bottom.pack(fill=tk.X, padx=10, pady=10)

        self.status_label = tk.Label(frame_bottom, text="就绪。源目录不会被修改。")

        self.status_label.pack(side=tk.LEFT)

        tk.Button(frame_bottom, text="开始执行", command=self.run_task).pack(side=tk.RIGHT)

    # 选择源目录

    def choose_src_folder(self):

        folder = filedialog.askdirectory(title="选择源目录（只读）")

        if folder:

            self.src_dir.set(folder)

    # 选择目标目录

    def choose_dst_folder(self):

        folder = filedialog.askdirectory(title="选择目标目录（用于保存结果文件）")

        if folder:

            self.dst_dir.set(folder)

    # 解析用户粘贴的文件名列表

    def parse_name_list(self, raw_text):

        """

        将用户粘贴的文本解析成一个 set（小写），

        默认按“每行一个名字”来处理，忽略空行。

        """

        names = set()

        for line in raw_text.splitlines():

            name = line.strip()

            if not name:

                continue

            names.add(name.lower())

        return names

    # 主逻辑：扫描源目录，根据名单筛选，复制到目标目录

    def run_task(self):

        src_dir = self.src_dir.get()

        dst_dir = self.dst_dir.get()

        if not src_dir or src_dir.startswith("（未选择"):

            messagebox.showwarning("提示", "请先选择源目录。")

            return

        if not dst_dir or dst_dir.startswith("（未选择"):

            messagebox.showwarning("提示", "请先选择目标目录。")

            return

        # 防止误选成同一个目录

        if os.path.abspath(src_dir) == os.path.abspath(dst_dir):

            messagebox.showwarning(

                "提示", "源目录和目标目录不能是同一个路径，请选择不同的目录。"

            )

            return

        raw_names = self.text_names.get("1.0", tk.END)

        target_names = self.parse_name_list(raw_names)

        if not target_names:

            messagebox.showwarning("提示", "请先在文本框中粘贴至少一个文件名。")

            return

        mode = self.mode.get()

        confirm_msg = [

            f"源目录：{src_dir}",

            f"目标目录：{dst_dir}",

            f"模式：{'只复制这些名字的文件' if mode == 'keep' else '复制除这些名字外的所有文件'}",

            f"导入的名字数量：{len(target_names)}",

            "",

            "确认要执行吗？注意：",

            "  - 源目录中的文件不会被修改、移动或删除；",

            "  - 只会向目标目录复制文件，如遇到重名会自动重命名（加“__数字”后缀）。",

        ]

        if not messagebox.askokcancel("确认执行", "\n".join(confirm_msg)):

            return

        # 确保目标目录存在

        os.makedirs(dst_dir, exist_ok=True)

        self.status_label.config(text="正在扫描源目录并复制文件，请稍候...")

        self.master.update_idletasks()

        # 1. 收集源目录下所有文件

        all_files = []  # (full_src_path, base_name_lower, file_name)

        for dirpath, dirnames, filenames in os.walk(src_dir):

            for fname in filenames:

                full_path = os.path.join(dirpath, fname)

                base, ext = os.path.splitext(fname)

                all_files.append((full_path, base.lower(), fname))

        total_files = len(all_files)

        if total_files == 0:

            messagebox.showinfo("结果", "在该源目录下没有找到任何文件。")

            self.status_label.config(text="源目录为空。")

            return

        copied = []

        skipped = []

        matched_name_set = set()  # 用于统计“哪些名单上的名字在源目录中出现过”

        # 2. 按模式决定要复制哪些文件

        for src_path, base_lower, fname in all_files:

            in_list = base_lower in target_names

            if mode == "keep":

                # 只复制名单中的文件

                if in_list:

                    matched_name_set.add(base_lower)

                    copied.append((src_path, fname))

                else:

                    skipped.append(src_path)

            else:  # mode == "exclude"

                # 排除名单中的文件，复制其它

                if in_list:

                    matched_name_set.add(base_lower)

                    skipped.append(src_path)

                else:

                    copied.append((src_path, fname))

        # 3. 执行复制到目标目录（扁平化，不建子目录）

        copy_count = 0

        rename_count = 0

        error_count = 0

        for src_path, fname in copied:

            dest_path = os.path.join(dst_dir, fname)

            # 如已有同名文件，则自动加 "__数字" 后缀，避免覆盖

            if os.path.exists(dest_path):

                base, ext = os.path.splitext(fname)

                idx = 1

                while True:

                    new_name = f"{base}__{idx}{ext}"

                    dest2 = os.path.join(dst_dir, new_name)

                    if not os.path.exists(dest2):

                        dest_path = dest2

                        rename_count += 1

                        break

                    idx += 1

            try:

                shutil.copy2(src_path, dest_path)

                copy_count += 1

            except Exception as e:

                print(f"复制失败：{src_path} -> {dest_path}，错误：{e}")

                error_count += 1

        # 4. 统计名单中“完全未匹配到任何源文件”的名字

        unmatched_names = sorted(target_names - matched_name_set)

        # 5. 汇总结果

        skipped_count = len(skipped)

        result_lines = [

            f"源目录：{src_dir}",

            f"目标目录：{dst_dir}",

            f"模式：{'只复制这些名字的文件' if mode == 'keep' else '复制除这些名字外的所有文件'}",

            f"总共扫描源文件数：{total_files}",

            f"计划复制的文件数：{len(copied)}",

            f"实际成功复制的文件数：{copy_count}",

            f"复制时因重名而自动重命名的文件数：{rename_count}",

            f"未复制（被忽略）的文件数：{skipped_count}",

            f"复制出错的文件数：{error_count}",

        ]

        if unmatched_names:

            result_lines.append("")

            result_lines.append(

                f"下列 {len(unmatched_names)} 个名字在源目录中没有找到对应文件（按“去后缀的文件名”匹配）："

            )

            preview = unmatched_names[:30]

            result_lines.extend("  - " + n for n in preview)

            if len(unmatched_names) > 30:

                result_lines.append(f"...（共 {len(unmatched_names)} 个未匹配名字）")

        messagebox.showinfo("执行结果", "\n".join(result_lines))

        self.status_label.config(text="完成。源目录未被修改。")


if __name__ == "__main__":

    root = tk.Tk()

    app = FileCopyFilterApp(root)

    root.mainloop()
 
