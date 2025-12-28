import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

class FileCleanerApp:
   def __init__(self, master):
       self.master = master
       master.title("文件批量保留/删除工具（按 Excel 名称）")
       master.geometry("700x500")
       # 当前选择的根目录
       self.root_dir = tk.StringVar(value="（未选择）")
       # 模式：keep = 保留这些名字；delete = 删除这些名字
       self.mode = tk.StringVar(value="keep")
       # ============== 根目录选择区域 ==============
       frame_top = tk.Frame(master)
       frame_top.pack(fill=tk.X, padx=10, pady=10)
       tk.Label(frame_top, text="根目录：").pack(side=tk.LEFT)
       tk.Label(frame_top, textvariable=self.root_dir, fg="blue").pack(
           side=tk.LEFT, expand=True, fill=tk.X
       )
       tk.Button(frame_top, text="选择文件夹", command=self.choose_folder).pack(side=tk.RIGHT)
       # ============== 模式选择 ==============
       frame_mode = tk.Frame(master)
       frame_mode.pack(fill=tk.X, padx=10)
       tk.Label(frame_mode, text="操作模式：").pack(side=tk.LEFT)
       tk.Radiobutton(
           frame_mode,
           text="保留这些名字的文件（删除其他）",
           variable=self.mode,
           value="keep",
       ).pack(side=tk.LEFT, padx=5)
       tk.Radiobutton(
           frame_mode,
           text="删除这些名字的文件（保留其他）",
           variable=self.mode,
           value="delete",
       ).pack(side=tk.LEFT, padx=5)
       # ============== 文件名输入区域 ==============
       frame_mid = tk.Frame(master)
       frame_mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
       tk.Label(
           frame_mid,
           text="请从 Excel 中复制“文件名（不含后缀）”这一列，粘贴到下面，每行一个：",
       ).pack(anchor="w")
       self.text_names = scrolledtext.ScrolledText(frame_mid, height=12)
       self.text_names.pack(fill=tk.BOTH, expand=True)
       # ============== 执行按钮 + 状态 ==============
       frame_bottom = tk.Frame(master)
       frame_bottom.pack(fill=tk.X, padx=10, pady=10)
       self.status_label = tk.Label(frame_bottom, text="就绪。")
       self.status_label.pack(side=tk.LEFT)
       tk.Button(frame_bottom, text="开始执行", command=self.run_task).pack(side=tk.RIGHT)
   # 选择根目录
   def choose_folder(self):
       folder = filedialog.askdirectory()
       if folder:
           self.root_dir.set(folder)
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
   # 主逻辑：扫描、删除/保留、扁平化子目录
   def run_task(self):
       root_dir = self.root_dir.get()
       if not root_dir or root_dir == "（未选择）":
           messagebox.showwarning("提示", "请先选择根目录。")
           return
       raw_names = self.text_names.get("1.0", tk.END)
       target_names = self.parse_name_list(raw_names)
       if not target_names:
           messagebox.showwarning("提示", "请先在文本框中粘贴至少一个文件名。")
           return
       mode = self.mode.get()
       # 简单确认提示
       confirm_msg = [
           f"根目录：{root_dir}",
           f"模式：{'保留这些名字（删除其他）' if mode == 'keep' else '删除这些名字（保留其他）'}",
           f"导入的目标名字数量：{len(target_names)}",
           "",
           "确认要执行吗？此操作会执行文件删除，请谨慎！",
       ]
       if not messagebox.askokcancel("确认执行", "\n".join(confirm_msg)):
           return
       self.status_label.config(text="正在扫描文件，请稍候...")
       self.master.update_idletasks()
       # 1. 收集所有文件
       all_files = []  # (full_path, base_name_lower)
       for dirpath, dirnames, filenames in os.walk(root_dir):
           for fname in filenames:
               full_path = os.path.join(dirpath, fname)
               base, ext = os.path.splitext(fname)
               all_files.append((full_path, base.lower()))
       total_files = len(all_files)
       if total_files == 0:
           messagebox.showinfo("结果", "在该根目录下没有找到任何文件。")
           self.status_label.config(text="没有文件。")
           return
       deleted = []
       kept = []
       # 为了统计“哪些名字没有匹配到任何文件”
       matched_name_set = set()
       # 2. 根据模式决定删除 / 保留
       for full_path, base_lower in all_files:
           # 是否属于“目标名字列表”
           in_list = base_lower in target_names
           if mode == "delete":
               # 删除这些名字
               if in_list:
                   try:
                       os.remove(full_path)
                       deleted.append(full_path)
                       matched_name_set.add(base_lower)
                   except Exception as e:
                       print(f"删除失败：{full_path}，错误：{e}")
               else:
                   kept.append(full_path)
           else:  # mode == 'keep'
               # 仅保留这些名字
               if in_list:
                   kept.append(full_path)
                   matched_name_set.add(base_lower)
               else:
                   try:
                       os.remove(full_path)
                       deleted.append(full_path)
                   except Exception as e:
                       print(f"删除失败：{full_path}，错误：{e}")
       # 3. 将“保留的文件”统一移动到根目录下（如果不在根目录）
       moved_count = 0
       rename_count = 0
       still_kept_paths = []
       for path in kept:
           dirpath, fname = os.path.split(path)
           # 已经在根目录：不动
           if os.path.abspath(dirpath) == os.path.abspath(root_dir):
               still_kept_paths.append(path)
               continue
           dest = os.path.join(root_dir, fname)
           # 如已有同名文件，自动在文件名后加 __数字
           if os.path.exists(dest):
               base, ext = os.path.splitext(fname)
               idx = 1
               while True:
                   new_name = f"{base}__{idx}{ext}"
                   dest2 = os.path.join(root_dir, new_name)
                   if not os.path.exists(dest2):
                       dest = dest2
                       rename_count += 1
                       break
                   idx += 1
           try:
               shutil.move(path, dest)
               moved_count += 1
               still_kept_paths.append(dest)
           except Exception as e:
               print(f"移动失败：{path} -> {dest}，错误：{e}")
               # 移动失败的文件也算“保留了但是没移动成功”
               still_kept_paths.append(path)
       # 4. 删除空子文件夹
       removed_dirs = 0
       for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
           if os.path.abspath(dirpath) == os.path.abspath(root_dir):
               continue
           # 只删除空目录
           try:
               if not os.listdir(dirpath):
                   os.rmdir(dirpath)
                   removed_dirs += 1
           except Exception:
               pass
       # 5. 找出用户给的名字中，哪些完全没匹配到任何文件
       unmatched_names = sorted(target_names - matched_name_set)
       # 统计结果
       deleted_count = len(deleted)
       kept_count = len(still_kept_paths)
       result_lines = [
           f"根目录：{root_dir}",
           f"模式：{'保留这些名字（删除其他）' if mode == 'keep' else '删除这些名字（保留其他）'}",
           f"总共扫描文件数：{total_files}",
           f"删除文件数：{deleted_count}",
           f"保留文件数：{kept_count}",
           f"其中，从子文件夹移动到根目录的文件数：{moved_count}",
           f"因重名而自动重命名的文件数：{rename_count}",
           f"删除的空子文件夹数量：{removed_dirs}",
       ]
       if unmatched_names:
           result_lines.append("")
           result_lines.append(
               f"下列 {len(unmatched_names)} 个名字在目录中没有找到对应文件（按“去后缀的文件名”匹配）："
           )
           # 避免一次性弹太长，只显示前 30 个
           preview = unmatched_names[:30]
           result_lines.extend("  - " + n for n in preview)
           if len(unmatched_names) > 30:
               result_lines.append(f"...（共 {len(unmatched_names)} 个未匹配名字）")
       messagebox.showinfo("执行结果", "\n".join(result_lines))
       self.status_label.config(text="完成。详情见弹窗。")

if __name__ == "__main__":
   root = tk.Tk()
   app = FileCleanerApp(root)
   root.mainloop()
