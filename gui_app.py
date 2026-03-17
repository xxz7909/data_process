"""
股票图片OCR批量识别工具 - GUI版本
使用 tkinter 构建界面，调用 ocr_processor 进行处理
"""
import os
import sys
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 获取脚本所在目录（支持打包后的路径）
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 确保能导入同目录下的模块
sys.path.insert(0, BASE_DIR)
from ocr_processor import process_images, detect_groups

# 获取当前Windows用户的真实桌面路径
def _get_desktop_dir():
    """动态获取当前登录用户的桌面目录（支持OneDrive重定向等场景）"""
    try:
        import ctypes.wintypes
        CSIDL_DESKTOPDIRECTORY = 0x0010
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, 0, buf)
        if buf.value:
            return buf.value
    except Exception:
        pass
    # 回退方案
    return os.path.join(os.path.expanduser('~'), 'Desktop')

DESKTOP_DIR = _get_desktop_dir()

# 配置文件路径（记忆上次选择）
CONFIG_FILE = os.path.join(BASE_DIR, '.ocr_config.json')


def load_config():
    """加载上次的路径配置"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_config(cfg):
    """保存路径配置"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# 默认图片目录：上级目录的 pics 文件夹
DEFAULT_PICS_DIR = os.path.join(os.path.dirname(BASE_DIR), 'pics')


class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("股票图片OCR批量识别工具")
        self.root.geometry("620x520")
        self.root.resizable(False, False)

        # 设置图标和字体
        self.default_font = ("微软雅黑", 10)
        self.root.option_add("*Font", self.default_font)

        # 加载上次配置
        self.config = load_config()

        self._build_ui()
        self._setup_drag_drop()
        self.processing = False

    def _build_ui(self):
        # ---- 图片目录选择 ----
        frame_dir = ttk.LabelFrame(self.root, text="图片目录（可拖拽文件夹到输入框）", padding=10)
        frame_dir.pack(fill="x", padx=15, pady=(15, 5))

        default_dir = self.config.get('img_dir', DEFAULT_PICS_DIR)
        self.dir_var = tk.StringVar(value=default_dir)
        self.entry_dir = tk.Entry(frame_dir, textvariable=self.dir_var, width=50,
                                  font=("微软雅黑", 9))
        self.entry_dir.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_browse = ttk.Button(frame_dir, text="浏览...", command=self._browse_dir)
        btn_browse.pack(side="right")

        # ---- 输出文件 ----
        frame_out = ttk.LabelFrame(self.root, text="输出Excel文件", padding=10)
        frame_out.pack(fill="x", padx=15, pady=5)

        default_out = self.config.get('output_path', '')
        # 如果记忆的输出路径所在目录不存在（换了用户/电脑），回退到当前用户桌面
        if not default_out or not os.path.isdir(os.path.dirname(default_out)):
            default_out = os.path.join(DESKTOP_DIR, 'result.xlsx')
        self.out_var = tk.StringVar(value=default_out)
        entry_out = ttk.Entry(frame_out, textvariable=self.out_var, width=50)
        entry_out.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_out = ttk.Button(frame_out, text="浏览...", command=self._browse_output)
        btn_out.pack(side="right")

        # ---- 信息区域 ----
        frame_info = ttk.LabelFrame(self.root, text="处理信息", padding=10)
        frame_info.pack(fill="both", expand=True, padx=15, pady=5)

        self.info_text = tk.Text(frame_info, height=10, state="disabled",
                                 wrap="word", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(frame_info, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.info_text.pack(fill="both", expand=True)

        # ---- 进度条 ----
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var,
                                            maximum=100, length=400)
        self.progress_bar.pack(padx=15, pady=5, fill="x")

        # ---- 状态标签 ----
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(self.root, textvariable=self.status_var,
                                 foreground="gray")
        status_label.pack(padx=15)

        # ---- 按钮区域 ----
        frame_btn = ttk.Frame(self.root, padding=10)
        frame_btn.pack(fill="x", padx=15, pady=(0, 15))

        self.btn_start = ttk.Button(frame_btn, text="开始识别", command=self._start)
        self.btn_start.pack(side="left", padx=5)

        btn_open = ttk.Button(frame_btn, text="打开结果文件", command=self._open_result)
        btn_open.pack(side="left", padx=5)

        btn_exit = ttk.Button(frame_btn, text="退出", command=self.root.quit)
        btn_exit.pack(side="right", padx=5)

    def _setup_drag_drop(self):
        """使用 windnd 实现 Windows 原生文件拖拽"""
        try:
            import windnd

            def on_drop_dir(files):
                """拖拽到图片目录输入框"""
                if files:
                    path = files[0].decode('gbk') if isinstance(files[0], bytes) else str(files[0])
                    if os.path.isdir(path):
                        self.dir_var.set(path)
                    elif os.path.isfile(path):
                        self.dir_var.set(os.path.dirname(path))
                    self._save_paths()

            def on_drop_out(files):
                """拖拽到输出文件输入框（取目录）"""
                if files:
                    path = files[0].decode('gbk') if isinstance(files[0], bytes) else str(files[0])
                    if os.path.isdir(path):
                        # 拖入目录则用默认文件名
                        self.out_var.set(os.path.join(path, 'result.xlsx'))
                    elif os.path.isfile(path):
                        self.out_var.set(path)
                    self._save_paths()

            windnd.hook_dropfiles(self.entry_dir, func=on_drop_dir)
            windnd.hook_dropfiles(self.root, func=on_drop_dir)  # 整个窗口也支持拖拽
        except ImportError:
            pass  # windnd 不可用则忽略

    def _browse_dir(self):
        current = self.dir_var.get().strip()
        init_dir = current if os.path.isdir(current) else BASE_DIR
        d = filedialog.askdirectory(title="选择图片目录", initialdir=init_dir)
        if d:
            self.dir_var.set(d)
            self._save_paths()

    def _browse_output(self):
        current = self.out_var.get().strip()
        init_dir = os.path.dirname(current) if current else DESKTOP_DIR
        init_file = os.path.basename(current) if current else 'result.xlsx'
        f = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialdir=init_dir,
            initialfile=init_file
        )
        if f:
            self.out_var.set(f)
            self._save_paths()

    def _save_paths(self):
        """保存当前路径到配置文件"""
        self.config['img_dir'] = self.dir_var.get().strip()
        self.config['output_path'] = self.out_var.get().strip()
        save_config(self.config)

    def _log(self, msg):
        """向信息区域追加一行文字（线程安全）"""
        def _append():
            self.info_text.configure(state="normal")
            self.info_text.insert("end", msg + "\n")
            self.info_text.see("end")
            self.info_text.configure(state="disabled")
        self.root.after(0, _append)

    def _update_progress(self, current, total, msg):
        """进度回调"""
        pct = (current / total) * 100 if total > 0 else 0
        self.root.after(0, lambda: self.progress_var.set(pct))
        self.root.after(0, lambda: self.status_var.set(msg))
        self._log(f"[{current}/{total}] {msg}")

    def _start(self):
        if self.processing:
            return

        img_dir = self.dir_var.get().strip()
        output_path = self.out_var.get().strip()

        # 验证输入
        if not img_dir or not os.path.isdir(img_dir):
            messagebox.showerror("错误", f"图片目录不存在：\n{img_dir}")
            return

        group_count = detect_groups(img_dir)
        if group_count == 0:
            messagebox.showerror("错误",
                                 f"目录中未找到有效图片文件\n"
                                 f"（格式：N_差额.png.bmp 或 N_差额.bmp）\n"
                                 f"（代码/市值文件同样支持上述两种命名）\n\n"
                                 f"目录：{img_dir}")
            return

        if not output_path:
            messagebox.showerror("错误", "请指定输出文件路径")
            return

        # 检查输出文件是否被占用
        if os.path.exists(output_path):
            try:
                with open(output_path, 'a'):
                    pass
            except PermissionError:
                messagebox.showerror("错误",
                                     f"输出文件已被其他程序占用，请关闭后重试：\n{output_path}")
                return

        # 清空信息区域
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.configure(state="disabled")
        self.progress_var.set(0)

        # 保存本次路径选择
        self._save_paths()

        self._log(f"检测到 {group_count} 组图片，开始处理...")
        self.processing = True
        self.btn_start.configure(state="disabled")

        # 在子线程中运行OCR
        thread = threading.Thread(target=self._run_ocr, args=(img_dir, output_path),
                                  daemon=True)
        thread.start()

    def _run_ocr(self, img_dir, output_path):
        """子线程中执行OCR处理"""
        try:
            success, message = process_images(img_dir, output_path, self._update_progress)
            self.root.after(0, lambda: self._on_complete(success, message))
        except Exception as e:
            self.root.after(0, lambda: self._on_complete(False, f"处理出错：{e}"))

    def _on_complete(self, success, message):
        """处理完成回调"""
        self.processing = False
        self.btn_start.configure(state="normal")
        self.progress_var.set(100 if success else 0)
        self.status_var.set("完成！" if success else "失败")
        self._log(message)

        if success:
            messagebox.showinfo("完成", message)
        else:
            messagebox.showerror("失败", message)

    def _open_result(self):
        """打开结果Excel文件"""
        path = self.out_var.get().strip()
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showwarning("提示", "结果文件尚未生成，请先执行识别。")


def main():
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
