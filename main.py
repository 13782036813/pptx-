import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import zipfile
import shutil
import filetype
from datetime import datetime
import traceback
import platform

class PPTXVideoExtractor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("胡抓 v1.0")
        self.root.geometry("600x400")
        self.setup_ui()
        self.errors = []
        self.output_dir = ""
        self.current_operation = ""
        
        # 设置跨平台拖拽支持
        # self.setup_drag_drop()

    def setup_ui(self):
        """初始化界面组件"""
        style = ttk.Style()
        style.configure("TLabel", font=("微软雅黑", 10))
        style.configure("TButton", font=("微软雅黑", 10))
        
        # 主容器
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 拖拽/选择区域
        self.drop_area = ttk.Label(
            main_frame,
            text="拖拽PPTX文件到此处（Windows）\n或点击下方按钮选择文件",
            relief="groove",
            anchor="center",
            style="DropArea.TLabel"
        )
        self.drop_area.pack(fill=tk.BOTH, expand=True, pady=10)
        self.drop_area.bind("<Button-1>", self.on_select_file)

        # 进度条
        self.progress = ttk.Progressbar(
            main_frame,
            orient=tk.HORIZONTAL,
            length=500,
            mode='determinate'
        )
        self.progress.pack(pady=10)

        # 按钮容器
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)

        ttk.Button(
            btn_frame,
            text="选择文件",
            command=self.on_select_file
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame,
            text="打开输出目录",
            command=self.open_output_dir
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame,
            text="退出",
            command=self.root.quit
        ).pack(side=tk.RIGHT, padx=5)

        # 状态栏
        self.status_var = tk.StringVar()
        ttk.Label(
            main_frame,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        ).pack(fill=tk.X, pady=10)

        # 样式配置
        style.configure("DropArea.TLabel",
                       background="#f0f0f0",
                       padding=20,
                       wraplength=400,
                       font=("微软雅黑", 12))

    def setup_drag_drop(self):
        """增强拖拽支持"""
        if platform.system() == "Windows":
            try:
                self.root.drop_target_register(tk.DND_FILES)
                self.root.dnd_bind("<<Drop>>", self.on_drag_drop)
            except Exception as e:
                self.update_status(f"拖拽初始化失败: {str(e)}")
        else:
            self.drop_area["text"] = "点击按钮或区域选择文件"

    def on_drag_drop(self, event):
        """增强拖拽处理"""
        try:
            raw_path = event.data.replace("\\", "/").strip("{}")
            for path in raw_path.split("} {"):
                if os.path.isfile(path) and path.lower().endswith(".pptx"):
                    self.process_file(path)
        except Exception as e:
            self.errors.append(f"拖拽错误: {str(e)}")
            messagebox.showerror("拖拽错误", "请检查文件是否为有效PPTX")

    def on_select_file(self, event=None):
        """处理文件选择"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PPTX Files", "*.pptx"), ("All Files", "*.*")]
        )
        if file_path:
            self.process_file(file_path)

    def update_status(self, message):
        """更新状态栏"""
        self.status_var.set(message)
        self.root.update_idletasks()

    def create_output_dir(self):
        """创建输出目录"""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = os.path.join(desktop, f"PPTX_Videos_{timestamp}")
        os.makedirs(self.output_dir, exist_ok=True)
        return self.output_dir

    def safe_extract(self, zip_ref, extract_path):
        """安全解压文件"""
        try:
            total = len(zip_ref.namelist())
            for i, member in enumerate(zip_ref.namelist()):
                if member.startswith("ppt/media/"):
                    zip_ref.extract(member, extract_path)
                # 更新解压进度（0-30%）
                self.progress["value"] = (i+1)/total * 30
                self.update_status(f"解压文件中... ({i+1}/{total})")
            return True
        except Exception as e:
            self.errors.append(f"解压失败: {str(e)}")
            return False

    def is_video_file(self, filepath):
        """检测视频文件"""
        try:
            kind = filetype.guess(filepath)
            return bool(kind and kind.mime.startswith("video/"))
        except Exception as e:
            self.errors.append(f"文件检测失败: {os.path.basename(filepath)}")
            return False

    def handle_duplicate(self, filename):
        """处理重复文件名"""
        base, ext = os.path.splitext(filename)
        counter = 1
        new_name = filename
        while os.path.exists(os.path.join(self.output_dir, new_name)):
            new_name = f"{base}({counter}){ext}"
            counter += 1
        return new_name

    def process_file(self, pptx_path):
        """处理主逻辑"""
        if not pptx_path.lower().endswith(".pptx"):
            messagebox.showerror("错误", "仅支持.pptx文件")
            return

        self.errors.clear()
        temp_dir = "temp_pptx_extract"
        output_dir = self.create_output_dir()

        try:
            # 重置进度
            self.progress["value"] = 0
            self.update_status("开始处理...")

            # 步骤1：解压文件
            with zipfile.ZipFile(pptx_path, "r") as zip_ref:
                if not self.safe_extract(zip_ref, temp_dir):
                    return

            # 步骤2：处理媒体文件
            media_path = os.path.join(temp_dir, "ppt", "media")
            if not os.path.exists(media_path):
                self.errors.append("未找到媒体目录")
                return

            media_files = os.listdir(media_path)
            total = len(media_files)
            extracted = 0

            for i, filename in enumerate(media_files):
                src = os.path.join(media_path, filename)
                if self.is_video_file(src):
                    dest_name = self.handle_duplicate(filename)
                    shutil.copy2(src, os.path.join(output_dir, dest_name))
                    extracted += 1

                # 更新处理进度（30-100%）
                progress = 30 + (i+1)/total * 70
                self.progress["value"] = progress
                self.update_status(f"处理文件中... ({i+1}/{total})")

            # 显示结果
            result_msg = [
                f"处理完成！",
                f"成功提取: {extracted} 个视频",
                f"输出目录: {output_dir}"
            ]
            if self.errors:
                result_msg.append(f"\n遇到 {len(self.errors)} 个错误")

            messagebox.showinfo("完成", "\n".join(result_msg))

            # 显示错误详情
            if self.errors:
                self.show_error_details()

        except Exception as e:
            error_msg = f"严重错误: {str(e)}\n\n{traceback.format_exc()}"
            self.errors.append(error_msg)
            messagebox.showerror("运行时错误", error_msg)
        finally:
            # 清理临时文件
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            self.progress["value"] = 0
            self.update_status("就绪")

    def show_error_details(self):
        """显示错误详情窗口"""
        error_win = tk.Toplevel(self.root)
        error_win.title("错误详情")
        
        text_frame = ttk.Frame(error_win)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text = tk.Text(text_frame, wrap=tk.WORD, width=80, height=20)
        vsb = ttk.Scrollbar(text_frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=vsb.set)
        
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        text.insert(tk.END, "\n\n".join(self.errors))
        text.configure(state="disabled")
        
        ttk.Button(
            error_win,
            text="关闭",
            command=error_win.destroy
        ).pack(pady=5)

    def open_output_dir(self):
        """打开输出目录"""
        if os.path.exists(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showinfo("提示", "尚未生成输出目录")

if __name__ == "__main__":
    app = PPTXVideoExtractor()
    app.root.mainloop()