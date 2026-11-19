import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import subprocess
import threading
from enum import Enum

from utils import log, setup_crash_handler, tk_exception_hook, get_base_dir

# True 才真正拉起 drag.exe；False 时点击按钮弹"暂未接入"提示，
# 便于 drag.exe 未打包时先测试主程序其他路径
STARTING_DRAG = True


class AppState(Enum):
    # 三态互斥：非 IDLE 状态下禁止启动新任务。
    # 两个布尔（is_processing / is_dragging）也可以表达，但组合起来有 4 种状态
    # 其中 1 种是非法组合；枚举从结构上排除非法态
    IDLE = "idle"
    DRAGGING = "dragging"
    PROCESSING = "processing"


def cross_platform_select_file(parent_window):
    dialog_kwargs = {
        "title": "选择需要识别的文档或图片",
        # 空格分隔的多 glob：macOS/Linux 的 Tk 不认分号，
        # 分号在 Windows 上能工作但在其他平台会导致过滤框为空
        "filetypes": [
            ("支持的文件", "*.pdf *.png *.jpg *.jpeg"),
            ("PDF 文件", "*.pdf"),
            ("图片文件", "*.png *.jpg *.jpeg"),
            ("所有文件", "*.*")
        ],
        # parent：macOS 上不传 parent 会导致弹窗跑到主窗口后面，
        # 用户以为程序卡死
        "parent": parent_window,
    }
    try:
        if sys.platform.startswith("win"):
            dialog_kwargs["initialdir"] = os.path.expanduser("~/Desktop")
            return filedialog.askopenfilename(**dialog_kwargs)

        elif sys.platform == "darwin":
            try:
                return filedialog.askopenfilename(**dialog_kwargs)
            except Exception as e:
                # macOS 某些 Tk 版本在特殊权限下会底层报错，
                # 直接捕获避免闪退到终端
                log.error("macOS 文件选择器失败: %s", e, exc_info=True)
                messagebox.showerror("错误", "文件选择器不可用，请检查系统权限。")
                return None

        elif sys.platform.startswith("linux"):
            # 无 DISPLAY 也没有 WAYLAND_DISPLAY 说明是纯命令行环境，
            # 调 askopenfilename 会直接抛 TclError
            if not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY"):
                messagebox.showwarning("环境警告", "当前环境缺少图形界面支持！")
                return None
            return filedialog.askopenfilename(**dialog_kwargs)

        else:
            return filedialog.askopenfilename(**dialog_kwargs)

    except Exception as e:
        log.error("文件选择器发生错误: %s", str(e), exc_info=True)
        messagebox.showerror("错误", f"无法打开文件选择器:\n{e}")
        return None


class PDF2WordApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("通用文档 OCR 转换工具")
        self.geometry("550x420")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # 接管 Tkinter 按钮回调异常：tkinter 默认会把异常打到 stderr 后继续跑，
        # --noconsole 打包后 stderr 是 None，回调异常会静默丢失。
        # 挂 tk_exception_hook 后所有回调异常会写入日志
        self.report_callback_exception = tk_exception_hook

        # 【关键】不能用 self.state：会覆盖 tkinter.Tk.state() 方法，
        # CustomTkinter 内部调 self.state(...) 时会因 AppState 枚举不可调用而崩
        self._app_state = AppState.IDLE
        self.file_list: list[str] = []
        # _closing 和 _confirming_close 两个标志语义不同：
        # - _closing = True 表示窗口即将销毁，_safe_after 调度和执行都会拦截
        # - _confirming_close = True 表示正在弹关闭确认框，
        #   此时后台线程完成转换不应该弹"成功"提示框，否则会与确认框叠在一起
        self._closing = False
        self._confirming_close = False
        self._drag_proc = None
        self._cancel_event = threading.Event()
        # _engine_loading 反映"模型是否正在加载"，
        # 取消时用它决定显示"引擎加载中，取消将在初始化完成后生效"
        # 还是"正在结束当前操作"
        self._engine_loading = False

        self._start_btn_default_fg = None
        self._start_btn_default_hover = None

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()

        # 启动按钮初始颜色记录：恢复按钮状态时不硬编码颜色，
        # 未来切换主题（如绿色主题）后恢复颜色不会脱节
        self._start_btn_default_fg = self.start_btn.cget("fg_color")
        self._start_btn_default_hover = self.start_btn.cget("hover_color")

    def _build_ui(self):
        ctk.CTkLabel(self, text="通用文档 OCR 转换器", font=("Microsoft YaHei", 22, "bold")).pack(pady=(20, 10))

        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(fill="x", padx=20, pady=10)

        # 多行 Textbox 而非单行 Entry：超长路径（如 Android 沙盒路径）会被 Entry
        # 撑爆或截断，Textbox 自动换行 + 可滚动
        self.file_textbox = ctk.CTkTextbox(self.file_frame, height=65, width=380, font=("Arial", 12), wrap="char")
        self.file_textbox.pack(side="left", padx=(10, 5), pady=10)
        self.file_textbox.insert("0.0", "尚未选择文件（支持 PDF / PNG / JPG / JPEG）")
        self.file_textbox.configure(state="disabled")

        self.browse_btn = ctk.CTkButton(self.file_frame, text="浏览...", width=80, command=self._select_file)
        self.browse_btn.pack(side="left", padx=(5, 10), pady=10)

        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.pack(fill="x", padx=20, pady=15)
        self.drag_btn = ctk.CTkButton(self.action_frame, text="独立弹窗拖拽", width=150, fg_color="#555555", command=self._open_drag_window)
        self.drag_btn.pack(side="left", padx=10)
        self.start_btn = ctk.CTkButton(self.action_frame, text="开始转换", width=150, font=("Microsoft YaHei", 15, "bold"), command=self._start_task)
        self.start_btn.pack(side="right", padx=10)

        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=20, pady=20)
        # 纯文字百分比而非进度条：进度条不够精确（用户看不出 47% 还是 48%），
        # 大字号数字更符合"精确卡点"的需求
        self.percent_label = ctk.CTkLabel(self.progress_frame, text="0%", font=("Arial", 56, "bold"), text_color="#2ECC71")
        self.percent_label.pack(pady=(5, 0))
        self.status_label = ctk.CTkLabel(self.progress_frame, text="就绪，请选择文件", font=("Microsoft YaHei", 14), text_color="gray")
        self.status_label.pack(pady=5)

    def _safe_after(self, fn, *args, **kwargs):
        # 所有后台线程更新 UI 都必须走这里，否则 tkinter 会因非主线程调用抛 TclError。
        # 双层 _closing 检查：
        # - 调度时窗口已关闭 → 直接丢弃，不排进 after 队列
        # - 排队后、真正执行前窗口关闭 → wrapper 内再检查一次，静默跳过
        if self._closing:
            return
        def wrapper():
            if self._closing:
                return
            try:
                fn(*args, **kwargs)
            except tk.TclError:
                # 窗口销毁边缘的 TclError（如 configure 已失效）不刷屏
                pass
        try:
            self.after(0, wrapper)
        except tk.TclError:
            pass

    def _is_valid_file(self, path: str) -> bool:
        # 白名单与 drag.py 的 VALID_EXTS 必须保持同步
        if not os.path.isfile(path):
            return False
        valid_exts = ('.pdf', '.png', '.jpg', '.jpeg')
        return path.lower().endswith(valid_exts)

    def _set_file_path(self, path: str, status_text: str = "已选择文件，点击开始转换"):
        try:
            self.file_textbox.configure(state="normal")
            self.file_textbox.delete("0.0", "end")
            self.file_textbox.insert("0.0", path)
            # self.file_list 在 insert 成功后才更新：
            # 若 insert 抛异常（极少见），file_list 不更新，
            # 保证"内部状态"与"UI 显示"始终一致
            self.file_list = [path]
        except Exception as e:
            log.error("更新文件路径失败", exc_info=True)
            self.file_list = []
            try:
                self.file_textbox.delete("0.0", "end")
                self.file_textbox.insert("0.0", "尚未选择文件（写入失败）")
            except Exception:
                pass
            return
        finally:
            self.file_textbox.configure(state="disabled")
        self.status_label.configure(text=status_text, text_color="#2ECC71")

    def _select_file(self):
        if self._app_state != AppState.IDLE:
            return
        path = cross_platform_select_file(self)
        if path:
            if not self._is_valid_file(path):
                messagebox.showwarning("格式错误", "请选择有效的文档或图片文件！")
                return
            self._set_file_path(path)

    def _open_drag_window(self):
        if self._app_state != AppState.IDLE:
            return
        if not STARTING_DRAG:
            messagebox.showinfo("提示", "拖拽模块暂未接入！")
            return
        base_dir = get_base_dir()
        drag_exe = os.path.join(base_dir, "drag.exe")
        if not os.path.exists(drag_exe):
            messagebox.showerror("错误", f"找不到拖拽组件: {drag_exe}")
            return

        self._app_state = AppState.DRAGGING
        self.drag_btn.configure(state="disabled", text="等待拖拽中...")
        self.start_btn.configure(state="disabled")
        self.status_label.configure(text="请在弹出的窗口中拖拽文件...", text_color="gray")

        def run_drag_thread():
            # 双通道通讯的清理阶段：启动 drag.exe 前先删掉上次可能残留的 drop_path.txt，
            # 否则若本次 drag.exe 未写文件（用户取消），会读到上一次的旧路径
            txt_path = os.path.join(base_dir, "drop_path.txt")
            try:
                if os.path.exists(txt_path):
                    os.remove(txt_path)
            except OSError:
                pass

            proc = None
            try:
                flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
                # text=True + encoding="utf-8"：
                # drag.exe 输出中文路径时，若用 locale 默认编码（Windows GBK），
                # 主进程按 UTF-8 解码会得到 \ufffd 替换字符，路径失效
                # errors="replace" 兜底：即便有非 UTF-8 字节也不崩
                proc = subprocess.Popen(
                    [drag_exe],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    text=True, encoding="utf-8", errors="replace",
                    creationflags=flags,
                )
                self._drag_proc = proc

                # 竞态：Popen 成功后，主线程可能在 communicate 期间调用 _on_close 清理。
                # 先检查 _closing，命中时立即 terminate 当前 proc 并返回，
                # 避免主进程退出后 drag.exe 变孤儿进程
                if self._closing:
                    try:
                        proc.terminate()
                        proc.wait(timeout=0.5)
                    except subprocess.TimeoutExpired:
                        try:
                            proc.kill()
                            proc.wait(timeout=0.5)
                        except Exception:
                            pass
                    except Exception:
                        pass
                    return

                # communicate 阻塞等 drag.exe 退出，timeout=300 秒
                # 这是主线程之外的后台线程，不会卡死 UI
                stdout, stderr = proc.communicate(timeout=300)
                returncode = proc.returncode

                if stderr.strip():
                    # 用退出码区分严重程度，而非匹配 stderr 文本：
                    # drag.exe 正常运行时也可能往 stderr 打调试信息，
                    # 匹配关键词会误报
                    if returncode != 0:
                        log.warning("drag.exe 非零退出 (code=%s): %s", returncode, stderr.strip())
                    else:
                        log.info("drag.exe 正常退出，stderr: %s", stderr.strip())

                # 双通道协议第一层：stdout 优先
                # lstrip("\ufeff\u200b\u200c\u200d")：去 BOM 和零宽字符，
                # 防止路径前有不可见字符导致 os.path.isfile 返回 False
                dropped_path = stdout.strip().lstrip("\ufeff\u200b\u200c\u200d")
                if "\n" in dropped_path:
                    # drag.exe 若同时输出多个路径（例如拖入多文件），
                    # 只取第一个
                    log.warning("drag.exe 输出多个路径，只取第一个: %r", dropped_path)
                    dropped_path = dropped_path.splitlines()[0].strip()
                if "\ufffd" in dropped_path:
                    # \ufffd = Unicode 替换字符，说明 stdout 某处编码失真
                    log.warning("编码异常: %r", dropped_path)

                # 双通道协议第二层：stdout 为空时读 txt 兜底
                # 场景：--noconsole 打包后 drag.exe 的 stdout 是 None，
                # 它只会写 txt，主进程必须能通过 txt 拿到路径
                if not dropped_path and os.path.exists(txt_path):
                    try:
                        with open(txt_path, "r", encoding="utf-8") as f:
                            dropped_path = f.read().strip().lstrip("\ufeff\u200b\u200c\u200d")
                    except OSError as e:
                        log.warning("读取 drop_path.txt 失败: %s", e)

                # 拿到路径后删掉 txt，避免残留污染下一次拖拽
                if dropped_path and os.path.exists(txt_path):
                    try:
                        os.remove(txt_path)
                    except OSError:
                        pass

                if dropped_path and self._is_valid_file(dropped_path):
                    self._safe_after(self._set_file_path, dropped_path, "拖拽成功！")
                elif not dropped_path and returncode != 0:
                    # 无路径 + 非零退出 = drag.exe 本身崩溃
                    self._safe_after(self.status_label.configure, text="拖拽工具异常退出，请查看日志", text_color="#E74C3C")
                else:
                    # 无路径 + 零退出 = 用户取消，或拖了不支持的文件格式
                    # Windows 260 字符路径限制在 Python 层不报错但下游 API 会拒绝，
                    # 出现时日志记录便于排查
                    if dropped_path and len(dropped_path) > 260:
                        log.warning("路径超过 260 字符，可能被 Windows API 拒绝: %r", dropped_path)
                    self._safe_after(self.status_label.configure, text="未获取到有效文件", text_color="gray")

            except subprocess.TimeoutExpired:
                # 超时后必须 kill+wait，否则 Linux 上子进程变僵尸
                if proc is not None:
                    try:
                        proc.kill()
                        proc.wait(timeout=0.5)
                    except Exception:
                        pass
                self._safe_after(self.status_label.configure, text="拖拽超时", text_color="#E74C3C")
            except Exception as e:
                log.error("拉起拖拽失败: %s", e)
                self._safe_after(self.status_label.configure, text="拖拽启动失败", text_color="#E74C3C")
            finally:
                self._drag_proc = None
                self._safe_after(self._on_drag_done)

        try:
            threading.Thread(target=run_drag_thread, daemon=True).start()
        except RuntimeError:
            # RuntimeError 常见于"can't start new thread"（线程耗尽），
            # 此时必须恢复 UI 状态，否则按钮永久禁用
            log.error("无法启动拖拽线程", exc_info=True)
            self._app_state = AppState.IDLE
            self.drag_btn.configure(state="normal", text="独立弹窗拖拽")
            self.start_btn.configure(
                state="normal", text="开始转换", command=self._start_task,
                fg_color=self._start_btn_default_fg, hover_color=self._start_btn_default_hover,
            )
            messagebox.showerror("错误", "系统资源不足，无法启动拖拽线程")

    def _on_drag_done(self):
        self._app_state = AppState.IDLE
        self.drag_btn.configure(state="normal", text="独立弹窗拖拽")
        # 拖拽期间只禁用 start_btn 的 state，不改 text/command，
        # 所以恢复时也只恢复 state，不重复设置其他属性
        self.start_btn.configure(
            state="normal", text="开始转换", command=self._start_task,
            fg_color=self._start_btn_default_fg, hover_color=self._start_btn_default_hover,
        )

    def _start_task(self):
        if self._app_state != AppState.IDLE:
            return

        if not self.file_list:
            messagebox.showwarning("提示", "请先选择一个文件！")
            return
        if not self._is_valid_file(self.file_list[0]):
            messagebox.showerror("错误", "当前路径不是支持的文件格式！")
            return

        self._app_state = AppState.PROCESSING
        # clear() 而非创建新 Event：复用同一个 Event 对象，
        # 避免上一轮的 _cancel_event 引用残留在其他线程里
        self._cancel_event.clear()
        self._engine_loading = False

        # 开始转换后按钮立即变红色"取消转换"，用户可以随时打断长任务
        self.start_btn.configure(state="normal", text="取消转换", command=self._cancel_task, fg_color="#E74C3C", hover_color="#C0392B")
        self.percent_label.configure(text="0%", text_color="#F1C40F")
        self.status_label.configure(text="正在准备 OCR 引擎...", text_color="gray")

        try:
            threading.Thread(target=self._run_conversion_thread, daemon=True).start()
        except RuntimeError:
            log.error("无法启动转换线程", exc_info=True)
            self._reset_task_state(False, "系统资源不足，无法启动转换线程", "失败")

    def _cancel_task(self):
        if self._closing:
            return
        self._cancel_event.set()
        self.start_btn.configure(state="disabled", text="取消中...")

        # 两种文案区分阶段：
        # - 模型加载中（_engine_loading=True）：PaddleOCR 是 C++ 调用，Python 层
        #   cancel_event 无法中断正在加载的模型，用户需等待当前加载完成
        # - 转换中：页面边界检查 cancel，响应快
        if self._engine_loading:
            self.status_label.configure(text="引擎加载中，取消将在初始化完成后生效...", text_color="#F39C12")
        else:
            self.status_label.configure(text="正在结束当前操作，请稍候...", text_color="gray")

    def _run_conversion_thread(self):
        # 延迟导入：OCR 引擎依赖 PaddleOCR/Paddle，导入耗时 2-5 秒，
        # 放在 __init__ 里会让 GUI 启动卡住。
        # 单独 try/except 很重要：ocr_engine 导入失败抛 ImportError 会绕开后面的
        # except ConversionCancelled/except Exception 的匹配（ConversionCancelled
        # 名字未定义），如果导入失败没有被这里捕获，整个线程会静默死掉
        try:
            from ocr_engine import get_engine, ConversionCancelled, is_engine_loaded
        except Exception as e:
            log.error("OCR 模块导入失败", exc_info=True)
            self._safe_after(self._reset_task_state, False, f"OCR 模块加载失败，请检查依赖是否安装：{e}", "失败")
            return

        try:
            if self._cancel_event.is_set():
                raise ConversionCancelled("启动前已取消")

            need_load = not is_engine_loaded()
            if need_load:
                self._engine_loading = True
                self._safe_after(self.status_label.configure, text="正在加载 OCR 模型（首次较慢，请稍候）...", text_color="gray")
            else:
                self._safe_after(self.status_label.configure, text="正在启动 OCR 引擎...", text_color="gray")

            # 【P2 修复】get_engine() 是毫秒级返回（只是创建实例，懒加载），
            # 模型真正加载发生在 convert_pdf/convert_image 内部。
            # 所以不能在这里复位 _engine_loading——那会让"引擎加载中"文案
            # 在模型真正开始加载之前就消失，用户取消时看到错误提示
            engine = get_engine()

            if self._cancel_event.is_set():
                raise ConversionCancelled("初始化后取消")

            input_path = self.file_list[0]
            output_dir = os.path.dirname(input_path)

            # 【P2 修复】首次 progress_cb 触发时复位 _engine_loading：
            # progress_cb 第一次被调用说明模型已加载完、第一页/图片处理完。
            # 用 list 包一层而非 nonlocal 闭包，兼容 Python 3.10 之前的旧版本
            first_progress = [True]
            def progress_cb(current, total):
                if first_progress[0]:
                    first_progress[0] = False
                    self._engine_loading = False
                percent = int(current / total * 100) if total > 0 else 0
                self._safe_after(self.percent_label.configure, text=f"{percent}%")
                self._safe_after(self.status_label.configure, text=f"正在识别：第 {current}/{total} 页")

            if input_path.lower().endswith(".pdf"):
                out_path = engine.convert_pdf(input_path, output_dir, progress_cb, self._cancel_event)
            else:
                # hasattr 防御：如果未来 ocr_engine 删除了 convert_image（理论上不会），
                # 这里会给出明确提示而非 AttributeError
                if not hasattr(engine, "convert_image"):
                    raise NotImplementedError("图片转换功能暂未接入，请先使用 PDF 文件进行转换！")
                out_path = engine.convert_image(input_path, output_dir, progress_cb, self._cancel_event)

            # 引擎契约要求成功时返回非空路径。None 说明内部异常被静默吞掉，
            # 在 GUI 层抛出避免"转换完成但找不到文件"的悬案
            if not out_path:
                raise RuntimeError("转换引擎未返回输出路径（可能内部异常被静默吞掉）")

            self._safe_after(self._reset_task_state, True, out_path, "100%")

        except ConversionCancelled:
            # 用户主动取消：灰色静默提示，不弹错误框
            self._safe_after(self._reset_task_state, False, "已取消", "取消", cancelled=True)
        except NotImplementedError as e:
            # 功能未实现是预期路径而非错误：info 级日志，
            # 走 cancelled=True 的灰色提示而非红色失败
            log.info("功能未实现: %s", e)
            self._safe_after(self._reset_task_state, False, str(e), "提示", cancelled=True)
        except Exception as e:
            # 双重区分：cancel_event 已置位说明异常是取消的副作用（如 PaddleOCR 内部
            # 被中断抛的 RuntimeError），仍走取消分支；否则才是真正的失败
            if self._cancel_event.is_set():
                self._safe_after(self._reset_task_state, False, "已取消", "取消", cancelled=True)
            else:
                log.error("转换异常", exc_info=True)
                self._safe_after(self._reset_task_state, False, f"转换失败: {e}", "失败")

    def _reset_task_state(self, success: bool, message: str, percent_text: str, cancelled: bool = False):
        if self._closing:
            return

        self._app_state = AppState.IDLE
        # 【兜底】异常路径或提前取消时 _engine_loading 可能仍为 True，
        # 下次点"开始"时 need_load 判断会基于错误的 True 值走错分支
        self._engine_loading = False
        self.start_btn.configure(
            state="normal", text="开始转换", command=self._start_task,
            fg_color=self._start_btn_default_fg, hover_color=self._start_btn_default_hover,
        )
        self.percent_label.configure(text=percent_text)

        if success:
            self.percent_label.configure(text_color="#2ECC71")
            self.status_label.configure(text="转换完成！", text_color="#2ECC71")
        elif cancelled:
            self.percent_label.configure(text_color="gray")
            self.status_label.configure(text=message, text_color="gray")
        else:
            self.percent_label.configure(text_color="#E74C3C")
            self.status_label.configure(text=message, text_color="#E74C3C")

        # 关闭确认框正弹出时不弹成功/失败框，避免两个模态框叠在一起
        if self._confirming_close:
            return

        if success:
            messagebox.showinfo("成功", f"转换完成！\n\n输出文件：\n{message}")
        elif not cancelled:
            messagebox.showerror("失败", f"处理过程中发生错误。\n{message}")

    def _on_close(self):
        # 防狂点 X：_confirming_close 为 True 时后续点击直接忽略
        if self._confirming_close:
            return
        self._confirming_close = True

        try:
            if self._app_state != AppState.IDLE:
                try:
                    confirmed = messagebox.askyesno("确认", "任务进行中，确定要退出吗？")
                except Exception as e:
                    # 弹框自身异常时保守不关闭：宁可让用户再点一次 X，
                    # 也不要因弹框异常丢正在跑的任务
                    log.error("关闭确认框异常", exc_info=True)
                    confirmed = False

                if not confirmed:
                    return

            self._closing = True
        finally:
            # 无论确认结果如何，_confirming_close 必须在 finally 里复位，
            # 否则用户点"否"后 _confirming_close 永久为 True，下次点 X 无响应
            self._confirming_close = False

        # 接管 _drag_proc 引用：先取出再置 None，
        # 避免后台线程在 finally 里把 self._drag_proc 置空导致后面调用 None.terminate()
        proc = self._drag_proc
        self._drag_proc = None

        if proc and proc.poll() is None:
            try:
                proc.terminate()
                proc.wait(timeout=0.5)
            except subprocess.TimeoutExpired:
                try:
                    proc.kill()
                    proc.wait(timeout=0.5)
                except Exception:
                    pass
            except Exception as e:
                log.error("清理子进程失败: %s", e)
                try:
                    proc.kill()
                except Exception:
                    pass

        # 通知转换线程退出：线程会在页面边界响应，最坏等情况当前页处理完
        self._cancel_event.set()
        self.destroy()


if __name__ == "__main__":
    # 第一件事挂载崩溃拦截器：后续任何未捕获异常（主线程/子线程/Tkinter回调）
    # 都会写日志 + 弹记事本，而不是闪退到桌面
    setup_crash_handler()
    app = PDF2WordApp()
    app.mainloop()