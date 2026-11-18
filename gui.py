import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import subprocess
import threading
from enum import Enum

from utils import log, setup_crash_handler, tk_exception_hook, get_base_dir
# 引入 lang.py 的接口
from lang import t, set_lang, load_lang_from_file, save_lang_to_file

# True 才真正拉起 drag.exe；False 时点击按钮弹"暂未接入"提示
STARTING_DRAG = True


class AppState(Enum):
    IDLE = "idle"
    DRAGGING = "dragging"
    PROCESSING = "processing"


def cross_platform_select_file(parent_window):
    dialog_kwargs = {
        "title": t("dialog_title"),
        "filetypes": [
            (t("filter_supported"), "*.pdf *.png *.jpg *.jpeg"),
            (t("filter_pdf"), "*.pdf"),
            (t("filter_image"), "*.png *.jpg *.jpeg"),
            (t("filter_all"), "*.*")
        ],
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
                log.error("macOS 文件选择器失败: %s", e, exc_info=True)
                messagebox.showerror(t("msg_error"), t("msg_mac_picker_fail"))
                return None

        elif sys.platform.startswith("linux"):
            if not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY"):
                messagebox.showwarning(t("msg_tip"), t("msg_no_display"))
                return None
            return filedialog.askopenfilename(**dialog_kwargs)

        else:
            return filedialog.askopenfilename(**dialog_kwargs)

    except Exception as e:
        log.error("文件选择器发生错误: %s", str(e), exc_info=True)
        messagebox.showerror(t("msg_error"), t("msg_picker_crash", err=e))
        return None


class PDF2WordApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("550x420")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.report_callback_exception = tk_exception_hook

        # 1. 从物理文件加载语言
        self._lang = load_lang_from_file()

        self._app_state = AppState.IDLE
        self.file_list: list[str] = []
        self._closing = False
        self._confirming_close = False
        self._drag_proc = None
        self._cancel_event = threading.Event()
        self._engine_loading = False

        self._start_btn_default_fg = None
        self._start_btn_default_hover = None

        # 状态栏语义记录：切语言时按 key 重新生成
        self._status_key = "status_ready"
        self._status_args = {}
        self._status_color = "gray"

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()

        # 2. 构建完毕后应用一次语言（会基于 _status_key 刷出初始状态文案）
        self._apply_language()

        self._start_btn_default_fg = self.start_btn.cget("fg_color")
        self._start_btn_default_hover = self.start_btn.cget("hover_color")

    def _build_ui(self):
        self.title_label = ctk.CTkLabel(self, text="", font=("Microsoft YaHei", 22, "bold"))
        self.title_label.pack(pady=(20, 10))

        # 语言切换下拉框
        self.lang_menu = ctk.CTkOptionMenu(
            self, values=["zh", "en"], command=self._change_language, width=80, height=28
        )
        self.lang_menu.pack(anchor="ne", padx=20, pady=(0, 0))
        self.lang_menu.set(self._lang)

        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(fill="x", padx=20, pady=10)

        self.file_textbox = ctk.CTkTextbox(self.file_frame, height=65, width=380, font=("Arial", 12), wrap="char")
        self.file_textbox.pack(side="left", padx=(10, 5), pady=10)
        self.file_textbox.insert("0.0", "")
        self.file_textbox.configure(state="disabled")

        self.browse_btn = ctk.CTkButton(self.file_frame, text="", width=80, command=self._select_file)
        self.browse_btn.pack(side="left", padx=(5, 10), pady=10)

        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.pack(fill="x", padx=20, pady=15)
        self.drag_btn = ctk.CTkButton(self.action_frame, text="", width=150, fg_color="#555555", command=self._open_drag_window)
        self.drag_btn.pack(side="left", padx=10)
        self.start_btn = ctk.CTkButton(self.action_frame, text="", width=150, font=("Microsoft YaHei", 15, "bold"), command=self._start_task)
        self.start_btn.pack(side="right", padx=10)

        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=20, pady=20)
        self.percent_label = ctk.CTkLabel(self.progress_frame, text="0%", font=("Arial", 56, "bold"), text_color="#2ECC71")
        self.percent_label.pack(pady=(5, 0))
        self.status_label = ctk.CTkLabel(self.progress_frame, text="", font=("Microsoft YaHei", 14), text_color="gray")
        self.status_label.pack(pady=5)

    # ================== 状态栏辅助 ==================
    def _set_status(self, key, text_color="gray", **fmt_args):
        """统一入口：记录 key + args，便于切语言时重新渲染。"""
        self._status_key = key
        self._status_args = fmt_args
        self._status_color = text_color
        try:
            self.status_label.configure(text=t(key, **fmt_args), text_color=text_color)
        except tk.TclError:
            pass

    def _refresh_status(self):
        """切语言后按记录的 key 重新生成当前状态栏文案。"""
        if not self._status_key:
            return
        try:
            self.status_label.configure(
                text=t(self._status_key, **self._status_args),
                text_color=self._status_color,
            )
        except tk.TclError:
            pass

    # ================== 语言核心逻辑 ==================
    def _change_language(self, new_lang):
        if new_lang == self._lang:
            return
        self._lang = new_lang
        set_lang(new_lang)
        save_lang_to_file(new_lang)
        self._apply_language()

    def _apply_language(self):
        """
        刷新静态文案 + 按当前 AppState 刷新按钮 text。
        注意：
          - drag_btn 的 state 由任务生命周期（_start_task / _reset_task_state）管理，这里只改 text
          - start_btn 在 PROCESSING 下需区分"取消转换"与"取消中..."（后者由 _cancel_task 置 disabled）
          - status_label 由 _refresh_status() 基于 _status_key 重新翻译，不覆盖终态信息
        """
        self.title(t("window_title"))
        self.title_label.configure(text=t("app_title"))
        self.browse_btn.configure(text=t("btn_browse"))

        if self._app_state == AppState.IDLE:
            self.start_btn.configure(text=t("btn_start"))
            self.drag_btn.configure(text=t("btn_drag"))
            if not self.file_list:
                # 区分"真的没选"与"写入失败"：写入失败时 file_list 也为空，但语义不同
                txt = t("msg_write_fail") if self._status_key == "msg_write_fail" else t("file_placeholder")
                self.file_textbox.configure(state="normal")
                self.file_textbox.delete("0.0", "end")
                self.file_textbox.insert("0.0", txt)
                self.file_textbox.configure(state="disabled")

        elif self._app_state == AppState.PROCESSING:
            # _cancel_task 会把 start_btn 置为 disabled 且不再复位；
            # 用 state 判别"已点取消、等待线程响应"状态，避免把"取消中..."打回"取消转换"
            if self.start_btn.cget("state") == "disabled":
                self.start_btn.configure(text=t("btn_canceling"))
            else:
                self.start_btn.configure(text=t("btn_cancel"))
            self.drag_btn.configure(text=t("btn_drag"))

        elif self._app_state == AppState.DRAGGING:
            self.start_btn.configure(text=t("btn_start"))
            self.drag_btn.configure(text=t("status_dragging"))

        self._refresh_status()

    # ================== 业务逻辑 ==================
    def _safe_after(self, fn, *args, **kwargs):
        if self._closing:
            return
        def wrapper():
            if self._closing:
                return
            try:
                fn(*args, **kwargs)
            except tk.TclError:
                pass
        try:
            self.after(0, wrapper)
        except tk.TclError:
            pass

    def _is_valid_file(self, path: str) -> bool:
        if not os.path.isfile(path):
            return False
        valid_exts = ('.pdf', '.png', '.jpg', '.jpeg')
        return path.lower().endswith(valid_exts)

    def _set_file_path(self, path: str, status_key: str = "status_file_selected"):
        try:
            self.file_textbox.configure(state="normal")
            self.file_textbox.delete("0.0", "end")
            self.file_textbox.insert("0.0", path)
            self.file_list = [path]
        except Exception as e:
            log.error("更新文件路径失败", exc_info=True)
            self.file_list = []
            try:
                self.file_textbox.delete("0.0", "end")
                self.file_textbox.insert("0.0", t("msg_write_fail"))
            except Exception:
                pass
            self._set_status("msg_write_fail", "#E74C3C")
            return
        finally:
            self.file_textbox.configure(state="disabled")
        self._set_status(status_key, "#2ECC71")

    def _select_file(self):
        if self._app_state != AppState.IDLE:
            return
        path = cross_platform_select_file(self)
        if path:
            if not self._is_valid_file(path):
                messagebox.showwarning(t("msg_format_error"), t("msg_select_valid_file"))
                return
            self._set_file_path(path)

    def _open_drag_window(self):
        if self._app_state != AppState.IDLE:
            return
        if not STARTING_DRAG:
            messagebox.showinfo(t("msg_tip"), t("msg_drag_not_ready"))
            return

        base_dir = get_base_dir()
        drag_name = "drag_en.exe" if self._lang == "en" else "drag.exe"
        drag_exe = os.path.join(base_dir, drag_name)

        if not os.path.exists(drag_exe):
            drag_exe = os.path.join(base_dir, "drag.exe")

        if not os.path.exists(drag_exe):
            messagebox.showerror(t("msg_error"), t("msg_drag_missing", path=drag_exe))
            return

        self._app_state = AppState.DRAGGING
        self.drag_btn.configure(state="disabled", text=t("status_dragging"))
        self.start_btn.configure(state="disabled")
        self._set_status("status_drag_prompt", "gray")

        def run_drag_thread():
            txt_path = os.path.join(base_dir, "drop_path.txt")
            try:
                if os.path.exists(txt_path):
                    os.remove(txt_path)
            except OSError:
                pass

            proc = None
            try:
                flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
                proc = subprocess.Popen(
                    [drag_exe],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    text=True, encoding="utf-8", errors="replace",
                    creationflags=flags,
                )
                self._drag_proc = proc

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

                stdout, stderr = proc.communicate(timeout=300)
                returncode = proc.returncode

                if stderr.strip():
                    if returncode != 0:
                        log.warning("drag.exe 非零退出 (code=%s): %s", returncode, stderr.strip())
                    else:
                        log.info("drag.exe 正常退出，stderr: %s", stderr.strip())

                dropped_path = stdout.strip().lstrip("\ufeff\u200b\u200c\u200d")
                if "\n" in dropped_path:
                    log.warning("drag.exe 输出多个路径，只取第一个: %r", dropped_path)
                    dropped_path = dropped_path.splitlines()[0].strip()
                if "\ufffd" in dropped_path:
                    log.warning("编码异常: %r", dropped_path)

                if not dropped_path and os.path.exists(txt_path):
                    try:
                        with open(txt_path, "r", encoding="utf-8") as f:
                            dropped_path = f.read().strip().lstrip("\ufeff\u200b\u200c\u200d")
                    except OSError as e:
                        log.warning("读取 drop_path.txt 失败: %s", e)

                if dropped_path and os.path.exists(txt_path):
                    try:
                        os.remove(txt_path)
                    except OSError:
                        pass

                if dropped_path and self._is_valid_file(dropped_path):
                    self._safe_after(self._set_file_path, dropped_path, "drag_success")
                elif not dropped_path and returncode != 0:
                    self._safe_after(self._set_status, "drag_error_exit", "#E74C3C")
                else:
                    if dropped_path and len(dropped_path) > 260:
                        log.warning("路径超过 260 字符，可能被 Windows API 拒绝: %r", dropped_path)
                    self._safe_after(self._set_status, "drag_invalid", "gray")

            except subprocess.TimeoutExpired:
                if proc is not None:
                    try:
                        proc.kill()
                        proc.wait(timeout=0.5)
                    except Exception:
                        pass
                self._safe_after(self._set_status, "drag_timeout", "#E74C3C")
            except Exception as e:
                log.error("拉起拖拽失败: %s", e)
                self._safe_after(self._set_status, "drag_failed", "#E74C3C")
            finally:
                self._drag_proc = None
                self._safe_after(self._on_drag_done)

        try:
            threading.Thread(target=run_drag_thread, daemon=True).start()
        except RuntimeError:
            log.error("无法启动拖拽线程", exc_info=True)
            self._app_state = AppState.IDLE
            self.drag_btn.configure(state="normal", text=t("btn_drag"))
            self.start_btn.configure(
                state="normal", text=t("btn_start"), command=self._start_task,
                fg_color=self._start_btn_default_fg, hover_color=self._start_btn_default_hover,
            )
            messagebox.showerror(t("msg_error"), t("msg_resource_low"))

    def _on_drag_done(self):
        self._app_state = AppState.IDLE
        self.drag_btn.configure(state="normal", text=t("btn_drag"))
        self.start_btn.configure(
            state="normal", text=t("btn_start"), command=self._start_task,
            fg_color=self._start_btn_default_fg, hover_color=self._start_btn_default_hover,
        )

    def _start_task(self):
        if self._app_state != AppState.IDLE:
            return

        if not self.file_list:
            messagebox.showwarning(t("msg_tip"), t("msg_no_file"))
            return
        if not self._is_valid_file(self.file_list[0]):
            messagebox.showerror(t("msg_error"), t("msg_invalid_file"))
            return

        self._app_state = AppState.PROCESSING
        self._cancel_event.clear()
        self._engine_loading = False

        self.start_btn.configure(state="normal", text=t("btn_cancel"), command=self._cancel_task, fg_color="#E74C3C", hover_color="#C0392B")
        # 转换期间禁用拖拽按钮（防并发）；恢复放到 _reset_task_state，避免与"切语言"耦合
        self.drag_btn.configure(state="disabled", text=t("btn_drag"))
        self.percent_label.configure(text="0%", text_color="#F1C40F")
        self._set_status("status_preparing", "gray")

        try:
            threading.Thread(target=self._run_conversion_thread, daemon=True).start()
        except RuntimeError:
            log.error("无法启动转换线程", exc_info=True)
            self._reset_task_state("failed", err=RuntimeError(t("msg_conv_thread_fail")))

    def _cancel_task(self):
        if self._closing:
            return
        self._cancel_event.set()
        self.start_btn.configure(state="disabled", text=t("btn_canceling"))

        if self._engine_loading:
            self._set_status("status_engine_loading", "#F39C12")
        else:
            self._set_status("status_canceling", "gray")

    def _run_conversion_thread(self):
        try:
            from ocr_engine import get_engine, ConversionCancelled, is_engine_loaded
        except Exception as e:
            log.error("OCR 模块导入失败", exc_info=True)
            self._safe_after(self._reset_task_state, "failed",
                             err=Exception(t("msg_ocr_import_fail", err=e)))
            return

        try:
            if self._cancel_event.is_set():
                raise ConversionCancelled("启动前已取消")

            need_load = not is_engine_loaded()
            if need_load:
                self._engine_loading = True
                self._safe_after(self._set_status, "status_loading", "gray")
            else:
                self._safe_after(self._set_status, "status_starting", "gray")

            engine = get_engine()

            if self._cancel_event.is_set():
                raise ConversionCancelled("初始化后取消")

            input_path = self.file_list[0]
            output_dir = os.path.dirname(input_path)

            first_progress = [True]
            def progress_cb(current, total):
                if first_progress[0]:
                    first_progress[0] = False
                    self._engine_loading = False
                percent = int(current / total * 100) if total > 0 else 0
                self._safe_after(self.percent_label.configure, text=f"{percent}%")
                self._safe_after(self._set_status, "status_recognizing", "gray",
                                 cur=current, total=total)

            if input_path.lower().endswith(".pdf"):
                out_path = engine.convert_pdf(input_path, output_dir, progress_cb, self._cancel_event)
            else:
                if not hasattr(engine, "convert_image"):
                    raise NotImplementedError(t("msg_img_not_ready"))
                out_path = engine.convert_image(input_path, output_dir, progress_cb, self._cancel_event)

            if not out_path:
                raise RuntimeError(t("msg_no_output"))

            self._safe_after(self._reset_task_state, "success", path=out_path)

        except ConversionCancelled:
            self._safe_after(self._reset_task_state, "cancelled")
        except NotImplementedError as e:
            log.info("功能未实现: %s", e)
            self._safe_after(self._reset_task_state, "not_implemented", msg=str(e))
        except Exception as e:
            if self._cancel_event.is_set():
                self._safe_after(self._reset_task_state, "cancelled")
            else:
                log.error("转换异常", exc_info=True)
                self._safe_after(self._reset_task_state, "failed", err=e)

    def _reset_task_state(self, outcome, *, path=None, err=None, msg=None):
        """
        outcome:
          - "success"         : path 为输出文件路径
          - "cancelled"       : 用户主动取消
          - "failed"          : err 为异常对象（None 时退化为通用失败文案）
          - "not_implemented" : msg 为已翻译的说明文本（不可再随语言刷新）
        """
        if self._closing:
            return

        self._app_state = AppState.IDLE
        self._engine_loading = False
        self.start_btn.configure(
            state="normal", text=t("btn_start"), command=self._start_task,
            fg_color=self._start_btn_default_fg, hover_color=self._start_btn_default_hover,
        )
        # 关键：恢复拖拽按钮状态（原版本遗漏，导致转换中切语言后拖拽按钮永久灰色）
        self.drag_btn.configure(state="normal", text=t("btn_drag"))

        if outcome == "success":
            self.percent_label.configure(text="100%", text_color="#2ECC71")
            self._set_status("msg_done", "#2ECC71")
        elif outcome == "cancelled":
            self.percent_label.configure(text=t("status_cancelled"), text_color="gray")
            self._set_status("status_cancelled", "gray")
        elif outcome == "not_implemented":
            self.percent_label.configure(text=t("msg_tip"), text_color="gray")
            # msg 已是翻译过的字符串，切语言无法重新生成，保持原样
            self._status_key = None
            self._status_args = {}
            self._status_color = "gray"
            try:
                self.status_label.configure(text=msg or "", text_color="gray")
            except tk.TclError:
                pass
        else:  # failed
            self.percent_label.configure(text=t("msg_failed"), text_color="#E74C3C")
            if err is not None:
                self._set_status("msg_conv_fail_body", "#E74C3C", err=err)
            else:
                self._set_status("msg_failed", "#E74C3C")

        if self._confirming_close:
            return

        if outcome == "success":
            messagebox.showinfo(t("msg_success_title"), t("msg_success_body", path=path))
        elif outcome == "failed":
            fail_msg = t("msg_conv_fail_body", err=err) if err is not None else t("msg_failed")
            messagebox.showerror(t("msg_fail_title"), t("msg_fail_body", msg=fail_msg))

    def _on_close(self):
        if self._confirming_close:
            return
        self._confirming_close = True

        try:
            if self._app_state != AppState.IDLE:
                try:
                    confirmed = messagebox.askyesno(t("msg_confirm_title"), t("msg_confirm_exit"))
                except Exception as e:
                    log.error("关闭确认框异常", exc_info=True)
                    confirmed = False

                if not confirmed:
                    return

            self._closing = True
        finally:
            self._confirming_close = False

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

        self._cancel_event.set()
        self.destroy()


if __name__ == "__main__":
    setup_crash_handler()
    app = PDF2WordApp()
    app.mainloop()