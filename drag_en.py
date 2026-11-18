import sys
import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES

# 【契约】VALID_EXTS 必须与 gui.py 的 _is_valid_file / cross_platform_select_file 保持同步。
VALID_EXTS = ('.pdf', '.png', '.jpg', '.jpeg')

# 错误提示停留时长（毫秒）
ERROR_RESET_MS = 2000
# 空闲超时自动关闭（毫秒）
IDLE_TIMEOUT_MS = 60000
# 超时前多久开始警告（毫秒）——即 60s 中的最后 20s
IDLE_WARNING_MS = 20000

# ================== 平台检测 ==================
CURRENT_PLATFORM = "Windows"
if sys.platform.startswith("linux"):
    CURRENT_PLATFORM = "Linux"
elif sys.platform == "darwin":
    CURRENT_PLATFORM = "macOS"

def _get_exe_dir() -> str:
    """获取 drag.exe 真实所在目录（与 utils.get_base_dir 五级兜底对齐）。"""
    env_dir = os.environ.get("PDF2WORD_BASE_DIR")
    if env_dir:
        return os.path.abspath(env_dir)

    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)

    mod = sys.modules.get(__name__)
    file_path = getattr(mod, "__file__", None) if mod else None
    if file_path and not file_path.startswith("<"):
        return os.path.dirname(os.path.abspath(file_path))

    if sys.argv and sys.argv[0]:
        return os.path.dirname(os.path.abspath(sys.argv[0]))

    return os.getcwd()


def _safe_write_stream(stream, text: str):
    """安全写入流（stdout/stderr 通用），--noconsole 下 stream 为 None 时跳过。"""
    if stream is None:
        return
    try:
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
        stream.write(text + "\n")
        stream.flush()
    except Exception:
        pass


def _safe_write_txt(txt_path: str, text: str):
    """txt 兜底通道，独立 try，保证一定写成功。"""
    try:
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass


def main():
    try:
        root = TkinterDnD.Tk()
    except Exception as e:
        if CURRENT_PLATFORM == "Linux":
            platform_tip = "Please install system dependency 'tkdnd'."
        elif CURRENT_PLATFORM == "macOS":
            platform_tip = "Please check your Tkinter environment."
        else:
            platform_tip = "Please check if tkinterdnd2 is installed correctly."
            
        _safe_write_stream(
            sys.stderr,
            f"\n tkinterdnd2 Init Failed】\n"
            f"Error: {e}\n"
            f"Tip: {platform_tip}\n"
        )
        sys.exit(1)

    root.title("Drag File")
    root.geometry("350x200")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    try:
        root.eval('tk::PlaceWindow . center')
    except Exception:
        pass

    label = tk.Label(
        root,
        text="Drag PDF or Image here\n(Release to drop)",
        font=("Microsoft YaHei", 14),
        bg="#f0f0f0",
        fg="#333333"
    )
    label.pack(expand=True, fill="both", padx=20, pady=20)

    exe_dir = _get_exe_dir()
    txt_path = os.path.join(exe_dir, "drop_path.txt")

    _reset_after_id = None
    _warning_after_id = None
    _timeout_after_id = None

    def _cancel_all_timers():
        nonlocal _reset_after_id, _warning_after_id, _timeout_after_id
        for after_id in (_reset_after_id, _warning_after_id, _timeout_after_id):
            if after_id is not None:
                try:
                    root.after_cancel(after_id)
                except Exception:
                    pass
        _reset_after_id = None
        _warning_after_id = None
        _timeout_after_id = None

    def _reset_to_idle_ui():
        try:
            label.configure(
                text="Drag PDF or Image here\n(Release to drop)",
                fg="#333333"
            )
        except Exception:
            pass

    def _on_idle_timeout():
        try:
            root.destroy()
        except Exception:
            pass

    def _show_idle_warning():
        """40 秒时显示警告文案，再等 20 秒后关闭。"""
        nonlocal _warning_after_id, _timeout_after_id
        _warning_after_id = None
        # 【P4 修复】label.configure 与 root.after 拆成两个 try
        # 即使 label 配置失败（窗口已销毁），也确保 timeout 逻辑独立执行
        try:
            label.configure(
                text="Closing soon...\nPlease drop a file (Click window to reset timer).",
                fg="#F39C12"
            )
        except Exception:
            pass
        try:
            _timeout_after_id = root.after(IDLE_WARNING_MS, _on_idle_timeout)
        except Exception:
            _timeout_after_id = None

    def _start_idle_timer():
        """重置空闲计时器：60 秒总长，40 秒时触发警告。"""
        nonlocal _warning_after_id
        _cancel_all_timers()
        _reset_to_idle_ui()
        try:
            warning_delay = IDLE_TIMEOUT_MS - IDLE_WARNING_MS
            _warning_after_id = root.after(warning_delay, _show_idle_warning)
        except Exception:
            pass

    def _reset_to_idle():
        """错误提示 2 秒后恢复默认 UI。"""
        nonlocal _reset_after_id
        _reset_after_id = None
        # 【P4 加注释】_warning_after_id 为 None 说明警告已触发（进入最后 20 秒），
        # 此时不应覆盖警告文案，所以只在计时器仍活跃时恢复默认 UI
        if _warning_after_id is not None:
            _reset_to_idle_ui()

    def _show_error(tip: str):
        """显示错误提示，2 秒后恢复默认 UI。不负责重置空闲计时。"""
        nonlocal _reset_after_id
        # 【保留的防御】不依赖调用方清理 _reset_after_id：
        # 未来若从其他入口（如定时检测）触发 _show_error，这里显式 cancel 旧恢复任务
        if _reset_after_id is not None:
            try:
                root.after_cancel(_reset_after_id)
            except Exception:
                pass
        _reset_after_id = None

        try:
            label.configure(text=tip, fg="#E74C3C")
            _reset_after_id = root.after(ERROR_RESET_MS, _reset_to_idle)
        except Exception:
            _reset_after_id = None

    def on_drop(event):
        # 任何拖拽动作统一在此重置空闲计时
        _start_idle_timer()

        try:
            files = root.tk.splitlist(event.data)
        except Exception:
            files = [event.data]

        if not files:
            root.destroy()
            return

        valid_file = None
        for f in files:
            f_clean = f.strip().lstrip("\ufeff\u200b\u200c\u200d")
            if f_clean.lower().endswith(VALID_EXTS):
                valid_file = f_clean
                break

        if valid_file is None:
            tip = "Unsupported file format\nOnly support PDF / PNG / JPG / JPEG"
            if len(files) > 1:
                tip += f"\n({len(files)} files, none supported)"
            _show_error(tip)
            return

        _cancel_all_timers()
        _safe_write_txt(txt_path, valid_file)
        _safe_write_stream(sys.stdout, valid_file)
        root.destroy()

    def _bind_drop(widget):
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind('<<Drop>>', on_drop)

    _bind_drop(root)
    _bind_drop(label)

    def _on_user_interact(event=None):
        """用户点击 / 鼠标进入窗口都重置空闲计时，符合"点击可重置"的直觉。"""
        _start_idle_timer()

    # 【P4 修复】<Button-1> 覆盖点击，<Enter> 覆盖鼠标进入窗口
    # 两者叠加 = "只要用户在操作窗口，就不会被超时关闭"
    root.bind("<Button-1>", _on_user_interact)
    label.bind("<Button-1>", _on_user_interact)
    root.bind("<Enter>", _on_user_interact)
    label.bind("<Enter>", _on_user_interact)

    root.bind("<Escape>", lambda e: root.destroy())
    root.protocol("WM_DELETE_WINDOW", root.destroy)

    _start_idle_timer()
    root.mainloop()


if __name__ == "__main__":
    main()