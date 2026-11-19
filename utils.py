import sys
import os
import time
import logging
import tempfile
import threading
import subprocess
import shutil
import atexit
from functools import lru_cache
from logging.handlers import RotatingFileHandler


class ModelDirError(RuntimeError):
    """自定义异常：用于精确表示 OCR 模型目录缺失或配置错误的情况"""
    pass


# ================= 全局状态与锁 =================
_crashed = False          # 标记：本次运行是否发生过任意崩溃
_fatal_crash = False      # 标记：是否发生致命崩溃（主线程崩溃，进程即将退出）
_logger_ready = False     # 标记：日志系统是否已完整初始化，防止并发重复挂载
_setup_lock = threading.Lock()  # 锁：保护日志系统的初始化过程


@lru_cache(maxsize=1)
def get_base_dir() -> str:
    """
    获取程序运行时的根目录（五级火箭兜底策略）。
    极其重要：决定了 models 文件夹和 logs 文件夹到底去哪里找。
    """
    # 1. 优先使用环境变量（方便测试或在特殊部署环境下强制指定）
    env_dir = os.environ.get("PDF2WORD_BASE_DIR")
    if env_dir:
        return os.path.abspath(env_dir)

    # 2. 打包成 exe 后（PyInstaller），必须以 exe 所在目录为准
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)

    # 3. 源码运行时，安全地获取当前模块的 __file__（防 python -c 时的 NameError）
    mod = sys.modules.get(__name__)
    file_path = getattr(mod, '__file__', None) if mod else None
    
    if file_path and not file_path.startswith("<"):
        return os.path.dirname(os.path.abspath(file_path))

    # 4. 退而求其次，找入口脚本所在的目录
    if sys.argv and sys.argv[0]:
        return os.path.dirname(os.path.abspath(sys.argv[0]))

    # 5. 极端环境兜底（如某些嵌入式解释器、动态 exec 伪造模块），只能用当前工作目录
    if sys.__stderr__:
        try:
            sys.__stderr__.write("[WARNING] 无法获取程序路径，已退回当前工作目录！\n")
        except Exception:
            pass
    return os.getcwd()


# 【关键操作】在模块导入期强制求值一次。
# 这样做的目的是把"当前工作目录（cwd）"锁死在内存里。
# 如果后续代码（或第三方库）调用了 os.chdir() 切换了目录，我们的 paths 依然是对的。
_ = get_base_dir()


def get_model_dir() -> str:
    """获取 PaddleOCR 模型所在的 models 文件夹路径"""
    return os.path.join(get_base_dir(), "models")


@lru_cache(maxsize=1)
def get_log_dir() -> str:
    """
    获取日志目录。带有降级策略：
    优先放在程序根目录下；如果根目录没权限（比如装在 C:\Program Files），降级到系统 Temp 目录。
    """
    log_dir = os.path.join(get_base_dir(), "logs")
    try:
        os.makedirs(log_dir, exist_ok=True)
        # 用临时文件测试一下是否有写入权限（比直接写日志测试更安全）
        with tempfile.NamedTemporaryFile(dir=log_dir, delete=True) as f:
            pass
    except OSError:
        # 降级到 Temp 目录。如果 Temp 也写不了，那就让它抛 OSError 早崩早发现，拒绝静默失败。
        log_dir = os.path.join(tempfile.gettempdir(), "PDF2Word_Logs")
        os.makedirs(log_dir, exist_ok=True)  
    return log_dir


def setup_logger() -> logging.Logger:
    """
    初始化日志系统。采用双检锁（Double-Checked Locking）防止并发重复初始化。
    """
    global _logger_ready
    logger = logging.getLogger("PDF2Word")
    
    # 第一重检查：如果已经初始化好了，直接返回
    if _logger_ready:
        return logger

    # 加锁：防止多线程同时进来导致重复挂载 Handler（会导致同一条日志打印两遍）
    with _setup_lock:
        # 第二重检查：确保拿到锁后，别的线程没有刚刚完成初始化
        if _logger_ready:
            return logger
        # 如果 logger 已经有 handler 了（比如被外部代码污染过），直接视为就绪
        if logger.handlers:
            _logger_ready = True
            return logger

        logger.setLevel(logging.INFO)
        # 阻止日志向上传递给 root logger，防止被其他库的 basicConfig 重复打印
        logger.propagate = False

        # 包含模块名和行号的格式，方便打包无控制台时精确定位报错
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s'
        )

        console_handler = None
        if sys.stdout is not None:  # 如果打包成 --noconsole，sys.stdout 是 None，必须跳过
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(formatter)

        log_file = os.path.join(get_log_dir(), "app.log")
        
        # 【关键操作】先构造 file_handler。如果磁盘满了抛异常，此时还没挂载，不会污染 logger.handlers
        file_handler = RotatingFileHandler(
            log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
        )
        file_handler.setFormatter(formatter)

        # 构造成功后，再统一挂载
        if console_handler is not None:
            logger.addHandler(console_handler)
        logger.addHandler(file_handler)

        logger.info("日志系统初始化完成。当前日志目录: %s", os.path.dirname(log_file))
        logger.info("Logging system initialized. Current log directory: %s", os.path.dirname(log_file))

        # 注册程序退出时的清理动作
        atexit.register(_log_normal_exit)
        _logger_ready = True
        return logger


def _log_normal_exit():
    """
    通过 atexit 在程序退出时记录最后一条状态。
    严格区分：正常退出 / 非致命崩溃后退出 / 致命崩溃退出，方便事后排查。
    """
    try:
        if _fatal_crash:
            # 致命崩溃时，_report_crash 已经写了 CRITICAL 日志，这里补一条结论
            log.warning("=== 进程因致命崩溃退出（详见上方 CRITICAL） ===")
            for h in log.handlers:
                h.flush()
            return
        
        if _crashed:
            # 子线程/UI回调崩溃了，但主程序活下来了，最后正常关窗退出
            log.warning("程序退出（本次运行期间发生过非致命崩溃，详见上文 CRITICAL）")
        else:
            # 全程无崩溃，完美退出
            log.info("程序正常退出")
            
        # 强制落盘，防止断电/强杀导致最后一条日志丢失
        for h in log.handlers:
            h.flush()
    except Exception:
        pass


# 实例化全局 logger 供外部导入（如 `from utils import log`）
log = setup_logger()


def ensure_model_dir() -> str:
    """
    模型根目录预检：只确保 models/ 存在。

    具体模型是否齐全（det / rec_ch / layout 等）由 ocr_engine._ensure_model
    负责——它会在模型缺失时自动下载，并给出精确的中文错误提示。

    本函数只做路径级检查，不做模型级检查：
    模型目录结构会随 fast_model/pro_model 分档、模型版本升级而变化，
    硬编码具体子目录必然过时。
    """
    model_dir = get_model_dir()
    if not os.path.isdir(model_dir):
        log.error("模型目录不存在: %s", model_dir)
        raise ModelDirError(
            f"模型目录不存在: {model_dir}\n"
            f"(base_dir = {get_base_dir()}；如需自定义请设置 PDF2WORD_BASE_DIR)\n"
            f"（首次运行会自动创建并下载模型）"
        )
    return model_dir


# =====================================================
# 全局崩溃拦截模块（三路钩子 + 并发安全）
# =====================================================
# 保存原始的 excepthook，用于链式调用（防覆盖）
_original_excepthook = sys.excepthook
_original_thread_excepthook = threading.excepthook

_crash_popup_opened = False         # 标记：是否已经弹过记事本了（防止弹窗风暴）
_crash_lock = threading.RLock()     # 锁：使用可重入锁(RLock)，防止未来嵌套调用导致死锁
_crash_handler_installed = False    # 标记：崩溃钩子是否已安装（保证幂等）


def get_log_file_path() -> str:
    """获取当前日志文件的绝对路径"""
    return os.path.join(get_log_dir(), "app.log")


def _open_log_in_notepad():
    """
    在致命崩溃时，自动用记事本打开日志文件。
    【核心亮点】：使用的是日志的"快照（副本）"，而不是原文件。
    因为 Windows 下用记事本打开原文件会加锁，导致正在运行的程序无法轮转日志，从而静默丢失数据。
    """
    global _crash_popup_opened
    
    if not sys.platform.startswith("win"):
        # 非 Windows 平台兜底：不弹窗，往标准错误流输出一行提示
        try:
            if sys.__stderr__:
                sys.__stderr__.write(f"致命崩溃，日志已写入: {get_log_file_path()}\n")
        except Exception:
            pass
        return
        
    # 进锁占位，防止 PaddleOCR 多个子线程同时崩溃导致弹出多个记事本
    with _crash_lock:
        if _crash_popup_opened:
            return
        _crash_popup_opened = True

    try:
        src_log = get_log_file_path()
        # 生成带 pid 和时间戳的快照路径，防止旧的快照没关导致新快照复制失败
        dst_snapshot = os.path.join(
            tempfile.gettempdir(), 
            f"PDF2Word_crash_snapshot_{os.getpid()}_{int(time.time())}.log"
        )
        
        try:
            shutil.copy2(src_log, dst_snapshot)
            target = dst_snapshot
        except Exception:
            # 如果复制快照失败（极端情况），退而求其次直接打开原文件
            target = src_log

        notepad = shutil.which("notepad.exe")
        if not notepad:
            os.startfile(target)
        else:
            # 加上 CREATE_NO_WINDOW 标志，防止打包成 --noconsole 时闪黑框
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.Popen(
                [notepad, target],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=flags,
            )
    except Exception:
        # 连记事本都打不开，回滚标记，并做最后的手动日志输出
        with _crash_lock:
            _crash_popup_opened = False
        try:
            log.error("无法自动打开日志文件: %s", get_log_file_path(), exc_info=True)
            if sys.__stderr__:
                sys.__stderr__.write(f"崩溃日志已写入: {get_log_file_path()}\n")
        except Exception:
            pass


def _report_crash(exc_type, exc_value, exc_tb, source: str = "主线程", fatal: bool = False):
    """统一的崩溃上报入口。"""
    global _crashed, _fatal_crash
    
    with _crash_lock:
        _crashed = True
        if fatal:
            _fatal_crash = True
        
    try:
        # 【核心修复】彻底参数化，绝不让路径中的 % 拼接进格式串，防止 logging 格式化异常导致日志丢失
        if fatal:
            log.critical(
                "%s 发生未捕获异常",
                source,
                exc_info=(exc_type, exc_value, exc_tb),
            )
        else:
            log.critical(
                "%s 发生非致命未捕获异常（日志路径: %s）",
                source, get_log_file_path(),
                exc_info=(exc_type, exc_value, exc_tb),
            )
            
        # 崩溃时必须强制刷盘，防止程序瞬间退出导致日志没写进去
        for h in log.handlers:
            try:
                h.flush()
            except Exception:
                pass
    except Exception:
        pass

    # 只有致命崩溃才弹记事本打扰用户；子线程崩溃不打断主程序运行
    if fatal:
        _open_log_in_notepad()


def setup_crash_handler():
    """安装全局崩溃拦截钩子。"""
    global _crash_handler_installed, _original_excepthook, _original_thread_excepthook
    
    # 【核心修复】将捕获上游 hook、定义闭包、赋值 sys.excepthook 完全放在锁内，消除并发竞态
    with _crash_lock:
        if _crash_handler_installed:
            return
        _crash_handler_installed = True
        
        # 重新捕获"上一级"的 hook，防止模块加载后、setup 调用前，被别的框架（如 PyQt）悄悄覆盖
        _original_excepthook = sys.excepthook
        _original_thread_excepthook = threading.excepthook

        def handle_exception(exc_type, exc_value, exc_tb):
            if issubclass(exc_type, (KeyboardInterrupt, SystemExit)):
                _original_excepthook(exc_type, exc_value, exc_tb)
                return
            _report_crash(exc_type, exc_value, exc_tb, source="主线程", fatal=True)
            # 链式调用原有的 hook，确保不破坏第三方库的原有逻辑
            _original_excepthook(exc_type, exc_value, exc_tb)

        sys.excepthook = handle_exception

        def handle_thread_exception(args):
            if issubclass(args.exc_type, (KeyboardInterrupt, SystemExit)):
                _original_thread_excepthook(args)
                return
            name = args.thread.name if args.thread else "unknown"
            _report_crash(args.exc_type, args.exc_value, args.exc_traceback, source=f"子线程[{name}]", fatal=False)
            _original_thread_excepthook(args)

        threading.excepthook = handle_thread_exception


def tk_exception_hook(exc_type, exc_value, exc_tb):
    """专门用于接管 Tkinter 按钮回调等 GUI 层异常的钩子"""
    if issubclass(exc_type, (KeyboardInterrupt, SystemExit)):
        return
    _report_crash(exc_type, exc_value, exc_tb, source="Tkinter GUI 回调", fatal=False)
