import sys
import os

# ==============================================================================
# 【极其关键】OpenMP 冲突防御
# ==============================================================================
# 必须在导入任何 paddle 或 gui 模块之前设置此环境变量！
# 原因：Paddle 底层依赖 MKL，加载 libiomp5md.dll 时会与系统/其他库的 OpenMP 冲突，
# 直接导致 OSError #15 甚至无提示闪退。此变量允许多个 OpenMP 运行时共存。
# 注意：ocr_engine.py 顶部也设了同样的变量，这是刻意双保险。
# 不要删 ocr_engine.py 里那一处，因为 ocr_engine 未来可能被其他入口（如 pytest）单独导入。
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

# 提前导入基建模块：获取日志实例和崩溃钩子安装器
from utils import setup_crash_handler, log, get_base_dir


def main():
    # ==============================================================================
    # 【第一步】挂载全局崩溃拦截钩子
    # ==============================================================================
    # 覆盖三路异常：主线程 (sys.excepthook)、子线程 (threading.excepthook)、
    # Tkinter 回调异常 (tk_exception_hook，需在 GUI 里单独赋值)。
    # 这样后续 GUI 加载失败（比如 customtkinter 没装）也能被捕获，而不是直接闪退。
    setup_crash_handler()
    
    # 记录启动上下文（带根目录，方便打包后排查“到底在哪找 models”）
    # 使用参数化 %s 而不是 f-string：防止 Windows 路径中自带的 % 触发格式化异常
    log.info("正在启动 PDF2Word 主程序... | 根目录: %s", get_base_dir())

    try:
        # ==============================================================================
        # 【第二步】延迟导入 GUI
        # ==============================================================================
        # 放在 try 内部：如果 gui.py 有语法错误或依赖缺失（如 customtkinter 未安装），
        # 异常能被下方的 except 捕获，避免在模块加载期直接死掉。
        from gui import PDF2WordApp
        
        # 实例化主窗口（此时 GUI 仅画界面，不会加载 Paddle 模型，启动很快）
        app = PDF2WordApp()
        
        # 进入 Tkinter 事件主循环，程序在此挂起，直到窗口关闭
        app.mainloop()
        
    except Exception:
        # ==============================================================================
        # 【第三步】启动失败处理（P2 修复核心）
        # ==============================================================================
        # 首先写一条带上下文（"主程序启动失败"）的 CRITICAL 日志。
        # exc_info=True 会自动把完整的 traceback 也带进日志文件。
        log.critical("主程序启动失败", exc_info=True)
        
        # 【关键】直接 raise，绝不使用 sys.exit(1)！
        # 原因：sys.exit(1) 会吞掉异常，导致 main() 内部结束，Python 解释器认为正常退出，
        # 从而【不会】触发 sys.excepthook。这会导致 --noconsole 打包下，用户只看到
        # 窗口闪一下消失，既没有控制台报错，也不会弹出记事本快照。
        #
        # 用 raise 让异常逃逸出 main()，Python 解释器接手后会自动调用我们在
        # setup_crash_handler() 里安装的 handle_exception：
        # 1. 触发 _report_crash(fatal=True) → 弹记事本显示日志快照！
        # 2. 写入第二条带完整 traceback 的 CRITICAL 日志。
        # 3. 退出码依然是 1。
        raise 

if __name__ == "__main__":
    main()