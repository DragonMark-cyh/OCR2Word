import os

# ================== 核心开关 ==================
# 平板上调试设为 True，回到 PC 端打包发版时改为 False
ENABLE_TEST = True

# ================== 语言字典 ==================
STRINGS = {
    "zh": {
        # --- 主界面 ---
        "window_title": "通用文档 OCR 转换工具",
        "app_title": "通用文档 OCR 转换器",
        "file_placeholder": "尚未选择文件（支持 PDF / PNG / JPG / JPEG）",
        "btn_browse": "浏览...",
        "btn_drag": "独立弹窗拖拽",
        "btn_start": "开始转换",
        "status_ready": "就绪，请选择文件",
        
        # --- 文件选择器 ---
        "dialog_title": "选择需要识别的文档或图片",
        "filter_supported": "支持的文件",
        "filter_pdf": "PDF 文件",
        "filter_image": "图片文件",
        "filter_all": "所有文件",
        "msg_mac_picker_fail": "文件选择器不可用，请检查系统权限。",
        "msg_no_display": "当前环境缺少图形界面支持！",
        "msg_picker_crash": "无法打开文件选择器:\n{err}",

        # --- 文件路径状态 ---
        "status_file_selected": "已选择文件，点击开始转换",
        "msg_write_fail": "尚未选择文件（写入失败）",

        # --- 浏览与校验 ---
        "msg_format_error": "格式错误",
        "msg_tip": "提示",
        "msg_select_valid_file": "请选择有效的文档或图片文件！",
        "msg_no_file": "请先选择一个文件！",
        "msg_invalid_file": "当前路径不是支持的文件格式！",

        # --- 拖拽模块 ---
        "msg_drag_not_ready": "拖拽模块暂未接入！",
        "msg_drag_missing": "找不到拖拽组件: {path}",
        "status_dragging": "等待拖拽中...",
        "status_drag_prompt": "请在弹出的窗口中拖拽文件...",
        "drag_success": "拖拽成功！",
        "drag_invalid": "未获取到有效文件",
        "drag_error_exit": "拖拽工具异常退出，请查看日志",
        "drag_timeout": "拖拽超时",
        "drag_failed": "拖拽启动失败",
        "msg_resource_low": "系统资源不足，无法启动拖拽线程",

        # --- 转换状态机 ---
        "btn_cancel": "取消转换",
        "btn_canceling": "取消中...",
        "status_preparing": "正在准备 OCR 引擎...",
        "status_canceling": "正在结束当前操作，请稍候...",
        "status_engine_loading": "引擎加载中，取消将在初始化完成后生效...",
        "msg_conv_thread_fail": "系统资源不足，无法启动转换线程",

        # --- 核心转换线程 ---
        "msg_ocr_import_fail": "OCR 模块加载失败，请检查依赖是否安装：{err}",
        "msg_failed": "失败",
        "status_loading": "正在加载 OCR 模型（首次较慢，请稍候）...",
        "status_starting": "正在启动 OCR 引擎...",
        "status_recognizing": "正在识别：第 {cur}/{total} 页",
        "msg_img_not_ready": "图片转换功能暂未接入，请先使用 PDF 文件进行转换！",
        "msg_no_output": "转换引擎未返回输出路径（可能内部异常被静默吞掉）",
        "status_cancelled": "已取消",
        "msg_conv_fail_body": "转换失败: {err}",

        # --- 状态重置与弹窗 ---
        "msg_done": "转换完成！",
        "msg_success_title": "成功",
        "msg_success_body": "转换完成！\n\n输出文件：\n{path}",
        "msg_fail_title": "失败",
        "msg_fail_body": "处理过程中发生错误。\n{msg}",

        # --- 关闭窗口 ---
        "msg_confirm_title": "确认",
        "msg_confirm_exit": "任务进行中，确定要退出吗？",
        "msg_error": "错误"
    },
    
    "en": {
        # --- Main UI ---
        "window_title": "Universal Document OCR Converter",
        "app_title": "Universal Document OCR Converter",
        "file_placeholder": "No file selected (Supports PDF / PNG / JPG / JPEG)",
        "btn_browse": "Browse...",
        "btn_drag": "Drag Window",
        "btn_start": "Start",
        "status_ready": "Ready. Select a file.",
        
        # --- File Picker ---
        "dialog_title": "Select document or image",
        "filter_supported": "Supported Files",
        "filter_pdf": "PDF Files",
        "filter_image": "Image Files",
        "filter_all": "All Files",
        "msg_mac_picker_fail": "File picker unavailable. Check system permissions.",
        "msg_no_display": "No GUI support in current environment!",
        "msg_picker_crash": "Failed to open file picker:\n{err}",

        # --- File Path Status ---
        "status_file_selected": "File selected. Click Start.",
        "msg_write_fail": "No file selected (Write failed)",

        # --- Browse & Validate ---
        "msg_format_error": "Format Error",
        "msg_tip": "Notice",
        "msg_select_valid_file": "Please select a valid document or image file!",
        "msg_no_file": "Please select a file first!",
        "msg_invalid_file": "Unsupported file format!",

        # --- Drag Module ---
        "msg_drag_not_ready": "Drag module is not ready!",
        "msg_drag_missing": "Drag component not found: {path}",
        "status_dragging": "Waiting for drag...",
        "status_drag_prompt": "Drag file in the popup window...",
        "drag_success": "Drag success!",
        "drag_invalid": "No valid file received",
        "drag_error_exit": "Drag tool crashed. Check logs.",
        "drag_timeout": "Drag timeout",
        "drag_failed": "Failed to start drag",
        "msg_resource_low": "Insufficient resources to start drag thread",

        # --- State Machine ---
        "btn_cancel": "Cancel",
        "btn_canceling": "Canceling...",
        "status_preparing": "Preparing OCR engine...",
        "status_canceling": "Finishing current operation, please wait...",
        "status_engine_loading": "Engine loading. Cancel will take effect after init...",
        "msg_conv_thread_fail": "Insufficient resources to start conversion",

        # --- Conversion Thread ---
        "msg_ocr_import_fail": "OCR module load failed. Check dependencies: {err}",
        "msg_failed": "Failed",
        "status_loading": "Loading OCR model (First time may be slow)...",
        "status_starting": "Starting OCR engine...",
        "status_recognizing": "Recognizing: Page {cur}/{total}",
        "msg_img_not_ready": "Image conversion not ready. Please use PDF.",
        "msg_no_output": "Engine returned no output path (possible swallowed error)",
        "status_cancelled": "Cancelled",
        "msg_conv_fail_body": "Conversion failed: {err}",

        # --- Popups ---
        "msg_done": "Conversion complete!",
        "msg_success_title": "Success",
        "msg_success_body": "Conversion complete!\n\nOutput file:\n{path}",
        "msg_fail_title": "Failed",
        "msg_fail_body": "An error occurred during processing.\n{msg}",

        # --- Close Window ---
        "msg_confirm_title": "Confirm",
        "msg_confirm_exit": "Task is running, are you sure to exit?",
        "msg_error": "Error"
    }
}

# ================== 全局变量 ==================
_current_lang = "en"


# ================== 核心接口 ==================

def set_lang(lang: str):
    """纯内存操作：修改当前语言状态"""
    global _current_lang
    if lang in ("zh", "en"):
        _current_lang = lang

def t(key: str, **kwargs) -> str:
    """翻译字典：获取当前语言的文案，支持格式化变量"""
    template = STRINGS.get(_current_lang, {}).get(key, key)
    return template.format(**kwargs) if kwargs else template

def load_lang_from_file() -> str:
    """启动时读取：找 zh 或 en 标记文件。如果都没有，默认 en 并生成 en 文件。"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    zh_file = os.path.join(base_dir, "zh")
    en_file = os.path.join(base_dir, "en")

    if os.path.exists(zh_file):
        set_lang("zh")
        return "zh"
    if os.path.exists(en_file):
        set_lang("en")
        return "en"
    
    try:
        open(en_file, "w").close()
    except Exception:
        pass
    set_lang("en")
    return "en"

def save_lang_to_file(lang: str):
    """切换语言时保存：物理清理旧文件，精准创建新文件"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    zh_file = os.path.join(base_dir, "zh")
    en_file = os.path.join(base_dir, "en")
    new_file = os.path.join(base_dir, lang)

    try:
        if os.path.exists(zh_file): os.remove(zh_file)
        if os.path.exists(en_file): os.remove(en_file)
        open(new_file, "w").close()
    except Exception:
        pass


# ================== 防中断的纯逻辑测试 ==================
if __name__ == "__main__":
    if ENABLE_TEST:
        print("=== 开始测试 lang.py 核心逻辑 ===")
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        zh_file = os.path.join(base_dir, "zh")
        en_file = os.path.join(base_dir, "en")

        # 测试 1：首次默认加载
        try:
            lang = load_lang_from_file()
            if lang == "en" and os.path.exists(en_file):
                print("OK: 初始化逻辑正确，默认生成了 en 文件")
            else:
                print(f"FAIL: 初始化逻辑错误，返回值为 {lang}")
        except Exception as e:
            print(f"FAIL: 初始化测试崩溃 - {e}")

        # 测试 2：切换到中文
        try:
            save_lang_to_file("zh")
            lang = load_lang_from_file()
            if lang == "zh" and os.path.exists(zh_file) and not os.path.exists(en_file):
                print("OK: 切换到 zh 成功，en 文件已被清理")
            else:
                print(f"FAIL: 切换到 zh 失败，返回值为 {lang}")
        except Exception as e:
            print(f"FAIL: 切换 zh 测试崩溃 - {e}")

        # 测试 3：再切回英文
        try:
            save_lang_to_file("en")
            lang = load_lang_from_file()
            if lang == "en" and os.path.exists(en_file) and not os.path.exists(zh_file):
                print("OK: 切换回 en 成功，zh 文件已被清理")
            else:
                print(f"FAIL: 切换回 en 失败，返回值为 {lang}")
        except Exception as e:
            print(f"FAIL: 切换 en 测试崩溃 - {e}")

        # 测试 4：验证带变量的翻译接口
        try:
            set_lang("zh")
            res = t("status_recognizing", cur=2, total=10)
            if "2" in res and "10" in res:
                print(f"OK: 翻译变量替换成功 -> {res}")
            else:
                print(f"FAIL: 翻译变量替换失败 -> {res}")
        except Exception as e:
            print(f"FAIL: 翻译接口测试崩溃 - {e}")
        # ================== 字典一致性校验 ==================
        zh_keys = set(STRINGS["zh"].keys())
        en_keys = set(STRINGS["en"].keys())
        print(f"📊 统计信息：中文 {len(zh_keys)} 条，英文 {len(en_keys)} 条")

        if zh_keys != en_keys:
            print(f"⚠️ 警告：中英文 Key 不匹配！")
            if zh_keys - en_keys:
                print(f"  中文有但英文缺: {zh_keys - en_keys}")
            if en_keys - zh_keys:
                print(f"  英文有但中文缺: {en_keys - zh_keys}")
        else:
            print("✅ 完美：中英文 Key 完全对齐，一条不多一条不少！")
# ======================================================
        print("=== 测试结束 ===")