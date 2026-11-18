import os
import sys

# 必须在任何 paddle 相关 import 之前设置：paddle 加载 libiomp5md.dll 后此变量无效，
# 且在 OpenMP 冲突场景下会直接 OSError #15 崩溃
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import shutil, subprocess
import re
import time
import ctypes
import threading
import tempfile
import shutil
import tarfile
import urllib.request
import numpy as np
import fitz
from PIL import Image, ImageOps
from docx import Document
from docx.oxml.ns import qn

import paddleocr
from paddleocr import PaddleOCR, PPStructure

from utils import log, get_base_dir

# ================== 平台检测 ==================
CURRENT_PLATFORM = "Windows"
if sys.platform.startswith("linux"):
    CURRENT_PLATFORM = "Linux"
elif sys.platform == "darwin":
    CURRENT_PLATFORM = "macOS"
    
def _detect_cjk_font() -> str | None:
    """
    探测当前平台可用的中文字体名。
    优先级：系统 API 动态查询 -> 退化到已知硬编码候选 -> 放弃 (None)
    """
    # ====== Windows ======
    if CURRENT_PLATFORM == "Windows":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                installed = {}
                i = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, i)
                        installed[name.lower()] = value.lower()
                        i += 1
                    except OSError:
                        break
            for candidate in ["yahei", "simhei", "simsun"]:
                for name, filename in installed.items():
                    if candidate in name or candidate in filename:
                        if candidate == "yahei":
                            return "Microsoft YaHei"
                        elif candidate == "simhei":
                            return "SimHei"
                        elif candidate == "simsun":
                            return "SimSun"
        except Exception:
            pass

        fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        for fname, font_name in [
            ("msyh.ttc",   "Microsoft YaHei"),
            ("simhei.ttf", "SimHei"),
            ("simsun.ttc", "SimSun"),
        ]:
            if os.path.exists(os.path.join(fonts_dir, fname)):
                return font_name
        return None

    # ====== macOS ======
    if CURRENT_PLATFORM == "macOS":
        if os.path.exists("/System/Library/Fonts/PingFang.ttc"):
            return "PingFang SC"
        if shutil.which("fc-match"):
            try:
                out = subprocess.run(["fc-match", ":lang=zh", "family"],
                                     capture_output=True, text=True, timeout=2).stdout
                if out.strip():
                    return out.splitlines()[0].split(",")[0].strip()
            except Exception:
                pass
        return None

    # ====== Linux ======
    if shutil.which("fc-match"):
        try:
            out = subprocess.run(["fc-match", ":lang=zh", "family"],
                                 capture_output=True, text=True, timeout=2).stdout
            if out.strip():
                return out.splitlines()[0].split(",")[0].strip()
        except Exception:
            pass
    if shutil.which("fc-list"):
        try:
            out = subprocess.run(["fc-list", ":lang=zh", "family"],
                                 capture_output=True, text=True, timeout=2).stdout
            if out.strip():
                return out.splitlines()[0].split(",")[0].strip()
        except Exception:
            pass
    for path, font_name in [
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "Noto Sans CJK SC"),
        ("/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc", "Noto Sans CJK SC"),
        ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", "WenQuanYi Zen Hei"),
    ]:
        if os.path.exists(path):
            return font_name
    return None   

def _apply_docx_font(doc_word):
    """
    根据探测结果应用字体。
    - 探测到具体字体名 → try 尝试设置
    - 探测不到（None）→ 直接跳过，走 docx 默认
    """
    font_name = _detect_cjk_font()
    if not font_name:
        return

    try:
        style = doc_word.styles['Normal']
        style.font.name = font_name
        style._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    except Exception as e:
        log.warning("设置默认字体失败 / Failed to set default font: %s", e)

class ConversionCancelled(Exception):
    pass


# 中英文双写：CDLA 训练模型输出 '正文'/'标题'/'表格'，
# PubLayNet 训练模型输出 'text'/'title'/'table'，
# 两种模型的 type 字段语言不同，任一命中即走对应分支
TEXT_REGION_TYPES = frozenset({
    'text', 'title', 'list',
    '正文', '标题', '列表',
})

TABLE_REGION_TYPES = frozenset({
    'table',
    '表格',
})

# 已知非文字区域：识别到后跳过，不写入 Word。
# 单独收集成集合是为了和“未知类型”区分——未知类型会打 info 日志便于排查新模型输出
SKIP_REGION_TYPES = frozenset({
    'figure', 'figure_caption', 'table_caption',
    'header', 'footer', 'reference', 'equation',
    '图片', '图注', '表注',
    '页眉', '页脚', '参考文献', '公式',
})


def _get_total_memory_gb() -> float:
    # Windows：GlobalMemoryStatusEx 返回 ullTotalPhys 单位是字节，
    # 除以 1024^3 得 GB；MEMORYSTATUSEX 结构体字段顺序不能乱，
    # dwLength 必须先填 size 否则 API 拒绝执行
    try:
        if sys.platform == "win32":
            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]
            stat = MEMORYSTATUSEX()
            stat.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat))
            return stat.ullTotalPhys / (1024 ** 3)
        elif sys.platform == "darwin":
            return os.sysconf('SC_PAGE_SIZE') * os.sysconf('SC_PHYS_PAGES') / (1024 ** 3)
        else:
            # /proc/meminfo 的 MemTotal 单位是 kB，除以 1024^2 得 GB
            with open('/proc/meminfo', 'r') as f:
                for line in f:
                    if line.startswith('MemTotal:'):
                        return int(line.split()[1]) / (1024 ** 2)
    except Exception as e:
        # 读不到内存时返回 0：调用方按 <=8GB 走 fast_model，
        # 宁可给低配模型也不抛异常阻断启动
        log.warning("内存检测失败，默认按低配处理: %s", e)
        log.warning("Memory detection failed, default to low configuration processing: %s", e)
    return 0.0


# timeout=1.5 而非 3.0：两个 URL 串行，最坏 1.5*2=3 秒。
# 第一个 URL 成功则立即返回，不再等第二个（串行相比并行 with 版的优势）
def _is_online(timeout: float = 1.5) -> bool:
    for url in [
        "https://paddleocr.bj.bcebos.com/",
        "https://www.baidu.com/",
    ]:
        try:
            # HEAD 而非 GET：只要 200 响应头，不拉 body，
            # 且 urllib 自动走 HTTPS_PROXY 环境变量，代理环境下能正确判在线
            req = urllib.request.Request(url, method="HEAD")
            urllib.request.urlopen(req, timeout=timeout).close()
            return True
        except Exception:
            continue
    return False


def _is_auto_download_enabled() -> bool:
    # 默认 "1"：无环境变量时开启自动下载；设为 "0" 可强制离线模式
    return os.environ.get("PDF2WORD_AUTO_DOWNLOAD", "1") == "1"


# 每个 key 对应 {model_root}/{key}/ 目录下的一个模型。
# name 是官方 tar 解压后的原始目录名，用于兼容旧版本直接解压到 model_root 下的布局
MODEL_REGISTRY = {
    "det": {
        "name": "ch_PP-OCRv4_det_infer",
        "url": "https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_det_infer.tar",
    },
    "rec_en": {
        "name": "en_PP-OCRv4_rec_infer",
        "url": "https://paddleocr.bj.bcebos.com/PP-OCRv4/english/en_PP-OCRv4_rec_infer.tar",
    },
    "rec_ch": {
        "name": "ch_PP-OCRv4_rec_infer",
        "url": "https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_rec_infer.tar",
    },
    "cls": {
        "name": "ch_ppocr_mobile_v2.0_cls_infer",
        "url": "https://paddleocr.bj.bcebos.com/dygraph_v2.0/ch/ch_ppocr_mobile_v2.0_cls_infer.tar",
    },
    "det_server": {
        "name": "ch_PP-OCRv4_det_server_infer",
        "url": "https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_det_server_infer.tar",
    },
    "rec_ch_server": {
        "name": "ch_PP-OCRv4_rec_server_infer",
        "url": "https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_rec_server_infer.tar",
    },
    # CDLA 训练而非 PubLayNet：CDLA 的 10 类（含页眉/页脚/公式）与 layout_cdla_dict.txt 对齐，
    # PubLayNet 只有 5 类，会把页眉页脚误认成正文
    "layout": {
        "name": "picodet_lcnet_x1_0_fgd_layout_cdla_infer",
        "url": "https://paddleocr.bj.bcebos.com/ppstructure/models/layout/picodet_lcnet_x1_0_fgd_layout_cdla_infer.tar",
    },
    "table": {
        "name": "ch_ppstructure_mobile_v2.0_SLANet_infer",
        "url": "https://paddleocr.bj.bcebos.com/ppstructure/models/slanet/ch_ppstructure_mobile_v2.0_SLANet_infer.tar",
    },
}


def _get_models_root() -> str:
    # 环境变量优先：测试或自定义部署时覆盖模型路径
    env_dir = os.environ.get("PDF2WORD_MODELS_DIR")
    if env_dir:
        return os.path.abspath(env_dir)

    # 以下五个候选依次尝试，第一个真实存在的目录胜出：
    # 1. exe/脚本同级（get_base_dir 内部做了 frozen/__file__/argv 兜底）
    # 2. 本模块 __file__ 同级
    # 3. 入口脚本 argv[0] 同级
    # 4. 当前 cwd
    # 5. 都找不到时返回第 1 个候选路径（调用方会 os.makedirs）
    candidates = []
    try:
        candidates.append(os.path.join(get_base_dir(), "models"))
    except Exception:
        pass

    mod = sys.modules.get(__name__)
    f = getattr(mod, "__file__", None) if mod else None
    if f and not f.startswith("<"):
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(f)), "models"))

    if sys.argv and sys.argv[0]:
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "models"))

    candidates.append(os.path.join(os.getcwd(), "models"))

    for c in candidates:
        if os.path.isdir(c):
            return c
    return candidates[0]


def _find_model_file_dir(base_path: str) -> str:
    # 支持两种推理模型文件：
    # det/rec/cls：inference.pdmodel
    # layout / table(ppstructure)：model.pdmodel
    infer1 = os.path.exists(os.path.join(base_path, "inference.pdmodel"))
    infer2 = os.path.exists(os.path.join(base_path, "model.pdmodel"))
    if infer1 or infer2:
        return base_path
    if os.path.isdir(base_path):
        for item in os.listdir(base_path):
            sub_path = os.path.join(base_path, item)
            if os.path.isdir(sub_path):
                s1 = os.path.exists(os.path.join(sub_path, "inference.pdmodel"))
                s2 = os.path.exists(os.path.join(sub_path, "model.pdmodel"))
                if s1 or s2:
                    return sub_path
    return base_path


def _resolve_local_model(model_root: str, key: str) -> str | None:
    meta = MODEL_REGISTRY.get(key)
    if not meta:
        return None

    # 先查语义目录 {model_root}/{key}/（如 fast_model/det/），
    # 这是本代码下载后归位的布局，用户手动放置也推荐这个
    legacy = os.path.join(model_root, key)
    p = _find_model_file_dir(legacy)
    has_infer = os.path.exists(os.path.join(p, "inference.pdmodel"))
    has_model = os.path.exists(os.path.join(p, "model.pdmodel"))
    if p and (has_infer or has_model):
        return p

    # 再查原名目录 {model_root}/{name}/（如 fast_model/ch_PP-OCRv4_det_infer/），
    # 兼容用户直接解压 tar 不解压到语义目录的场景
    modern = os.path.join(model_root, meta["name"])
    p = _find_model_file_dir(modern)
    has_infer = os.path.exists(os.path.join(p, "inference.pdmodel"))
    has_model = os.path.exists(os.path.join(p, "model.pdmodel"))
    if p and (has_infer or has_model):
        return p

    return None


def _download_and_extract(url: str, target_key_dir: str, timeout: int = 60) -> bool:
    tmp_tar = None
    tmp_extract = None
    try:
        log.info("开始下载模型: %s", url)
        log.info("Start downloading the models: %s", url)

        # delete=False：临时文件需要持有一个路径供 tarfile 打开，
        # 关闭 NamedTemporaryFile 的句柄时不能自动删
        with tempfile.NamedTemporaryFile(suffix=".tar", delete=False) as tmp:
            tmp_tar = tmp.name

        with urllib.request.urlopen(url, timeout=timeout) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            last_log = 0
            # 每次 read 8192 字节，每累计 1MB 打一条进度日志：
            # 模型 4-100MB，1MB 粒度足以让用户感知进度且不刷屏
            with open(tmp_tar, "wb") as f:
                while True:
                    chunk = resp.read(8192)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if downloaded - last_log >= 1024 * 1024:
                        log.info("下载进度: %.1f / %.1f MB",
                                 downloaded / 1024 / 1024,
                                 total / 1024 / 1024 if total else 0)
                        log.info("Download progress: %.1f / %.1f MB",
                                 downloaded / 1024 / 1024,
                                 total / 1024 / 1024 if total else 0)
                        last_log = downloaded

        log.info("下载完成，开始解压...")
        log.info("Download completed, start decompressing...")

        tmp_extract = tempfile.mkdtemp(prefix="pdf2word_model_")
        with tarfile.open(tmp_tar, "r") as tar:
            # 防 CVE-2007-4559（tar 路径穿越）：tar 成员名可能含 ../ 或绝对路径，
            # 逐个校验解压后的绝对路径必须落在 extract_root 内
            extract_root = os.path.abspath(tmp_extract)
            for m in tar.getmembers():
                target = os.path.abspath(os.path.join(extract_root, m.name))
                if not target.startswith(extract_root + os.sep) and target != extract_root:
                    raise RuntimeError(f"tar 内含不安全路径 / Unsafe path in tar: {m.name}")
            # filter="data" 是 Python 3.12+ 的官方安全解压模式；
            # 低版本抛 TypeError，退回普通 extractall（已由上面的路径校验兜底）
            try:
                tar.extractall(extract_root, filter="data")
            except TypeError:
                tar.extractall(extract_root)

        # 兼容两种模型标志：inference.pdmodel（ocr） / model.pdmodel(layout/table)
        src_dir = None
        for dirpath, _, filenames in os.walk(tmp_extract):
            if "inference.pdmodel" in filenames or "model.pdmodel" in filenames:
                src_dir = dirpath
                break

        if not src_dir:
            raise RuntimeError("解压后未找到推理权重文件 / No inference weight file found after extraction(inference.pdmodel / model.pdmodel)")

        # 下载完成后移动到语义目录 {model_root}/{key}/：
        # 先 rmtree 旧目录，避免用户之前放了残缺模型残留文件；
        # move 后 src_dir 自带权重，_resolve_local_model 一定能命中
        if os.path.exists(target_key_dir):
            shutil.rmtree(target_key_dir, ignore_errors=True)
        os.makedirs(os.path.dirname(target_key_dir), exist_ok=True)
        shutil.move(src_dir, target_key_dir)
        log.info("模型已就位: %s", target_key_dir)
        log.info("The model is in place: %s", target_key_dir)
        return True

    except Exception as e:
        log.error("模型下载/解压失败: %s", e)
        log.error("Model download/decompression failed: %s", e)
        return False
    finally:
        # 无论成功失败都清理：避免残留几个 GB 的 tar 和临时目录塞满用户硬盘
        if tmp_tar and os.path.exists(tmp_tar):
            try:
                os.remove(tmp_tar)
            except Exception:
                pass
        if tmp_extract and os.path.exists(tmp_extract):
            try:
                shutil.rmtree(tmp_extract, ignore_errors=True)
            except Exception:
                pass


def _ensure_model(model_root: str, key: str, online_fn, auto_download: bool,
                  invalidate_online_fn=None) -> str | None:
    local = _resolve_local_model(model_root, key)
    if local:
        log.info("模型 [%s] 本地命中: %s", key, local)
        log.info("Model [%s] Local hit: %s", key, local)
        return local

    meta = MODEL_REGISTRY.get(key)
    if not meta:
        log.error("未知模型 key: %s", key)
        log.error("Unknown model key: %s", key)
        return None

    if not auto_download:
        log.error("模型 [%s] 缺失，且 PDF2WORD_AUTO_DOWNLOAD=0 禁用下载", key)
        log.error("Model [%s] missing, and PDF2WORD_AUTO_DOWNLOAD=0 disables download", key)
        return None

    # 缺失时才检测网络：本地命中时不触发网络请求，避免每次启动多花 1.5-3 秒
    if not online_fn():
        log.error("模型 [%s] 缺失，当前离线模式无法下载。请手动放置到: %s",
                  key, os.path.join(model_root, key))
        log.error("Model [%s] missing. Offline mode cannot download. Please place it manually at: %s",
                  key, os.path.join(model_root, key))
        return None

    target_dir = os.path.join(model_root, key)
    log.warning("模型 [%s] 缺失，开始自动下载 → %s", key, target_dir)
    log.warning("Model [%s] missing，Start auto download → %s", key, target_dir)
    if _download_and_extract(meta["url"], target_dir):
        return _resolve_local_model(model_root, key)

    # 下载失败 = "假在线"（能 ping 通但下载被防火墙拦）：
    # 清空在线缓存，下次 _check_online 会重新请求，避免一直误判为在线反复重试
    log.warning("模型 [%s] 下载失败，清空网络状态缓存以便下次重测", key)
    log.warning("Model [%s] Download failed, clear network status cache for next retest", key)
    if invalidate_online_fn:
        try:
            invalidate_online_fn()
        except Exception:
            pass
    return None


def _bbox_to_xyxy(bbox):
    # 兼容两种 bbox 格式：
    # A. [x1, y1, x2, y2]                （PP-Structure layout 输出）
    # B. [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]（PaddleOCR det 输出，四点）
    # 四点取 min/max 转成轴对齐矩形；解析失败返回 None 由调用方跳过该区域
    if not bbox:
        return None
    try:
        if len(bbox) == 4 and not isinstance(bbox[0], (list, tuple)):
            return float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])
        if len(bbox) >= 4 and isinstance(bbox[0], (list, tuple)):
            xs = [float(p[0]) for p in bbox]
            ys = [float(p[1]) for p in bbox]
            return min(xs), min(ys), max(xs), max(ys)
    except Exception as e:
        log.warning("bbox 解析失败: %r (%s)", bbox, e)
        log.warning("bbox parsing failed: %r (%s)", bbox, e)
    return None


def _region_sort_key(region):
    # round(y1/20)：把 y 坐标按 20 像素分桶，
    # 同一行（y 差 < 20）的左右双栏区域会被分到同一桶，桶内按 x 排序，
    # 避免双栏文档被排成"左一段、右一段、左一段、右一段"的交错顺序
    xyxy = _bbox_to_xyxy(region.get('bbox'))
    if not xyxy:
        return (float('inf'), float('inf'))
    x1, y1, _, _ = xyxy
    return (round(y1 / 20), x1)


def _find_dict(name: str) -> str | None:
    # 字典文件（layout_cdla_dict.txt / table_structure_dict_ch.txt / ppocr_keys_v1.txt 等）
    # 随 paddleocr 包分发，路径是 paddleocr/ppocr/utils/xxx。
    # 逐层 os.walk 查找而非写死路径，兼容 pip 与便携版的不同安装布局
    root = os.path.dirname(paddleocr.__file__)
    for dirpath, _, filenames in os.walk(root):
        if name in filenames:
            return os.path.join(dirpath, name)
    log.warning("未找到字典文件: %s", name)
    log.warning("Dictionary file not found: %s", name)
    return None


_engine_instances = {}
_engine_lock = threading.Lock()


def get_engine(lang='ch') -> "OCREngine":
    # 双重检查锁定：首次创建实例时加锁，后续命中直接返回。
    # 实例只创建不加载模型（懒加载），创建本身毫秒级
    global _engine_instances
    with _engine_lock:
        if lang not in _engine_instances:
            _engine_instances[lang] = OCREngine(default_lang=lang)
        return _engine_instances[lang]


def is_engine_loaded(lang='ch') -> bool:
    # 判断"模型是否已加载到内存"而非"实例是否已创建"：
    # 实例创建是毫秒级，模型加载是 10-30 秒，GUI 侧靠这个区分
    # "正在加载模型（首次较慢）"和"正在启动引擎"两种提示
    engine = _engine_instances.get(lang)
    if engine is None:
        return False
    if engine.mode == "layout_bilingual":
        return engine._layout_engine is not None and engine._ocr_ch is not None
    return engine._ocr_single is not None


class OCREngine:
    # 在线状态缓存 300 秒：网络环境变化（如从 WiFi 切到离线）后最多 5 分钟会重新检测。
    # 太短会导致频繁发 HEAD 请求，太长会导致网络恢复后模型缺失仍然判离线
    ONLINE_TTL_SECONDS = 300

    def __init__(self, default_lang='ch'):
        self._ocr_single = None
        self._ocr_ch = None
        self._ocr_en = None
        self._layout_engine = None
        self._pro_models_cache = None

        self.default_lang = default_lang

        self._online_cached = None
        self._online_ts = 0.0
        self.auto_download = _is_auto_download_enabled()

        # 阈值 8GB：<=8G 只加载轻量单模型（省内存），>8G 加载 PP‑Structure + 双模型仲裁（高精度）
        mem_gb = _get_total_memory_gb()
        models_root = _get_models_root()

        if mem_gb <= 8.0:
            self.model_dir = os.path.join(models_root, "fast_model")
            self.mode = "ocr_single"
        else:
            self.model_dir = os.path.join(models_root, "pro_model")
            self.mode = "layout_bilingual"

        try:
            os.makedirs(self.model_dir, exist_ok=True)
        except OSError as e:
            # 装到 C:\Program Files 等无写权限目录时提前报错，
            # 而非等到下载模型时才在深层调用栈抛裸 PermissionError
            raise RuntimeError(
                f"无法创建模型目录: {self.model_dir}\n"
                f"Failed to create model directory: {self.model_dir}\n"
                f"请检查程序目录写权限，或设置环境变量 PDF2WORD_MODELS_DIR 指向可写目录。\n"
                f"Please check directory write permissions, or set PDF2WORD_MODELS_DIR to a writable path.\n"
                f"Original error: {e}\n"
                f"原始错误: {e}"
            ) from e

        log.info(
            "引擎初始化: 内存=%.1fGB | 模式=%s | 自动下载=%s | 模型目录=%s",
            mem_gb, self.mode, self.auto_download, self.model_dir
        )
        log.info(
            "Engine initialization: memory=%.1fGB | pattern=%s | auto-download=%s | Model Catalog=%s",
            mem_gb, self.mode, self.auto_download, self.model_dir
        )

    def _check_online(self, force: bool = False) -> bool:
        now = time.time()
        expired = (now - self._online_ts) > self.ONLINE_TTL_SECONDS
        if force or self._online_cached is None or expired:
            log.info("检测网络状态...")
            log.info("Detecting network status...")
            self._online_cached = _is_online()
            self._online_ts = now
            log.info("网络状态: %s", "在线" if self._online_cached else "离线")
            log.info("network status: %s", "online" if self._online_cached else "offline")
        return self._online_cached

    def _invalidate_online(self):
        # 下载失败时由 _ensure_model 调用：清空缓存让下次 _check_online 重新发请求
        self._online_cached = None
        self._online_ts = 0.0

    def _pro_cache_valid(self) -> bool:
        # 缓存后每次使用前校验：用户在两份 PDF 之间手动删了模型会导致路径失效，
        # 缓存命中后 PaddleOCR 加载会失败，不如重新查盘走自动下载
        if self._pro_models_cache is None:
            return False
        for p in self._pro_models_cache:
            if p is None:
                continue
            exist1 = os.path.exists(os.path.join(p, "inference.pdmodel"))
            exist2 = os.path.exists(os.path.join(p, "model.pdmodel"))
            if not (exist1 or exist2):
                log.warning("pro 模型缓存失效（模型被删除或移动）: %s", p)
                log.warning("pro model cache failure (model deleted or moved): %s", p)
                return False
        return True

    def _setup_models_single(self):
        det = _ensure_model(self.model_dir, "det", self._check_online,
                            self.auto_download, self._invalidate_online)
        if not det:
            raise FileNotFoundError(
                f"缺少 det 检测模型，无法启动 OCR。\n"
                f"Missing det detection model, unable to start OCR.\n"
                f"请检查目录: {self.model_dir}/det/\n"
                f"Please check directory: {self.model_dir}/det/\n"
                f"（离线环境请手动放置模型，或将 PDF2WORD_AUTO_DOWNLOAD 保持开启并联网重试）\n"
                f"(For offline environment, please place the model manually, or keep PDF2WORD_AUTO_DOWNLOAD enabled and retry online)"
            )

        # rec_en 优先：英语老师是主要用户，纯英文文档 en 模型识别率高于 ch 模型。
        # rec_en 命中后不再尝试 rec_ch（避免多下载 10MB ch 模型浪费首屏时间）
        rec_en = _ensure_model(self.model_dir, "rec_en", self._check_online,
                               self.auto_download, self._invalidate_online)
        if rec_en:
            rec_path, rec_lang = rec_en, 'en'
        else:
            log.warning("rec_en 缺失，回退尝试 rec_ch（lang=ch，英文识别精度略低）")
            log.warning("rec_en missing, falling back to rec_ch (lang=ch, slightly lower English accuracy)")
            rec_ch = _ensure_model(self.model_dir, "rec_ch", self._check_online,
                                   self.auto_download, self._invalidate_online)
            if not rec_ch:
                raise FileNotFoundError(
                    f"缺少识别模型 rec_en / rec_ch。\n"
                    f"Missing recognition model rec_en / rec_ch.\n"
                    f"请检查目录: {self.model_dir}/\n"
                    f"Please check directory: {self.model_dir}/"
                )
            rec_path, rec_lang = rec_ch, 'ch'

        # cls 方向分类是可选模型：缺失时 use_angle_cls=False，
        # 整页 OCR 无法矫正倒置文字，但不会因联网下载阻塞启动
        cls = _ensure_model(self.model_dir, "cls", self._check_online,
                            self.auto_download, self._invalidate_online)
        use_cls = cls is not None
        if not use_cls:
            log.warning("cls 模型缺失，单模型模式跳过方向矫正（倒置扫描件可能识别不准）")
            log.warning("cls model missing, single model mode skipping direction correction (inverted scan may not be recognized accurately)")

        return det, rec_path, rec_lang, cls, use_cls

    def _setup_models_pro(self):
        # 缓存：_init_layout_engine 和 _init_bilingual_ocr 各调一次，
        # 缓存后第二次直接返回，省 5 次查盘 + 5 条日志
        if self._pro_cache_valid():
            return self._pro_models_cache
        self._pro_models_cache = None

        layout = _ensure_model(self.model_dir, "layout", self._check_online,
                               self.auto_download, self._invalidate_online)
        if not layout:
            raise FileNotFoundError(
                f"pro_model 缺少 layout 版面分析模型。\n"
                f"pro_model is missing the layout analysis model.\n"
                f"请检查目录: {self.model_dir}/layout/\n"
                f"Please check directory: {self.model_dir}/layout/\n"
                f"（离线环境请手动放置；或将内存阈值调至 8GB 以下使用 fast_model）\n"
                f"(For offline environment, please place it manually; or lower the memory threshold below 8GB to use fast_model)"
            )

        # table 可选：缺失时 table=False，PP‑Structure 跳过表格识别，
        # 文字区域不受影响
        table = _ensure_model(self.model_dir, "table", self._check_online,
                              self.auto_download, self._invalidate_online)
        if not table:
            log.warning("table 模型缺失，表格识别已禁用（文字区域不受影响）")
            log.warning("table model missing, table recognition is disabled (text areas are not affected)")

        # det_server 优先：高精度检测，缺失退回普通 det 保证能跑
        det = _ensure_model(self.model_dir, "det_server", self._check_online,
                            self.auto_download, self._invalidate_online)
        if not det:
            log.info("det_server 缺失，回退到普通 det")
            log.info("det_server is missing, falling back to regular det")
            det = _ensure_model(self.model_dir, "det", self._check_online,
                                self.auto_download, self._invalidate_online)
        if not det:
            raise FileNotFoundError("缺少 det 检测模型，无法启动 OCR / Missing det detection model, can't start OCR")

        rec_en = _ensure_model(self.model_dir, "rec_en", self._check_online,
                               self.auto_download, self._invalidate_online)

        # rec_ch_server 优先：高精度中文识别，缺失退回普通 rec_ch。
        # rec_ch 是双模型仲裁的"必须"模型（en 缺失时降级为单 ch 也要用它）
        rec_ch = _ensure_model(self.model_dir, "rec_ch_server", self._check_online,
                               self.auto_download, self._invalidate_online)
        if not rec_ch:
            log.info("rec_ch_server 缺失，回退到普通 rec_ch")
            log.info("rec_ch_server is missing, falling back to regular rec_ch")
            rec_ch = _ensure_model(self.model_dir, "rec_ch", self._check_online,
                                   self.auto_download, self._invalidate_online)
        if not rec_ch:
            raise FileNotFoundError("缺少 rec_ch 中文识别模型，无法启动双模型仲裁 / Missing rec_ch Chinese recognition model, can't start dual-model arbitration")

        result = (layout, table, det, rec_en, rec_ch)
        self._pro_models_cache = result
        return result

    def _init_ocr_single(self, cancel_event=None):
        if self._ocr_single is not None:
            return
        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("模型加载前取消 / Cancel before the model loads")

        det, rec_path, rec_lang, cls, use_cls = self._setup_models_single()

        try:
            log.info("正在加载单模型 OCR: lang=%s, use_cls=%s", rec_lang, use_cls)
            log.info("Loading single-model OCR: lang=%s, use_cls=%s", rec_lang, use_cls)
            # PaddleOCR 初始化内部若缺模型会 sys.exit()，
            # SystemExit 不被 except Exception 捕获，必须单独拦
            self._ocr_single = PaddleOCR(
                use_angle_cls=use_cls,
                lang=rec_lang,
                det_model_dir=det,
                rec_model_dir=rec_path,
                cls_model_dir=cls,
                show_log=False,
                use_gpu=False,
            )
        except SystemExit as e:
            log.error("PaddleOCR 内部强制退出: %s", e)
            log.error("PaddleOCR internal forced exit: %s", e)
            raise RuntimeError(f"OCR 引擎内部异常退出，请检查模型完整性。 / The OCR engine exited unexpectedly, please check if the model is intact.\n 错误码 / Error code: {e}")
        except Exception:
            log.error("PaddleOCR 初始化失败", exc_info=True)
            log.error("PaddleOCR failed to initialize", exc_info=True)
            raise

        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("模型加载后取消 / Cancel after the model loads")

    def _ensure_ppstructure_det_cache(self, det_path: str = None) -> bool:
        """把本地 det 模型复制到 PPStructure 内部硬编码要的默认缓存目录。

        背景：PaddleOCR 2.9.1 的 TableSystem 内部 copy.deepcopy(args) 会丢失
        det_model_dir，TextDetector 回退到 ~/.paddleocr/whl/det/ch/ch_PP-OCRv4_det_infer/。
        本方法在该路径预先放好文件，避免 TextDetector 因找不到文件而崩。

        前置保证：_setup_models_pro 已确保本地 det_server 或 det 存在。
        原则：只复制 det 一个模型，目标文件必须是本地文件的副本，绝不联网。

        参数：
          det_path: 调用方已解析好的 det 实际路径（_setup_models_pro 返回值）。
                    传入则直接用，不传则用 _resolve_local_model 回退解析。

        返回：
          True  = 目标目录已完整（已存在或复制并验证成功），可安全启用 table
          False = 任何一步失败（源缺失/权限/磁盘/复制后验证失败），调用方禁用 table
        """
        try:
            # ① 找源：优先用调用方传入的 det_path（_setup_models_pro 已解析过，
            # 兼容语义 key 名 + 官方原名两种布局）；未传时用 _resolve_local_model
            # 重新解析，同样兼容双布局，避免自己手写目录遍历漏掉官方原名布局。
            src_dir = None
            if det_path and os.path.isdir(det_path):
                src_dir = det_path
                log.info("PPStructure det 缓存源使用传入路径: %s", src_dir)
                log.info("PPStructure det caches the source using the incoming path: %s", src_dir)
            else:
                for key in ("det_server", "det"):
                    resolved = _resolve_local_model(self.model_dir, key)
                    if resolved:
                        src_dir = resolved
                        break

            if src_dir is None:
                log.error("本地未找到 det/det_server 模型，无法填充 PPStructure 缓存")
                log.error("The det/det_server model wasn't found locally, so can't fill the PPStructure buffer.")
                return False

            # ② 目标路径 + 完整性检查
            target_dir = os.path.join(
                os.path.expanduser("~"),
                ".paddleocr", "whl", "det", "ch", "ch_PP-OCRv4_det_infer",
            )
            required = ("inference.pdmodel", "inference.pdiparams")

            if os.path.isdir(target_dir) and all(
                os.path.exists(os.path.join(target_dir, f)) for f in required
            ):
                log.info("PPStructure det 缓存已完整，跳过复制: %s", target_dir)
                log.info("PPStructure det cache is complete, skipping copy: %s", target_dir)
                return True

            # ③ 复制（文件名 model.* → inference.*）
            os.makedirs(target_dir, exist_ok=True)

            def _copy_any(src_names, dst_name):
                for src_name in src_names:
                    src = os.path.join(src_dir, src_name)
                    if os.path.exists(src):
                        shutil.copy2(src, os.path.join(target_dir, dst_name))
                        return True
                return False

            _copy_any(["inference.pdmodel", "model.pdmodel"], "inference.pdmodel")
            _copy_any(["inference.pdiparams", "model.pdiparams"], "inference.pdiparams")
            _copy_any(
                ["inference.pdiparams.info", "model.pdiparams.info"],
                "inference.pdiparams.info",
            )

            # ④ 复制后验证：文件真的在了才算成功
            missing = [f for f in required if not os.path.exists(os.path.join(target_dir, f))]
            if missing:
                log.warning(
                    "PPStructure det 缓存复制后验证失败，缺失: %s（目标: %s）",
                    missing, target_dir,
                )
                log.warning(
                    "PPStructure det cache verification failed after copy, missing: %s (target: %s)",
                    missing, target_dir,
                )
                return False

            log.info("已从本地复制 det 到 PPStructure 缓存: %s → %s", src_dir, target_dir)
            log.info("Copied det from local to PPStructure cache: %s → %s", src_dir, target_dir)
            return True

        except PermissionError as e:
            log.warning("默认缓存目录不可写，禁用表格识别: %s", e)
            log.warning("The default cache directory is not writable, so table recognition is disabled: %s", e)
            return False
        except OSError as e:
            log.warning("默认缓存填充时磁盘/路径异常，禁用表格识别: %s", e)
            log.warning("Disable table recognition when there's a disk/path error during default cache filling: %s", e)
            return False
        except Exception as e:
            log.warning("默认缓存填充发生未知异常，禁用表格识别: %s", e)
            log.warning("An unknown error occurred during default cache filling, table recognition is disabled: %s", e)
            return False

    def _init_layout_engine(self, cancel_event=None):
        if self._layout_engine is not None:
            return
        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("版面分析模型加载前取消 / Cancel before loading the layout analysis model")

        layout, table, det, rec_en, rec_ch = self._setup_models_pro()

        # ==========新增：畸形/特殊字符Windows用户名前置拦截，规避TableSystem C++崩溃 ==========
        import string
        home_path = os.path.expanduser("~")
        basename_home = os.path.basename(home_path)
        # 判定规则：用户名首字符是空白 / 路径含中文 / 用户名以 . 开头
        has_leading_space = bool(basename_home) and basename_home[0].isspace()
        has_cn_chars = any('\u4e00' <= c <= '\u9fff' for c in home_path)
        start_with_dot = basename_home.startswith(".")

        if has_leading_space or has_cn_chars or start_with_dot:
            log.warning("检测到Windows用户目录存在畸形/特殊字符[%s]，为规避PPStructure TableSystem C++层崩溃，强制禁用表格识别", home_path)
            log.warning("Detected malformed/special characters [%s] in the Windows user directory. To prevent PPStructure TableSystem C layer from crashing, table recognition is forcibly disabled.", home_path)
            table = None
        # =====================================================================================


        # ocr=False：PP‑Structure 只做版面检测（layout）+ 表格结构（table），
        # 文字识别交给外层的双模型仲裁，避免 PP‑Structure 内部再加载一套 rec 模型
        kwargs = dict(
            show_log=False,
            use_gpu=False,
            use_mp=False,
            image_orientation=False,
            layout=True,
            table=bool(table),
            ocr=False,
            layout_model_dir=layout,
        )
        if table:
            kwargs['table_model_dir'] = table

        # 字典文件找不到时不传该参数，让 PPStructure 用内部默认路径
        layout_dict = _find_dict('layout_cdla_dict.txt')
        table_dict = _find_dict('table_structure_dict_ch.txt')
        if layout_dict:
            kwargs['layout_dict_path'] = layout_dict
        if table_dict and table:
            kwargs['table_char_dict_path'] = table_dict

        # 【联动】table 启用前先确保 PPStructure 内部要的 det 缓存到位。
        # TextDetector 无视外层 det_model_dir，会回退到 ~/.paddleocr/whl/det/ch/ch_PP-OCRv4_det_infer；
        # 复制失败（权限/磁盘/源缺失）时禁用 table，避免 TableSystem 启动时崩。
        if table:
            if self._ensure_ppstructure_det_cache(det):
                log.info("PPStructure det 缓存就绪，启用表格识别")
                log.info("PPStructure det cache is ready, enable table recognition")
            else:
                log.warning("PPStructure det 缓存填充失败，本次禁用表格识别，文字区域不受影响")
                log.warning("PPStructure det cache fill failed, table recognition is disabled this time, text areas are not affected")
                table = None
                kwargs['table'] = False
                kwargs.pop('table_model_dir', None)
                kwargs.pop('table_char_dict_path', None)

        log.info("正在加载 PP‑Structure (layout+table, ocr=False)...")
        log.info("Loading PP‑Structure (layout+table, ocr=False)...")
        try:
            self._layout_engine = PPStructure(**kwargs)
        except SystemExit as e:
            log.error("PPStructure 内部强制退出: %s", e)
            log.error("PPStructure internal forced exit: %s", e)
            raise RuntimeError(f"版面分析引擎内部异常退出 / Layout analysis engine exited unexpectedly: {e}")
        except Exception:
            log.error("PPStructure 初始化失败", exc_info=True)
            log.error("PPStructure initialization failed", exc_info=True)
            raise

        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("版面分析模型加载后取消 / Cancel after the layout analysis model loads")

    def _init_bilingual_ocr(self, cancel_event=None):
        if self._ocr_ch is not None:
            return
        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("双模型加载前取消 / Cancel before loading dual models")

        layout, table, det, rec_en, rec_ch = self._setup_models_pro()

        try:
            log.info("正在加载中文识别模型 (rec_ch)...")
            log.info("Loading Chinese recognition model (rec_ch)...")
            # use_angle_cls=False：区域图从正确方向的页面裁出，无需再判 180°，
            # 省掉每区域 100‑200ms 的方向分类耗时
            self._ocr_ch = PaddleOCR(
                use_angle_cls=False,
                lang='ch',
                det_model_dir=det,
                rec_model_dir=rec_ch,
                show_log=False,
                use_gpu=False,
            )
        except SystemExit as e:
            log.error("中文模型内部强制退出: %s", e)
            log.error("Chinese model forced exit: %s", e)
            raise RuntimeError(f"中文模型初始化异常退出 / Chinese model initialization exited unexpectedly: {e}")
        except Exception:
            log.error("中文模型初始化失败", exc_info=True)
            log.error("Failed to initialize Chinese model", exc_info=True)
            raise

        if rec_en:
            try:
                log.info("正在加载英文识别模型 (rec_en)...")
                log.info("Loading English recognition model (rec_en)...")
                self._ocr_en = PaddleOCR(
                    use_angle_cls=False,
                    lang='en',
                    det_model_dir=det,
                    rec_model_dir=rec_en,
                    show_log=False,
                    use_gpu=False,
                )
            except SystemExit as e:
                log.warning("英文模型初始化异常，降级为单 ch 仲裁: %s", e)
                log.warning("English model initialization error, downgraded to single-channel arbitration: %s", e)
                
                self._ocr_en = None
            except Exception as e:
                log.warning("英文模型初始化失败，降级为单 ch 仲裁: %s", e)
                log.warning("English model initialization failed, downgraded to single-channel arbitration: %s", e)
                self._ocr_en = None
        else:
            # rec_en 加载失败不阻断：纯 ch 也能识别英文（略降精度），
            # 好过整个转换流程无法进行
            log.warning(">8GB 模式未找到 rec_en，降级为单 ch 仲裁")
            log.warning(">8GB mode did not find rec_en, downgrading to single channel arbitration")

        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("双模型加载后取消 / Cancel after loading both models")

    def _ocr_one_image(self, ocr_engine, img_np, use_cls: bool = False):
        # use_cls 参数的两条路径语义：
        # - 整页图（convert_pdf 的 _ocr_single）：True，扫描件可能倒置
        # - 区域图（_arbitrate_text 的 _ocr_ch/_ocr_en）：False，方向已由 PyMuPDF 渲染保证
        if ocr_engine is None:
            return []
        result = ocr_engine.ocr(img_np, cls=use_cls)
        if not result or not result[0]:
            return []
        # 置信度 0.5 以下丢弃：OCR 吐出的低置信度结果多为噪声或边角符号
        return [line[1][0] for line in result[0] if line[1][1] >= 0.5]

    def _arbitrate_text(self, img_np, cancel_event=None):
        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("区域识别前取消 / Cancel before area recognition")

        # 先跑 ch：ch 字典含英文，纯英文区域也能出可用结果，
        # 且能拿到中文比例判断该区域的语言属性
        texts_ch = self._ocr_one_image(self._ocr_ch, img_np, use_cls=False)
        ch_text = "".join(texts_ch)

        if ch_text:
            # 【修复】原代码 [\u4e00‑\u9fff] 中间是 U+2011（非标准连字符），
            # 会导致字符类不是范围而是三个字面字符，中文占比恒为 0。
            # 正确写法是 U+002D，表示 \u4e00 到 \u9fff 的范围。
            cn_count = len(re.findall(r'[\u4e00-\u9fff]', ch_text))
            # 分母排除空白：区域图里空格占位多，"Hello 你好" 若按全长算比例会被稀释，
            # 排除空白后更接近"有效字符中的中文占比"
            non_space = len(re.sub(r'\s', '', ch_text))
            if non_space > 0:
                if non_space < 3:
                    # 极短区域（1‑2 字符）：任何 1 个汉字都是强信号。
                    # "O" 被误认成 "口" 这类误判在 1‑2 字符区间概率很低，
                    # 直接判中文比走 10% 阈值更可靠
                    if cn_count >= 1:
                        return texts_ch
                else:
                    # 中文占比 > 10% 判中文区域：纯英文区域 ch 偶发吐 1‑2 个汉字
                    # （扫描噪声/字体问题），10% 阈值能过滤掉这种偶发误识别，
                    # 只有中英混排或纯中文区域才会走 ch
                    if cn_count / non_space > 0.1:
                        return texts_ch

        if self._ocr_en is None:
            return texts_ch

        # 纯英文区域：再跑 en 精修。en 专用模型对英文标点、连写单词、专有名词
        # 识别率高于 ch 模型
        texts_en = self._ocr_one_image(self._ocr_en, img_np, use_cls=False)
        en_text = "".join(texts_en)

        if en_text:
            # 【修复】同上，[\x20‑\x7e] 中间是 U+2011 而非 U+002D。
            en_valid = len(re.findall(r'[\x20-\x7e]', en_text))
            en_total = len(en_text.strip())
            if en_total > 0 and en_valid / en_total >= 0.7:
                return texts_en

        return texts_ch

    def _crop_region(self, img_np, bbox):
        # bbox 坐标可能有负数或超出图片边界（det 模型膨胀后的框），
        # 逐坐标 clamp 到 [0, w/h]，越界裁剪返回 None 由调用方跳过
        xyxy = _bbox_to_xyxy(bbox)
        if not xyxy:
            return None
        h, w = img_np.shape[:2]
        x1, y1, x2, y2 = xyxy
        x1 = max(0, min(int(x1), w))
        y1 = max(0, min(int(y1), h))
        x2 = max(0, min(int(x2), w))
        y2 = max(0, min(int(y2), h))
        if x2 <= x1 or y2 <= y1:
            return None
        return img_np[y1:y2, x1:x2]

    def _write_layout_result(self, result, full_img_np, doc_word, cancel_event=None, source_type="pdf"):
    # 排序失败（result 结构异常）时不阻断：用原始顺序写入，
    # 输出可能乱序，但内容不丢
        try:
            sorted_regions = sorted(result, key=_region_sort_key)
        except Exception as e:
            log.warning("区域排序失败，使用原始顺序: %s", e)
            log.warning("Region sorting failed, using original order: %s", e)
            sorted_regions = result

        for region in sorted_regions:
            r_type = (region.get('type') or '').lower()

            if r_type in TEXT_REGION_TYPES:
                # 优先用 region['img']：PP‑Structure ocr=False 时通常自带裁剪图，
                # 直接用它省一次 numpy 切片；缺失时退回 _crop_region
                region_img = region.get('img')
                if region_img is None:
                    region_img = self._crop_region(full_img_np, region.get('bbox'))

                if region_img is None or region_img.size == 0:
                    log.warning("区域裁剪失败，跳过: type=%s, bbox=%r", r_type, region.get('bbox'))
                    log.warning("Region cropping failed, skipping: type=%s, bbox=%r", r_type, region.get('bbox'))
                    region.pop('img', None)
                    continue

                texts = self._arbitrate_text(region_img, cancel_event)
                if texts:
                    # 同一区域内多行文本合并为一个段落：
                    # 避免 Word 里每个文本行都成为独立段落，视觉上过于破碎
                    doc_word.add_paragraph(" ".join(texts))

            elif r_type in TABLE_REGION_TYPES:
                res = region.get('res')
                html = ''
                extracted_list_text = ''

                if isinstance(res, dict):
                    html = res.get('html', '')
                elif isinstance(res, list):
                    # 兼容部分版本 res 为 list 的情况
                    parts = []
                    for item in res:
                        if isinstance(item, dict):
                            t = item.get('text', '')
                            if t:
                                parts.append(str(t))
                        elif item:
                            parts.append(str(item))
                    extracted_list_text = " ".join(parts).strip()

                if html:
                    # 标准 PP‑Structure 返回 dict，html 字段含表格 HTML 字符串，
                    # 用正则剥标签后以纯文本写入 Word（不是 Word 表格，但内容不丢）
                    text = re.sub(r'<[^>]+>', ' ', html)
                    text = re.sub(r'\s+', ' ', text).strip()
                    doc_word.add_paragraph(text if text else "[表格区域]")
                elif extracted_list_text:
                        doc_word.add_paragraph(extracted_list_text)
                else:
                    # 【P-1级漏洞修复】无 HTML 时的兜底逻辑（模型误判/禁用了表格）
                    log.warning("【兜底触发】发现无 HTML 的 Table 区域，强制降级分析！")
                    log.warning("[Fallback] Found Table region without HTML, forcing text extraction!")
                
                    # 如果是 PDF，先输出 5 个空行留白排版
                    if source_type == "pdf":
                        for _ in range(5):
                            doc_word.add_paragraph("")
                        
                # 无论是 IMG 还是 PDF，都强行把这块图裁下来，当文字跑 OCR
                    region_img = region.get('img')
                    if region_img is None:
                        region_img = self._crop_region(full_img_np, region.get('bbox'))
                    
                    if region_img is not None and region_img.size > 0:
                        texts = self._arbitrate_text(region_img, cancel_event)
                        if texts:
                            log.info("【兜底成功】从 Table 区域救回了 %d 行文本", len(texts))
                            log.info("[Fallback Success] Rescued %d lines of text from Table region", len(texts))
                            doc_word.add_paragraph(" ".join(texts))
                        else:
                            # 只有真正连一个字都没认出来的时候，才留占位符（或留白）
                            if source_type == "image":
                                doc_word.add_paragraph("[表格区域]")
                            else:
                                log.info("【兜底留白】该 Table 区域未提取到文字，直接跳过。")
                    else:
                        log.warning("【兜底失败】Table 区域图片裁剪失败。")
                        if source_type == "image":
                            doc_word.add_paragraph("[表格区域]")

            elif r_type in SKIP_REGION_TYPES:
                # 已知非文字区域（图片/页眉/页脚等）跳过：debug 级，
                # 正常转换不刷屏，需要排查时把 logger 调到 DEBUG 可见
                log.debug("跳过已知非文字区域: %s", r_type)
                log.debug("Skip known non-text areas: %s", r_type)
            else:
                # 未知类型：info 级，便于发现 CDLA/PubLayNet 输出了白名单未覆盖的新类型
                log.info("跳过未知区域类型: %s, bbox=%r", r_type, region.get('bbox'))
                log.info("Skip unknown area types: %s, bbox=%r", r_type, region.get('bbox'))

            # 显式释放区域裁剪图，降低一页处理过程中的峰值内存
            # （10 个区域 * 5‑20MB 裁剪图，不 pop 时 GC 来不及回收）
            region.pop('img', None)

    def convert_pdf(self, pdf_path: str, output_dir: str, progress_cb=None, cancel_event=None) -> str:
        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(f"PDF 不存在 / Doesn't exist: {pdf_path}")

        # try 从 fitz.open 之后开始：加密检查、0 页检查、模型加载、处理循环
        # 任何一步抛异常，finally 都会 close 文件句柄，
        # 防止 Windows 下模型缺失时 PDF 被锁住用户无法删除
        doc_pdf = fitz.open(pdf_path)
        try:
            if doc_pdf.is_encrypted:
                raise RuntimeError("PDF 已加密，请先解除密码保护后再转换 / The PDF is password-protected. Please remove the password before converting.")

            total = doc_pdf.page_count
            if total == 0:
                raise RuntimeError("PDF 没有页面，文件可能已损坏 / The PDF has no pages; the file might be corrupted.")

            # 先扫描全部页判断是否真的需要 OCR：
            # 纯数字版 PDF 每页都有文本层，应该完全不加载 OCR 模型（省 20 秒加载时间）
            page_texts = []
            need_ocr = False
            need_layout = False

            for i in range(total):
                try:
                    t = doc_pdf[i].get_text().strip()
                except Exception as e:
                    log.warning("扫描第 %d 页 get_text() 失败: %s", i + 1, e)
                    log.warning("Failed to get text when scanning page %d: %s", i + 1, e)
                    t = ""
                page_texts.append(t)
                # 10 字符阈值：过滤只有页码/单个标点的伪文本层，
                # 同时不会把 "Chapter 1"（9 字符）这种短标题误判成扫描件
                if len(t) < 10:
                    need_ocr = True
                    if self.mode == "layout_bilingual":
                        need_layout = True

            # 按需加载模型：三种情况加载不同组件
            if need_ocr:
                if need_layout:
                    # >8GB + 扫描件：版面分析 + 双模型仲裁
                    self._init_layout_engine(cancel_event)
                    self._init_bilingual_ocr(cancel_event)
                else:
                    # <=8GB + 扫描件：单模型整页 OCR
                    self._init_ocr_single(cancel_event)

            doc_word = Document()
            _apply_docx_font(doc_word)

            for i in range(total):
                if cancel_event and cancel_event.is_set():
                    raise ConversionCancelled("用户取消了转换 / The user canceled the conversion")

                page = doc_pdf[i]
                doc_word.add_heading(f"--- Page {i + 1} ---", level=2)

                embedded_text = page_texts[i]

                if len(embedded_text) >= 10:
                    # 文本层优先：数字版 PDF 每页 10ms，
                    # 走 OCR 每页 2 秒，相差 200 倍
                    log.info("第 %d 页检测到文本层（%d 字符），跳过 OCR", i + 1, len(embedded_text))
                    log.info("Text layer detected on page %d (%d characters), skipping OCR", i + 1, len(embedded_text))
                    for line in embedded_text.splitlines():
                        if line.strip():
                            doc_word.add_paragraph(line.strip())

                elif self.mode == "layout_bilingual":
                    log.info("第 %d 页无文本层，PP‑Structure + 双模型仲裁", i + 1)
                    log.info("Page %d has no text layer, PP‑Structure dual-model arbitration", i + 1)
                    # 3 倍缩放：小字号扫描件在 1 倍分辨率下 OCR 精度显著下降，
                    # 3 倍是精度与内存的平衡点（A4 页面 3 倍约 20‑40MB numpy）
                    pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    img_np = np.array(img)
                    del img

                    result = self._layout_engine(img_np)
                    self._write_layout_result(result, img_np, doc_word, cancel_event, source_type="pdf")
                    del img_np

                else:
                    log.info("第 %d 页无文本层，单模型 OCR", i + 1)
                    log.info("No text layer on page %d, single-model OCR", i + 1)
                    pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    img_np = np.array(img)
                    del img

                    # 整页图 use_cls=True：扫描件可能上下颠倒，需要方向分类矫正
                    texts = self._ocr_one_image(self._ocr_single, img_np, use_cls=True)
                    del img_np

                    if texts:
                        doc_word.add_paragraph(" ".join(texts))
                    else:
                        doc_word.add_paragraph("[本页未识别到文字] / [No text detected on this page]")

                # progress_cb(i+1, total)：一页处理完后回调一次，
                # GUI 侧用这个更新百分比和"正在识别：第 X/N 页"文案
                if progress_cb:
                    progress_cb(i + 1, total)

            # 输出重名防御：同名 .docx 已存在时自动加 _1、_2 后缀，
            # 不覆盖用户已有文件
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_path = os.path.join(output_dir, f"{base_name}.docx")
            counter = 1
            while os.path.exists(out_path):
                out_path = os.path.join(output_dir, f"{base_name}_{counter}.docx")
                counter += 1

            doc_word.save(out_path)
            log.info("转换完成: %s", out_path)
            log.info("Conversion complete: %s", out_path)
            return out_path

        finally:
            # close 自身失败也不能抛：例如用户提前关闭文件句柄时，
            # 这里再抛会覆盖掉原有异常，让排查时丢失真正的错误
            try:
                doc_pdf.close()
            except Exception:
                pass

    def convert_image(self, input_path: str, output_dir: str, progress_cb=None, cancel_event=None) -> str:
        if not os.path.isfile(input_path):
            raise FileNotFoundError(f"图片不存在 / The image doesn't exist: {input_path}")

        # ImageOps.exif_transpose：手机竖拍照片的像素在存储时往往是横躺的，
        # 真实方向记在 EXIF 的 Orientation 标签里。
        # 不处理时 OCR 会把竖排文字当横排识别，出乱码
        try:
            pil_img = ImageOps.exif_transpose(Image.open(input_path)).convert("RGB")
        except Exception as e:
            # from e 保留原始异常链：损坏的图片能追溯到具体解码失败位置
            raise RuntimeError(f"无法打开图片文件 / Can't open the image file {input_path}: {e}") from e

        # 按内存模式初始化模型（与 convert_pdf 一致的按需加载策略）
        if self.mode == "layout_bilingual":
            self._init_layout_engine(cancel_event)
            self._init_bilingual_ocr(cancel_event)
        else:
            self._init_ocr_single(cancel_event)

        if cancel_event and cancel_event.is_set():
            raise ConversionCancelled("模型加载后取消 / Cancel after the model loads")

        img_np = np.array(pil_img)
        del pil_img

        doc_word = Document()
        _apply_docx_font(doc_word)
        

        try:
            if self.mode == "layout_bilingual":
                log.info("图片识别模式: PP‑Structure + 双模型仲裁 (>8GB)")
                log.info("Image Recognition Mode: PP‑Structure, Dual Model Arbitration (>8GB)")
                result = self._layout_engine(img_np)
                self._write_layout_result(result, img_np, doc_word, cancel_event, source_type="image")
            else:
                log.info("图片识别模式: 单模型 OCR (<=8GB)")
                log.info("Image recognition mode: Single-model OCR (<=8GB)")
                if cancel_event and cancel_event.is_set():
                    raise ConversionCancelled("OCR 前取消 / Cancel before OCR")
                # 图片没有"页"的概念，use_cls=True 无意义（页面方向由 EXIF 已修正），
                # 传 True 只是保持与 convert_pdf 整页图路径的参数一致
                texts = self._ocr_one_image(self._ocr_single, img_np, use_cls=True)
                if cancel_event and cancel_event.is_set():
                    raise ConversionCancelled("OCR 后取消 / Cancel after OCR")
                if texts:
                    doc_word.add_paragraph(" ".join(texts))
                else:
                    doc_word.add_paragraph("[未识别到文字] / [No text recognized]")
        finally:
            # 图片 numpy 可能几十 MB，显式 del 让 GC 立即回收，
            # 不等 _write_layout_result 返回后才被清理
            del img_np

        # progress_cb(1, 1) 放在 OCR 完成后：
        # 若放在 OCR 前，GUI 大字号百分比会提前跳到 100% 却还在识别中，
        # 用户以为卡死。放在 OCR 后，100% 出现时 Word 已生成完毕
        if progress_cb:
            try:
                progress_cb(1, 1)
            except Exception:
                pass

        # 输出重名防御（同 convert_pdf）
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        out_path = os.path.join(output_dir, f"{base_name}.docx")
        counter = 1
        while os.path.exists(out_path):
            out_path = os.path.join(output_dir, f"{base_name}_{counter}.docx")
            counter += 1

        doc_word.save(out_path)
        log.info("图片转换完成: %s", out_path)
        log.info("Image conversion complete: %s", out_path)
        return out_path