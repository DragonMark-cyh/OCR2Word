# OCR2Word
Convert scanned PDFs and images to editable Word documents, powered by PaddleOCR 2.9.1.
# DragonMark‑cyh / Project: OCR2Word
> An open‑source GUI OCR tool converts scanned PDFs & images into editable Microsoft Word (.docx) documents.
> International‑oriented, with memory‑aware adaptive model loading.

## Features
- Supported input formats: **PDF, PNG, JPG, JPEG**
- Export output: Editable `.docx` Word document
- **Intelligent memory‑adaptive model loading**
  - RAM ≤ 8 GB: lightweight `fast_model`, single‑language (**English default**). Chinese recognition quality will be limited. Layout analysis and table recognition are disabled.
  - RAM > 8 GB: full‑featured `pro_model`. Enable layout analysis + **Chinese‑English dual‑model arbitration**. Automatically select optimal OCR output by confidence and text character statistics for mixed Chinese‑English content.
- Bilingual UI: English / 中文. Persist language setting by empty marker file (`en` / `zh`), no JSON config file.
- Two file import methods: system file browser or standalone drag‑and‑drop popup window.
- Smart PDF processing: Reads native embedded text layer directly for digital PDFs. OCR inference only triggers for scanned‑only pages to save waiting time.
- Auto‑download PaddleOCR inference models on first‑run. Offline deployment is supported.
- Comprehensive crash‑resistant mechanism: Global multi‑layer exception hook, rotating persistent log files. On Windows, fatal crash will automatically open log snapshot via Notepad for debugging.
- Auto detect system CJK fonts to render Chinese correctly inside exported docx.
- Duplicate filename protection: Append numeric suffix instead of overwriting existing user files.
- Support cancel ongoing conversion task or model loading procedure.
- **Cross‑platform compatible: Windows, macOS, Linux.**

> Important note: This project locks **PaddleOCR 2.9.1**. PaddleOCR 3.x introduced large‑scale breaking API refactor, migration is postponed for stability reason. In 2.x series, Chinese and English recognition models are completely separated. Mixed‑language recognition is implemented at application code layer, not built‑in inside model weight.

## Repository structure
```
DragonMark-cyh/OCR2Word/
├── main.py               # Application entry point
├── utils.py              # Path resolver, logging system & global crash handler
├── lang.py               # i18n module, empty marker‑file language persistence
├── gui.py                # Main GUI implemented by CustomTkinter
├── drag.py               # Chinese drag‑and‑drop popup sub‑program (needs separate exe build)
├── drag_en.py            # English drag‑and‑drop popup sub‑program (needs separate exe build)
├── ocr_engine.py         # Core OCR engine, memory judge, model download logic
├── requirements.txt      # Runtime dependencies
├── requirements-dev.txt  # Packaging / dev dependencies
├── en                    # Empty marker file, default startup language = English
├── models/               # Auto‑generated folder for fast_model / pro_model weights
└── logs/                 # Auto‑generated application log directory
```

## Install from source
### 1. Clone repository
```bash
git clone [https://github.com/DragonMark-cyh/OCR2Word.git](https://github.com/DragonMark-cyh/OCR2Word.git)
cd OCR2Word
```

### 2. Create virtual environment (strongly recommended)
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```
> For packaging with PyInstaller:
> ```bash
> pip install -r requirements-dev.txt
> ```

### 4. Launch application
```bash
python main.py
```

### requirements.txt
```txt
albucore==0.0.13
albumentations==1.4.10
altgraph==0.17.5
annotated-types==0.8.0
anyio==4.15.1
astor==0.8.1
beautifulsoup4==4.15.0
certifi==2026.10.23
charset-normalizer==3.5.2
cloudpickle==3.1.2
colorama==0.4.6
contourpy==1.3.2
customtkinter==5.2.2
cycler==0.12.1
Cython==3.3.0
darkdetect==0.8.0
decorator==5.3.1
exceptiongroup==1.3.1
fire==0.7.1
fonttools==4.65.0
h11==0.16.0
httpcore==1.0.9
httpx==0.28.1
idna==3.20
ImageIO==2.38.0
imgaug==0.4.0
joblib==1.6.0
kiwisolver==1.5.1
lazy-loader==0.6
lmdb==3.0.0
lxml==6.1.3
matplotlib==3.10.9
msvc-runtime==14.34.31931
networkx==3.4.2
numpy==1.24.4
opencv-python==4.6.0.66
opt-einsum==3.3.0
packaging==26.3
paddleocr==2.9.1
paddlepaddle==2.5.2
pefile==2024.8.26
pillow==10.4.0
protobuf==3.20.2
pyclipper==1.4.0
pydantic==2.14.2
pydantic_core==2.50.2
pyinstaller==6.10.0
pyinstaller-hooks-contrib==2026.9
PyMuPDF==1.24.9
PyMuPDFb==1.24.9
pyparsing==3.3.3
python-dateutil==3.0.0
python-docx==1.1.2
PyYAML==6.0.3
RapidFuzz==3.14.5
requests==2.34.2
scikit-image==0.25.2
scikit-learn==1.7.2
scipy==1.15.3
shapely==2.1.2
six==1.17.0
soupsieve==2.10
sqlite-bro==0.12.2
termcolor==3.3.0
threadpoolctl==3.7.0
tifffile==2025.5.10
tkinterdnd2==0.6.2
tomli==2.5.0
tqdm==4.70.1
typing-inspection==0.4.4
typing_extensions==4.16.0
tzdata==2026.6
urllib3==2.8.0
```

### requirements-dev.txt
```txt
pyinstaller
pyinstaller-hooks-contrib
```

## Environment Variables for custom configuration
- `PDF2WORD_BASE_DIR`: Override application root working directory
- `PDF2WORD_MODELS_DIR`: Override model weight storage directory
- `PDF2WORD_AUTO_DOWNLOAD=0`: Disable automatic model download for offline usage. You need manually place model files into target folder.

## 📦 Windows PyInstaller Packaging Guide
> **⚠️ Important**: These PyInstaller commands are Windows‑only. To build binaries for macOS / Linux you must run PyInstaller on that corresponding operating system. Cross‑compiling binaries from another OS is **not supported**.
>
> **Critical points**:
> 1. `drag.py` and `drag_en.py` **must be packaged individually into separate exe files**.
> 2. `--collect-all tkinterdnd2` is mandatory, tkinterdnd2 contains C‑extension binary resources; omitting this argument causes drag‑window runtime crash.
> 3. Use `-d` flag for debug builds during development for troubleshooting.
> 4. **Use identical virtual‑environment for all three packaging jobs (`main.py`, `drag.py`, `drag_en.py`)**. Mismatched dependency versions will produce runtime incompatibility.
> 5. After build, put `drag.exe` and `drag_en.exe` **flat at same directory level as main.exe**. Do NOT nest them inside sub‑folders.

```bash
# Build main program (main.py)
python -m PyInstaller -D --name OCR2Word --contents-directory "." --collect-all tkinterdnd2 --collect-all customtkinter --collect-all paddle --collect-all paddleocr --collect-all cv2 --hidden-import numpy --hidden-import scipy --hidden-import shapely --hidden-import pyclipper --hidden-import RapidFuzz --hidden-import lmdb --hidden-import fitz --hidden-import skimage --hidden-import PIL --hidden-import docx --hidden-import protobuf --hidden-import requests --hidden-import tqdm --hidden-import fire main.py

# Build Chinese drag popup (drag.py)
python -m PyInstaller -D --name drag--collect-all tkinterdnd2 drag.py

# Build English drag popup (drag_en.py)
python -m PyInstaller -D --name drag --collect-all tkinterdnd2 drag_en.py
```

After packaging:
1. Locate output files from each generated `dist/` folder.
2. Gather `main.exe`, `drag.exe`, `drag_en.exe` and place them into one flat output directory.
3. Keep the empty marker file `en` inside this folder.
4. `models/` and `logs/` directories will be auto‑created on first program launch.

> Production release tip: Remove `-d` flag for final release builds to strip debug symbols.

## Runtime Notes
1. **First‑time startup**: Network connection required, application will auto‑download OCR inference model weights.
2. **≤ 8 GB RAM**: Fast‑model mode uses English‑oriented single model. Chinese OCR result quality is limited. Layout analysis and table recognition functions are disabled.
3. **> 8 GB RAM**: Pro‑model mode is activated. Layout analysis, table parsing and Chinese‑English dual‑model arbitration become available.
4. On Windows, table recognition function may be automatically disabled when user home path contains special / non‑ASCII characters, to prevent underlying PPStructure C‑level crash. Text OCR feature keeps working normally.
5. Drag‑drop popup window depends on `tkinterdnd2`. Linux users need to install system package `tkdnd`.
6. `drag.exe` / `drag_en.exe` must reside same folder as main executable, otherwise main GUI cannot launch drag‑and‑drop window.

## Language switching
Use dropdown menu at top‑right corner inside GUI. Program will create / delete empty marker file (`en` or `zh`) under project root folder to remember your language choice.

## Log mechanism
Application logs store in `logs/` folder. If program directory is read‑only access denied, logs will fallback to system temporary directory.
Windows fatal crash will pop‑up Notepad with crash log snapshot.

## Known limitations
- PaddleOCR 2.x keeps Chinese and English model weights separate, no native single mixed‑language model. Mixed‑language capability is implemented by custom dual‑model arbitration logic.
- Table extraction quality depends on layout model. Fallback secondary OCR will run when table structure parsing fails.
- GPU inference is not enabled; all computations run on CPU.

## License
This project is licensed under **Apache‑2.0 License**, see the [`LICENSE`](LICENSE) file.
