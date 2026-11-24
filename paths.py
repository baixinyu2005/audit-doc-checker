# -*- coding: utf-8 -*-
"""统一路径：源码运行时用程序目录；打包成 exe 后，数据写 exe 旁边，只读资源在 _internal。"""
import sys
from pathlib import Path

FROZEN = bool(getattr(sys, "frozen", False))
BASE = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
RESOURCE = Path(getattr(sys, "_MEIPASS", BASE))
