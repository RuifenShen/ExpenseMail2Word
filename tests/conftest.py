"""pytest 配置：将 scripts 加入路径"""
import sys
from pathlib import Path

root = Path(__file__).resolve().parent.parent
scripts = root / "scripts"
if str(scripts) not in sys.path:
    sys.path.insert(0, str(scripts))
