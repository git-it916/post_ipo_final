"""
Post IPO Screener Package
=========================
2년 이내 신규상장 종목 스크리닝

데이터 소스:
    - Bloomberg API (xbbg): 가격, 거래량, RSI, 변동성
"""

__version__ = "2.0.0"

from .config import Config
from .utils import setup_logging

__all__ = [
    "Config",
    "setup_logging",
]
