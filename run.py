#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Post IPO Screener - Main Entry Point
=====================================
2년 이내 신규상장 종목 스크리닝

실행:
    python run.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from screen_ipo import IPOScreener, ScreenerConfig
from post_ipo_daily import setup_logging


def main():
    """메인 실행 함수"""
    setup_logging()

    print("=" * 60)
    print("       Post IPO Screener")
    print("       (2년 이내 신규 상장주)")
    print("=" * 60)

    config = ScreenerConfig(
        IPO_DAYS_LIMIT=730,
        TOP_N_RESULTS=20,
    )

    screener = IPOScreener(config)
    result = screener.run()

    if result is not None and not result.empty:
        print(f"\n총 {len(result)}개 종목 스크리닝 완료")
    else:
        print("\n스크리닝 결과가 없습니다.")


if __name__ == "__main__":
    main()
