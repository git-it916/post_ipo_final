#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Post IPO Monitor - Main Entry Point
====================================
2년 이내 신규상장 종목 모니터링

실행:
    python run.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from screen_ipo import IPOMonitor
from post_ipo_daily import Config, setup_logging


def main():
    """메인 실행 함수"""
    setup_logging()

    print("=" * 60)
    print("Post IPO Monitor")
    print("2년 이내 신규 상장주)")
    print("=" * 60)

    config = Config()
    monitor = IPOMonitor(config)
    result = monitor.run()

    if result is not None and not result.empty:
        print(f"\n모니터링 완료: {len(result)}개 종목")
    else:
        print("\n모니터링 결과가 없습니다.")


if __name__ == "__main__":
    main()
