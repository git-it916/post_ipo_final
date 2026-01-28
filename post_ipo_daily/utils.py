"""
Post IPO Monitor - Utilities
유틸리티 함수
"""
import logging
import sys
from datetime import datetime, timedelta

from .config import Config


def print_progress_bar(current: int, total: int, prefix: str = '', length: int = 40):
    """진행률 바 출력"""
    percent = current / total
    filled = int(length * percent)
    bar = '█' * filled + '░' * (length - filled)
    sys.stdout.write(f'\r  {prefix} [{bar}] {current}/{total}')
    sys.stdout.flush()
    if current == total:
        print()  # 완료 시 줄바꿈


def get_previous_business_day(config: Config = None) -> datetime:
    """전 영업일 계산 (주말 + 공휴일 제외)"""
    config = config or Config()
    today = datetime.now()
    prev_day = today - timedelta(days=1)

    # 주말 또는 공휴일이면 건너뛰기
    while (prev_day.weekday() >= 5 or
           prev_day.strftime('%Y-%m-%d') in config.KR_HOLIDAYS):
        prev_day -= timedelta(days=1)

    return prev_day


def setup_logging(
    config: Config = None,
    log_level: int = logging.INFO,
    log_to_file: bool = False,
) -> logging.Logger:
    """
    로깅 설정

    Args:
        config: Config 인스턴스
        log_level: 로그 레벨
        log_to_file: 파일 로깅 여부

    Returns:
        설정된 root logger
    """
    config = config or Config()
    config.ensure_directories()

    # 루트 로거 설정
    logger = logging.getLogger()
    logger.setLevel(log_level)

    # 기존 핸들러 제거
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # 포매터
    formatter = logging.Formatter(
        config.LOG_FORMAT,
        datefmt=config.LOG_DATE_FORMAT
    )

    # 콘솔 핸들러
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 파일 핸들러
    if log_to_file:
        log_filename = f"post_ipo_{datetime.now().strftime('%Y%m%d')}.log"
        log_path = config.LOG_DIR / log_filename

        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger
