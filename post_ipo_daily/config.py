"""
Post IPO Monitor - Configuration
설정값 관리
"""
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field


@dataclass
class Config:
    """설정 관리 클래스"""

    # ==========================================================================
    # 파일 경로
    # ==========================================================================
    BASE_DIR: Path = field(default_factory=lambda: Path(__file__).parent.parent)

    @property
    def PARENT_DIR(self) -> Path:
        """상위 디렉토리 (ssh_project/)"""
        return self.BASE_DIR.parent

    @property
    def OUTPUT_DIR(self) -> Path:
        return self.PARENT_DIR / "POST IPO 결과"

    @property
    def LOG_DIR(self) -> Path:
        return self.BASE_DIR / "logs"

    @property
    def IPO_FILE(self) -> Path:
        """최초상장일.xlsx 파일 경로"""
        return self.PARENT_DIR / "최초상장일.xlsx"

    @property
    def SUPPLY_FILE(self) -> Path:
        """수급.xlsx 파일 경로"""
        return self.PARENT_DIR / "수급.xlsx"

    # ==========================================================================
    # IPO 필터 설정
    # ==========================================================================
    IPO_DAYS_LIMIT: int = 730  # 2년 = 730일

    # 제외할 키워드 (ETF/ETN/스팩/리츠/펀드)
    EXCLUDE_KEYWORDS: tuple = (
        'KODEX', 'TIGER', 'ACE', 'RISE', 'SOL', 'KBSTAR', 'HANARO',
        'ARIRANG', 'KOSEF', 'PLUS', 'KoAct', 'WON', 'ITF', 'TREX',
        'ETN', '스팩', 'SPAC', '호스팩', '기업인수',
        '리츠', 'REIT', 'REITs',
        'TIME', 'TRUSTONE', 'KIWOOM', 'UNICORN',
        'DAISHIN', 'BNK', '액티브', '밸류업'
    )

    # ==========================================================================
    # 스코어링 가중치
    # ==========================================================================
    MOMENTUM_WEIGHT: float = 0.30  # 모멘텀 스코어 가중치
    SUPPLY_WEIGHT: float = 0.50    # 수급 스코어 가중치
    VOLUME_WEIGHT: float = 0.20    # 거래량 스코어 가중치

    # 수급 세부 가중치
    SUPPLY_DAILY_WEIGHT: float = 0.40   # 일간
    SUPPLY_5D_WEIGHT: float = 0.35      # 5일
    SUPPLY_20D_WEIGHT: float = 0.25     # 20일

    # ==========================================================================
    # 출력 설정
    # ==========================================================================
    TOP_N_RESULTS: int = 30
    BATCH_SIZE: int = 50
    EXCEL_COLUMN_WIDTH: int = 100  # 픽셀

    # ==========================================================================
    # 포맷 설정
    # ==========================================================================
    DATE_FORMAT: str = "%Y-%m-%d"
    LOG_FORMAT: str = "%(asctime)s - %(levelname)s - %(message)s"
    LOG_DATE_FORMAT: str = "%Y-%m-%d %H:%M:%S"

    # ==========================================================================
    # 한국 공휴일 (2024~2026년)
    # ==========================================================================
    KR_HOLIDAYS: frozenset = frozenset({
        # 2024년
        '2024-01-01', '2024-02-09', '2024-02-10', '2024-02-11', '2024-02-12',
        '2024-03-01', '2024-04-10', '2024-05-05', '2024-05-06', '2024-05-15',
        '2024-06-06', '2024-08-15', '2024-09-16', '2024-09-17', '2024-09-18',
        '2024-10-03', '2024-10-09', '2024-12-25',
        # 2025년
        '2025-01-01', '2025-01-28', '2025-01-29', '2025-01-30',
        '2025-03-01', '2025-03-03', '2025-05-05', '2025-05-06', '2025-06-06',
        '2025-08-15', '2025-10-03', '2025-10-06', '2025-10-07', '2025-10-08',
        '2025-10-09', '2025-12-25',
        # 2026년
        '2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18',
        '2026-03-01', '2026-03-02', '2026-05-05', '2026-05-24', '2026-06-06',
        '2026-08-15', '2026-08-17', '2026-09-24', '2026-09-25', '2026-09-26',
        '2026-10-03', '2026-10-05', '2026-10-09', '2026-12-25',
    })

    def ensure_directories(self) -> None:
        """필요한 디렉토리 생성"""
        for directory in [self.OUTPUT_DIR, self.LOG_DIR]:
            directory.mkdir(parents=True, exist_ok=True)

    def get_output_filename(self, prefix: str = "ipo_monitoring") -> str:
        """출력 파일명 생성"""
        today = datetime.now().strftime("%Y%m%d")
        return f"{prefix}_{today}.xlsx"
