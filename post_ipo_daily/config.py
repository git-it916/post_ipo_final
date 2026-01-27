"""
Post IPO Screener - Configuration
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
        return self.BASE_DIR / "output"

    @property
    def LOG_DIR(self) -> Path:
        return self.BASE_DIR / "logs"

    @property
    def IPO_FILE(self) -> Path:
        """최초상장일.xlsx 파일 경로"""
        return self.PARENT_DIR / "최초상장일.xlsx"

    # ==========================================================================
    # IPO 필터 설정
    # ==========================================================================
    IPO_DAYS_LIMIT: int = 730  # 2년 = 730일

    # ==========================================================================
    # 출력 설정
    # ==========================================================================
    TOP_N_RESULTS: int = 20

    # ==========================================================================
    # 포맷 설정
    # ==========================================================================
    DATE_FORMAT: str = "%Y-%m-%d"
    LOG_FORMAT: str = "%(asctime)s - %(levelname)s - %(message)s"
    LOG_DATE_FORMAT: str = "%Y-%m-%d %H:%M:%S"

    def ensure_directories(self) -> None:
        """필요한 디렉토리 생성"""
        for directory in [self.OUTPUT_DIR, self.LOG_DIR]:
            directory.mkdir(parents=True, exist_ok=True)

    def get_output_filename(self, prefix: str) -> str:
        """출력 파일명 생성"""
        today = datetime.now().strftime("%Y%m%d")
        return f"{prefix}_{today}.xlsx"
