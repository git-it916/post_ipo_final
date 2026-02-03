#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Post IPO Monitor
================
2년 이내 신규상장 종목의 수급/거래량/RSI/변동성 모니터링

데이터 소스:
    - 최초상장일.xlsx: IPO 종목 리스트
    - 수급.xlsx: 기관/외국인 순매수 데이터
    - Bloomberg API: 가격, 거래량, RSI, 변동성

실행:
    python run.py
"""
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import re

from post_ipo_daily import Config
from post_ipo_daily.utils import print_progress_bar, get_previous_business_day


class IPOMonitor:
    """IPO 종목 모니터링 클래스"""

    def __init__(self, config: Config = None):
        """
        Args:
            config: Config 인스턴스
        """
        self.config = config or Config()
        self.config.ensure_directories()

        # 데이터 저장
        self._ipo_df = None
        self._supply_df = None
        self._supply_date = None
        self._bloomberg_df = None
        self._result_df = None
        self._ref_date = None

    # =========================================================================
    # Step 1: IPO 종목 로드 (2년 이내, 일반 주식만)
    # =========================================================================
    def load_ipo_universe(self) -> pd.DataFrame:
        """2년 이내 IPO 종목 로드 (ETF/스팩/리츠 제외)"""
        print("\n" + "=" * 60)
        print("[Step 1] IPO Universe 로드")
        print("=" * 60)

        df = pd.read_excel(self.config.IPO_FILE, skiprows=5)
        df.columns = ['코드', '코드명', '최초상장일', '상장일']

        # datetime 변환
        df['최초상장일_dt'] = pd.to_datetime(
            df['최초상장일'].fillna(0).astype(int).astype(str),
            format='%Y%m%d',
            errors='coerce'
        )

        # 2년 이내 필터링
        from datetime import timedelta
        cutoff = datetime.now() - timedelta(days=self.config.IPO_DAYS_LIMIT)
        recent = df[df['최초상장일_dt'] >= cutoff].copy()
        recent = recent.dropna(subset=['코드'])

        # A+6자리 숫자만
        def is_regular_code(code):
            return bool(re.match(r'^A\d{6}$', str(code)))

        regular = recent[recent['코드'].apply(is_regular_code)].copy()

        # ETF/스팩/리츠 제외
        def is_regular_stock(name):
            if not isinstance(name, str):
                return True
            name_upper = name.upper()
            for kw in self.config.EXCLUDE_KEYWORDS:
                if kw.upper() in name_upper:
                    return False
            return True

        stocks = regular[regular['코드명'].apply(is_regular_stock)].copy()

        # 정리
        stocks = stocks[['코드', '코드명', '최초상장일', '최초상장일_dt']].copy()
        stocks['days_since_ipo'] = (datetime.now() - stocks['최초상장일_dt']).dt.days
        stocks['ticker_ks'] = stocks['코드'].str[1:] + ' KS Equity'
        stocks['ticker_kq'] = stocks['코드'].str[1:] + ' KQ Equity'
        stocks = stocks.sort_values('최초상장일_dt', ascending=False)
        stocks = stocks.reset_index(drop=True)

        self._ipo_df = stocks

        print(f"기준일: {cutoff.strftime('%Y-%m-%d')} 이후 상장")
        print(f"대상 종목: {len(stocks)}개")

        return stocks

    # =========================================================================
    # Step 2: 수급 데이터 로드
    # =========================================================================
    def load_supply_data(self) -> pd.DataFrame:
        """수급.xlsx에서 기관/외국인 순매수 데이터 로드"""
        print("\n" + "=" * 60)
        print("[Step 2] 수급 데이터 로드")
        print("=" * 60)

        # 전체 데이터 읽기
        raw = pd.read_excel(self.config.SUPPLY_FILE, header=None)

        # 메타데이터 추출 (행 8~13)
        codes = raw.iloc[8, :].values
        names = raw.iloc[9, :].values
        item_codes = raw.iloc[11, :].values

        # 데이터 부분 (행 14부터)
        data = raw.iloc[14:, :].copy()
        data.columns = range(len(data.columns))

        # 날짜 컬럼 (0번)
        data[0] = pd.to_datetime(data[0], errors='coerce')
        data = data.dropna(subset=[0])
        data = data.rename(columns={0: '날짜'})

        # 가장 최근 날짜 데이터만 추출
        latest_date = data['날짜'].max()
        latest = data[data['날짜'] == latest_date].iloc[0]

        self._supply_date = latest_date
        print(f"최신 데이터 날짜: {latest_date.strftime('%Y-%m-%d')}")

        # 종목별 수급 데이터 정리
        supply_records = []

        i = 1
        while i < len(codes):
            code = codes[i]
            if pd.isna(code) or not str(code).startswith('A'):
                i += 1
                continue

            record = {'코드': code, '코드명': names[i]}

            j = i
            while j < len(codes) and codes[j] == code:
                item_code = item_codes[j]
                value = latest[j] if j < len(latest) else None

                if item_code == 'CI20003020':
                    record['기관_일간'] = value
                elif item_code == 'CI20003021':
                    record['기관_5일'] = value
                elif item_code == 'CI20003022':
                    record['기관_20일'] = value
                elif item_code == 'CI20113020':
                    record['외국인_일간'] = value
                elif item_code == 'CI20113021':
                    record['외국인_5일'] = value
                elif item_code == 'CI20113022':
                    record['외국인_20일'] = value

                j += 1

            supply_records.append(record)
            i = j

        supply_df = pd.DataFrame(supply_records)

        # 숫자 컬럼 변환
        numeric_cols = ['기관_일간', '기관_5일', '기관_20일', '외국인_일간', '외국인_5일', '외국인_20일']
        for col in numeric_cols:
            if col in supply_df.columns:
                supply_df[col] = pd.to_numeric(supply_df[col], errors='coerce')

        self._supply_df = supply_df

        print(f"수급 데이터: {len(supply_df)}개 종목")

        return supply_df

    # =========================================================================
    # Step 3: Bloomberg 데이터 수집 (전일 T-1 기준)
    # =========================================================================
    def fetch_bloomberg_data(self) -> pd.DataFrame:
        """Bloomberg에서 전일(T-1) 가격, 거래량, RSI, 변동성 수집"""
        from xbbg import blp

        print("\n" + "=" * 60)
        print("[Step 3] Bloomberg 데이터 수집 (T-1)")
        print("=" * 60)

        if self._ipo_df is None:
            print("오류: IPO Universe를 먼저 로드하세요.")
            return pd.DataFrame()

        # 전일 기준
        ref_date = get_previous_business_day(self.config)
        ref_date_str = ref_date.strftime('%Y-%m-%d')
        self._ref_date = ref_date

        print(f"기준일: {ref_date_str} (T-1)")

        # 수급 데이터 날짜 검증
        if self._supply_date is not None:
            supply_date_str = self._supply_date.strftime('%Y-%m-%d')
            if supply_date_str != ref_date_str:
                print(f"  ⚠️  수급 데이터 날짜({supply_date_str})와 기준일 불일치!")
                print(f"      수급.xlsx 파일을 최신으로 업데이트하세요.")

        tickers = self._ipo_df['ticker_ks'].tolist()

        # BDH용 필드 (히스토리컬)
        bdh_fields = ['PX_LAST', 'PX_VOLUME', 'CHG_PCT_1D']

        # BDP용 필드 (기술적 지표)
        bdp_fields = [
            'VOLUME_AVG_20D',
            'VOLATILITY_30D',
            'RSI_14D',
            'CHG_PCT_20D',
            'MOV_AVG_10D',
        ]

        print(f"대상: {len(tickers)}개 종목\n")

        results = []
        batch_size = self.config.BATCH_SIZE
        total_batches = (len(tickers) + batch_size - 1) // batch_size

        # BDH: 전일 가격/거래량 데이터
        for i in range(0, len(tickers), batch_size):
            batch = tickers[i:i + batch_size]
            batch_num = i // batch_size + 1

            print_progress_bar(batch_num, total_batches, prefix='BDH 수집')

            try:
                hist_data = blp.bdh(batch, bdh_fields, ref_date_str, ref_date_str)
                if not hist_data.empty:
                    # BDH 결과: 인덱스=날짜, 컬럼=MultiIndex(ticker, field)
                    # 우리가 원하는 형태: 인덱스=ticker, 컬럼=field

                    # 멀티레벨 컬럼인 경우 (ticker, field) 구조
                    if hist_data.columns.nlevels > 1:
                        # 첫 번째 행(날짜) 데이터를 unstack하고 전치
                        hist_flat = hist_data.iloc[0].unstack(level=0).T
                    else:
                        # 단일 티커인 경우
                        hist_flat = hist_data.T
                        hist_flat.index = batch[:len(hist_flat)]

                    if not hist_flat.empty:
                        results.append(hist_flat)
            except Exception as e:
                print(f"\n  ⚠️ BDH 오류 (배치 {batch_num}): {e}")

        # BDP: 기술적 지표 수집
        bdp_results = []
        for i in range(0, len(tickers), batch_size):
            batch = tickers[i:i + batch_size]
            batch_num = i // batch_size + 1

            print_progress_bar(batch_num, total_batches, prefix='BDP 수집')

            try:
                bdp_data = blp.bdp(batch, bdp_fields)
                if not bdp_data.empty:
                    bdp_results.append(bdp_data)
            except Exception as e:
                print(f"\n  ⚠️ BDP 오류 (배치 {batch_num}): {e}")

        # 결과 병합
        if not results:
            print("BDH 데이터 수집 실패")
            return pd.DataFrame()

        df_hist = pd.concat(results)

        # 디버깅: BDH 결과 확인
        print(f"\n[DEBUG] BDH 결과:")
        print(f"  - Shape: {df_hist.shape}")
        print(f"  - Columns: {list(df_hist.columns)}")
        print(f"  - Index 샘플: {list(df_hist.index[:3])}")

        if bdp_results:
            df_bdp = pd.concat(bdp_results)

            # 디버깅: BDP 결과 확인
            print(f"\n[DEBUG] BDP 결과:")
            print(f"  - Shape: {df_bdp.shape}")
            print(f"  - Columns: {list(df_bdp.columns)}")
            print(f"  - Index 샘플: {list(df_bdp.index[:3])}")

            # 인덱스 매칭 확인
            common = set(df_hist.index) & set(df_bdp.index)
            print(f"  - 공통 인덱스 수: {len(common)}")

            df = df_hist.join(df_bdp, how='left')

            # 병합 후 확인
            print(f"\n[DEBUG] 병합 후:")
            print(f"  - Columns: {list(df.columns)}")
        else:
            print("\n[DEBUG] BDP 결과 없음!")
            df = df_hist

        # KS로 안 되는 종목은 KQ로 재시도
        missing = [t for t in tickers if t not in df.index]
        if missing:
            print(f"\nKOSDAQ 티커로 재시도: {len(missing)}개")
            kq_tickers = [t.replace(' KS ', ' KQ ') for t in missing]

            try:
                kq_hist = blp.bdh(kq_tickers, bdh_fields, ref_date_str, ref_date_str)
                kq_bdp = blp.bdp(kq_tickers, bdp_fields)

                if not kq_hist.empty:
                    # 멀티레벨 컬럼인 경우
                    if kq_hist.columns.nlevels > 1:
                        kq_flat = kq_hist.iloc[0].unstack(level=0).T
                    else:
                        kq_flat = kq_hist.T
                        kq_flat.index = kq_tickers[:len(kq_flat)]

                    if not kq_bdp.empty:
                        kq_flat = kq_flat.join(kq_bdp, how='left')
                    df = pd.concat([df, kq_flat])
            except Exception as e:
                print(f"  KQ 재시도 오류: {e}")

        # 컬럼명 소문자로
        df.columns = df.columns.str.lower()

        # RVOL(20) 계산: 당일거래량 / 20일평균거래량
        if 'px_volume' in df.columns and 'volume_avg_20d' in df.columns:
            df['rvol_20'] = (
                df['px_volume'] / df['volume_avg_20d']
            ).round(2)

        # 10일선 돌파 여부
        if 'px_last' in df.columns and 'mov_avg_10d' in df.columns:
            df['above_ma10'] = (df['px_last'] > df['mov_avg_10d']).astype(int)

        self._bloomberg_df = df.reset_index().rename(columns={'index': 'ticker'})

        print(f"\nBloomberg 데이터: {len(df)}개 종목 (기준일: {ref_date_str})")

        # Bloomberg 원본 데이터 저장
        self._save_bloomberg_raw_data(ref_date_str)

        return self._bloomberg_df

    def _save_bloomberg_raw_data(self, ref_date_str: str):
        """Bloomberg 원본 데이터를 Excel로 저장"""
        if self._bloomberg_df is None or self._bloomberg_df.empty:
            return

        # raw data 폴더 생성
        raw_data_dir = self.config.OUTPUT_DIR / "raw data"
        raw_data_dir.mkdir(parents=True, exist_ok=True)

        filename = f"bloomberg_raw_{ref_date_str.replace('-', '')}.xlsx"
        filepath = raw_data_dir / filename

        try:
            self._bloomberg_df.to_excel(filepath, index=False, sheet_name='Bloomberg_Data')
            print(f"  Bloomberg 원본 저장: {filepath}")
        except Exception as e:
            print(f"  Bloomberg 원본 저장 실패: {e}")

    # =========================================================================
    # Step 4: 데이터 병합 및 최종 결과 생성
    # =========================================================================
    def merge_data(self) -> pd.DataFrame:
        """IPO + 수급 + Bloomberg 데이터 병합"""
        print("\n" + "=" * 60)
        print("[Step 4] 데이터 병합")
        print("=" * 60)

        if self._ipo_df is None:
            print("오류: IPO Universe가 없습니다.")
            return pd.DataFrame()

        # 기본 IPO 데이터
        result = self._ipo_df[['코드', '코드명', '최초상장일_dt', 'days_since_ipo', 'ticker_ks']].copy()
        result = result.rename(columns={'최초상장일_dt': '상장일'})

        # 수급 데이터 병합
        if self._supply_df is not None and not self._supply_df.empty:
            supply_cols = ['코드', '기관_일간', '기관_5일', '기관_20일', '외국인_일간', '외국인_5일', '외국인_20일']
            supply_cols = [c for c in supply_cols if c in self._supply_df.columns]
            result = result.merge(self._supply_df[supply_cols], on='코드', how='left')

        # Bloomberg 데이터 병합
        if self._bloomberg_df is not None and not self._bloomberg_df.empty:
            bbg = self._bloomberg_df.copy()
            bbg['코드'] = 'A' + bbg['ticker'].str.split(' ').str[0]

            bbg_cols = ['코드', 'px_last', 'px_volume', 'volume_avg_20d', 'rvol_20',
                        'volatility_30d', 'rsi_14d', 'chg_pct_1d', 'chg_pct_20d',
                        'mov_avg_10d', 'above_ma10']
            bbg_cols = [c for c in bbg_cols if c in bbg.columns]

            result = result.merge(bbg[bbg_cols], on='코드', how='left')

        # 컬럼명 정리
        rename_map = {
            'px_last': '현재가',
            'px_volume': '거래량',
            'volume_avg_20d': '20일평균거래량',
            'rvol_20': 'RVOL(20)',
            'volatility_30d': '변동성(30D)',
            'rsi_14d': 'RSI(14)',
            'chg_pct_1d': '1일수익률(%)',
            'chg_pct_20d': '20일수익률(%)',
            'mov_avg_10d': '10일이평',
            'above_ma10': '10일선돌파'
        }
        result = result.rename(columns=rename_map)

        # 상장일 형식 변경
        if '상장일' in result.columns:
            result['상장일'] = pd.to_datetime(result['상장일']).dt.strftime('%Y-%m-%d')

        # 단위 변환 (백만원)
        for col in ['기관_일간', '기관_5일', '기관_20일', '외국인_일간', '외국인_5일', '외국인_20일']:
            if col in result.columns:
                result[col] = (result[col] / 1000).round(1)

        # ticker_ks 컬럼 제거
        if 'ticker_ks' in result.columns:
            result = result.drop(columns=['ticker_ks'])

        # 스코어 계산
        print("스코어 계산 중...")
        result = self.calculate_scores(result)

        self._result_df = result

        print(f"최종 데이터: {len(result)}개 종목")

        # 등급별 분포 출력
        if '등급' in result.columns:
            grade_dist = result['등급'].value_counts().sort_index(ascending=False)
            print(f"등급 분포: {dict(grade_dist)}")

        return result

    # =========================================================================
    # Step 5: 스코어링 시스템
    # =========================================================================
    def _normalize_score(self, series: pd.Series, higher_is_better: bool = True) -> pd.Series:
        """0~100 점수로 정규화 (백분위 기반)"""
        if series.isna().all():
            return pd.Series([50] * len(series), index=series.index)

        pct_rank = series.rank(pct=True, na_option='keep')

        if higher_is_better:
            return (pct_rank * 100).round(1)
        else:
            return ((1 - pct_rank) * 100).round(1)

    def calculate_scores(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        스코어링 시스템
        - 모멘텀 스코어 (30%): 20일수익률 + 10일선돌파 + RSI
        - 수급 스코어 (50%): 기관 + 외인 (일간/5일/20일)
        - 거래량 스코어 (20%): RVOL(20)
        """
        result = df.copy()

        # ===== 1. 모멘텀 스코어 =====
        momentum_components = []

        if '20일수익률(%)' in result.columns:
            result['_m1'] = self._normalize_score(result['20일수익률(%)'], higher_is_better=True)
            momentum_components.append('_m1')

        if '10일선돌파' in result.columns:
            result['_m2'] = result['10일선돌파'].fillna(0) * 100
            momentum_components.append('_m2')

        if 'RSI(14)' in result.columns:
            rsi = result['RSI(14)'].fillna(50)
            result['_m3'] = np.where(
                rsi < 30, 20,
                np.where(rsi > 70, 30,
                         np.where(rsi >= 50, 80, 60))
            )
            momentum_components.append('_m3')

        if momentum_components:
            result['모멘텀스코어'] = result[momentum_components].mean(axis=1).round(1)
        else:
            result['모멘텀스코어'] = 50

        # ===== 2. 수급 스코어 =====
        supply_components = []

        # 기관
        if '기관_일간' in result.columns:
            result['_s1'] = self._normalize_score(result['기관_일간'], higher_is_better=True)
            supply_components.append(('_s1', self.config.SUPPLY_DAILY_WEIGHT))
        if '기관_5일' in result.columns:
            result['_s2'] = self._normalize_score(result['기관_5일'], higher_is_better=True)
            supply_components.append(('_s2', self.config.SUPPLY_5D_WEIGHT))
        if '기관_20일' in result.columns:
            result['_s3'] = self._normalize_score(result['기관_20일'], higher_is_better=True)
            supply_components.append(('_s3', self.config.SUPPLY_20D_WEIGHT))

        # 외국인
        if '외국인_일간' in result.columns:
            result['_s4'] = self._normalize_score(result['외국인_일간'], higher_is_better=True)
            supply_components.append(('_s4', self.config.SUPPLY_DAILY_WEIGHT))
        if '외국인_5일' in result.columns:
            result['_s5'] = self._normalize_score(result['외국인_5일'], higher_is_better=True)
            supply_components.append(('_s5', self.config.SUPPLY_5D_WEIGHT))
        if '외국인_20일' in result.columns:
            result['_s6'] = self._normalize_score(result['외국인_20일'], higher_is_better=True)
            supply_components.append(('_s6', self.config.SUPPLY_20D_WEIGHT))

        if supply_components:
            inst_cols = [c for c, _ in supply_components if c in ['_s1', '_s2', '_s3']]
            frgn_cols = [c for c, _ in supply_components if c in ['_s4', '_s5', '_s6']]

            inst_score = 50
            frgn_score = 50

            if inst_cols:
                inst_weights = [self.config.SUPPLY_DAILY_WEIGHT, self.config.SUPPLY_5D_WEIGHT, self.config.SUPPLY_20D_WEIGHT][:len(inst_cols)]
                inst_score = sum(result[c] * w for c, w in zip(inst_cols, inst_weights)) / sum(inst_weights)

            if frgn_cols:
                frgn_weights = [self.config.SUPPLY_DAILY_WEIGHT, self.config.SUPPLY_5D_WEIGHT, self.config.SUPPLY_20D_WEIGHT][:len(frgn_cols)]
                frgn_score = sum(result[c] * w for c, w in zip(frgn_cols, frgn_weights)) / sum(frgn_weights)

            result['수급스코어'] = ((inst_score + frgn_score) / 2).round(1)
        else:
            result['수급스코어'] = 50

        # ===== 3. 거래량 스코어 =====
        if 'RVOL(20)' in result.columns:
            result['거래량스코어'] = self._normalize_score(result['RVOL(20)'], higher_is_better=True)
        else:
            result['거래량스코어'] = 50

        # ===== 4. 종합 스코어 =====
        result['종합스코어'] = (
            result['모멘텀스코어'] * self.config.MOMENTUM_WEIGHT +
            result['수급스코어'] * self.config.SUPPLY_WEIGHT +
            result['거래량스코어'] * self.config.VOLUME_WEIGHT
        ).round(1)

        # 등급 부여
        grade = pd.cut(
            result['종합스코어'],
            bins=[0, 30, 50, 65, 80, 100],
            labels=['F', 'D', 'C', 'B', 'A'],
            include_lowest=True
        )
        # Categorical을 문자열로 변환 후 NaN 처리
        result['등급'] = grade.astype(str)
        result.loc[result['종합스코어'].isna(), '등급'] = '-'

        # 임시 컬럼 삭제
        temp_cols = [c for c in result.columns if c.startswith('_')]
        result = result.drop(columns=temp_cols)

        return result

    def _create_recommendation(self, df: pd.DataFrame) -> pd.DataFrame:
        """종합스코어 기반 추천종목 선정"""
        if '종합스코어' not in df.columns:
            return None

        top = df.nlargest(self.config.TOP_N_RESULTS, '종합스코어').copy()
        top['순위'] = range(1, len(top) + 1)

        display_cols = [
            '순위', '등급', '코드', '코드명', '상장일', '현재가',
            '종합스코어', '모멘텀스코어', '수급스코어', '거래량스코어',
            '20일수익률(%)', '10일선돌파', 'RSI(14)',
            '기관_일간', '외국인_일간', 'RVOL(20)', '변동성(30D)'
        ]
        display_cols = [c for c in display_cols if c in top.columns]

        return top[display_cols].reset_index(drop=True)

    # =========================================================================
    # Step 6: 결과 저장 및 출력
    # =========================================================================
    def _set_column_width(self, writer, sheet_name: str):
        """Excel 시트의 모든 컬럼 폭을 설정"""
        from openpyxl.utils import get_column_letter

        excel_width = self.config.EXCEL_COLUMN_WIDTH / 7

        ws = writer.sheets[sheet_name]
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = excel_width

    def _load_rsi_history_from_excel(self) -> pd.DataFrame:
        """이전 엑셀 파일에서 RSI 히스토리 로드"""
        import glob

        pattern = str(self.config.OUTPUT_DIR / "ipo_monitoring_*.xlsx")
        files = glob.glob(pattern)

        if not files:
            return pd.DataFrame(columns=['날짜', '분석종목수', 'RSI65이상', '과매수비율(%)', '전일대비'])

        latest_file = max(files)

        try:
            df = pd.read_excel(latest_file, sheet_name='RSI추이')
            df['날짜'] = pd.to_datetime(df['날짜'])
            return df
        except (ValueError, KeyError):
            return pd.DataFrame(columns=['날짜', '분석종목수', 'RSI65이상', '과매수비율(%)', '전일대비'])

    def _calculate_rsi_history(self) -> pd.DataFrame:
        """RSI 65 이상 종목수 히스토리 (평소: 당일만, 초기: 30일 채움)"""
        from xbbg import blp

        if self._ipo_df is None:
            return pd.DataFrame(columns=['날짜', '분석종목수', 'RSI65이상', '과매수비율(%)', '전일대비'])

        tickers = self._ipo_df['ticker_ks'].tolist()
        end_date = self._ref_date if self._ref_date else datetime.now()
        end_str = end_date.strftime('%Y-%m-%d')

        # 1. 기존 히스토리 로드
        existing = self._load_rsi_history_from_excel()

        # 2. 초기 실행 여부 확인 (기존 데이터가 5개 미만이면 초기로 간주)
        is_initial = len(existing) < 5

        if is_initial:
            # 초기: 30일 히스토리 채우기
            print(f"\nRSI 히스토리 초기화 ({self.config.RSI_CALC_DAYS}일 계산)...")
            start_date = end_date - pd.Timedelta(days=45)
            start_str = start_date.strftime('%Y-%m-%d')
        else:
            # 평소: 당일만 계산
            print(f"\nRSI 히스토리 업데이트 (당일)...")
            start_str = end_str

        try:
            rsi_data = blp.bdh(tickers, 'RSI_14D', start_str, end_str)

            if rsi_data.empty:
                print("  RSI 데이터 없음")
                return existing if not existing.empty else pd.DataFrame(columns=['날짜', '분석종목수', 'RSI65이상', '과매수비율(%)', '전일대비'])

            # 멀티인덱스 처리
            if rsi_data.columns.nlevels > 1:
                rsi_data.columns = rsi_data.columns.droplevel(1)

            # 각 날짜별 RSI 65 이상 종목수 계산
            calc_results = []
            for date in rsi_data.index:
                row_data = rsi_data.loc[date]
                valid_count = row_data.notna().sum()  # RSI 데이터가 있는 종목 수
                overbought_count = (row_data >= self.config.RSI_THRESHOLD).sum()
                ratio = (overbought_count / valid_count * 100) if valid_count > 0 else 0
                calc_results.append({
                    '날짜': pd.to_datetime(date),
                    '분석종목수': int(valid_count),
                    'RSI65이상': int(overbought_count),
                    '과매수비율(%)': round(ratio, 1)
                })

            calculated = pd.DataFrame(calc_results)

            if is_initial:
                # 초기: 30일만 유지
                calculated = calculated.tail(self.config.RSI_CALC_DAYS)
                print(f"  초기 {len(calculated)}일 계산 완료")
                history = calculated
            else:
                # 평소: 기존 + 당일 병합
                print(f"  당일 데이터 추가")
                # 당일 날짜가 이미 있으면 제거
                today_str = end_str
                existing = existing[existing['날짜'].dt.strftime('%Y-%m-%d') != today_str]
                history = pd.concat([existing, calculated], ignore_index=True)

            # 날짜순 정렬
            history = history.sort_values('날짜').reset_index(drop=True)

            # 90일 롤링 유지
            cutoff_date = end_date - pd.Timedelta(days=self.config.RSI_HISTORY_DAYS)
            history = history[history['날짜'] >= cutoff_date]

            # 전일대비 계산
            history['전일대비'] = history['RSI65이상'].diff().fillna(0).astype(int)

            print(f"  RSI 히스토리 총 {len(history)}일")

            return history

        except Exception as e:
            print(f"  RSI 히스토리 계산 오류: {e}")
            return existing if not existing.empty else pd.DataFrame(columns=['날짜', '분석종목수', 'RSI65이상', '과매수비율(%)', '전일대비'])

    def save_results(self, filename: str = None) -> Path:
        """결과를 Excel로 저장"""
        print("\n" + "=" * 60)
        print("[Step 6] 결과 저장")
        print("=" * 60)

        if self._result_df is None or self._result_df.empty:
            print("저장할 데이터가 없습니다.")
            return None

        df = self._result_df.copy()

        if filename is None:
            filename = self.config.get_output_filename()

        filepath = self.config.OUTPUT_DIR / filename

        sheet_names = []

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # 추천종목
            recommend_df = self._create_recommendation(df)
            if recommend_df is not None and not recommend_df.empty:
                recommend_df.to_excel(writer, sheet_name='추천종목', index=False)
                sheet_names.append('추천종목')

            # A등급
            if '등급' in df.columns:
                grade_a = df[df['등급'] == 'A'].sort_values('종합스코어', ascending=False)
                if not grade_a.empty:
                    grade_a.to_excel(writer, sheet_name='A등급', index=False)
                    sheet_names.append('A등급')

            # 전체
            df.to_excel(writer, sheet_name='전체', index=False)
            sheet_names.append('전체')

            # RSI 과매도/과매수
            if 'RSI(14)' in df.columns:
                oversold = df[df['RSI(14)'] < 30].sort_values('RSI(14)')
                if not oversold.empty:
                    oversold.to_excel(writer, sheet_name='RSI과매도', index=False)
                    sheet_names.append('RSI과매도')

                overbought = df[df['RSI(14)'] > 70].sort_values('RSI(14)', ascending=False)
                if not overbought.empty:
                    overbought.to_excel(writer, sheet_name='RSI과매수', index=False)
                    sheet_names.append('RSI과매수')

            # 외국인/기관 순매수 TOP
            if '외국인_일간' in df.columns:
                foreign_top = df.dropna(subset=['외국인_일간']).nlargest(20, '외국인_일간')
                if not foreign_top.empty:
                    foreign_top.to_excel(writer, sheet_name='외국인순매수TOP', index=False)
                    sheet_names.append('외국인순매수TOP')

            if '기관_일간' in df.columns:
                inst_top = df.dropna(subset=['기관_일간']).nlargest(20, '기관_일간')
                if not inst_top.empty:
                    inst_top.to_excel(writer, sheet_name='기관순매수TOP', index=False)
                    sheet_names.append('기관순매수TOP')

            # RVOL TOP
            if 'RVOL(20)' in df.columns:
                vol_top = df.dropna(subset=['RVOL(20)']).nlargest(20, 'RVOL(20)')
                if not vol_top.empty:
                    vol_top.to_excel(writer, sheet_name='RVOL_TOP', index=False)
                    sheet_names.append('RVOL_TOP')

            # RSI 히스토리 (RSI 65 이상 종목수 일별 추이, 30일)
            rsi_history = self._calculate_rsi_history()
            if not rsi_history.empty:
                rsi_display = rsi_history.copy()
                rsi_display['날짜'] = rsi_display['날짜'].dt.strftime('%Y-%m-%d')
                rsi_display = rsi_display.sort_values('날짜', ascending=False)
                rsi_display.to_excel(writer, sheet_name='RSI추이', index=False)
                sheet_names.append('RSI추이')

            # 컬럼 폭 적용
            for sheet in sheet_names:
                self._set_column_width(writer, sheet)

        ref_date_str = self._ref_date.strftime('%Y-%m-%d') if self._ref_date else '(미정)'
        print(f"저장 완료: {filepath}")
        print(f"  (데이터 기준일: {ref_date_str})")

        return filepath

    def print_summary(self, n: int = 10):
        """요약 출력"""
        if self._result_df is None or self._result_df.empty:
            print("데이터가 없습니다.")
            return

        df = self._result_df.copy()

        print("\n" + "=" * 80)
        print("                    IPO 모니터링 요약")
        print("=" * 80)
        ref_date_str = self._ref_date.strftime('%Y-%m-%d') if self._ref_date else '(미정)'
        print(f"데이터 기준일: {ref_date_str} (T-1)")
        print(f"실행일: {datetime.now().strftime('%Y-%m-%d')}")
        print(f"총 종목 수: {len(df)}개")

        if '등급' in df.columns:
            grade_dist = df['등급'].value_counts().sort_index(ascending=False)
            print(f"\n등급 분포: {dict(grade_dist)}")

        # 추천종목 TOP
        recommend = self._create_recommendation(df)
        if recommend is not None and not recommend.empty:
            print(f"\n[★ 추천종목 TOP {min(n, len(recommend))}] (종합스코어 기준)")
            print("-" * 80)
            print(f"  {'순위':>4} {'등급':>4} {'종목명':<12} {'종합':>6} {'모멘텀':>6} {'수급':>6} {'거래량':>6}")
            print("-" * 80)
            for _, row in recommend.head(n).iterrows():
                grade = row.get('등급', '-')
                print(f"  {int(row['순위']):>4} {grade:>4} {row['코드명']:<12} "
                      f"{row['종합스코어']:>6.1f} {row.get('모멘텀스코어', 0):>6.1f} "
                      f"{row.get('수급스코어', 0):>6.1f} {row.get('거래량스코어', 0):>6.1f}")

        # RSI 과매도
        if 'RSI(14)' in df.columns:
            oversold = df[df['RSI(14)'] < 30].sort_values('RSI(14)')
            if not oversold.empty:
                print(f"\n[RSI 과매도 (< 30)] {len(oversold)}개")
                print("-" * 70)
                for _, row in oversold.head(n).iterrows():
                    print(f"  {row['코드명']:<15} RSI: {row['RSI(14)']:>5.1f}  변동성: {row.get('변동성(30D)', 0):>5.1f}%")

        # 외국인 순매수 TOP
        if '외국인_일간' in df.columns:
            foreign_top = df.dropna(subset=['외국인_일간']).nlargest(n, '외국인_일간')
            if not foreign_top.empty:
                print(f"\n[외국인 순매수 TOP {n}] (단위: 백만원)")
                print("-" * 70)
                for _, row in foreign_top.iterrows():
                    print(f"  {row['코드명']:<15} 일간: {row['외국인_일간']:>8.1f}  5일: {row.get('외국인_5일', 0):>8.1f}  20일: {row.get('외국인_20일', 0):>8.1f}")

        # 기관 순매수 TOP
        if '기관_일간' in df.columns:
            inst_top = df.dropna(subset=['기관_일간']).nlargest(n, '기관_일간')
            if not inst_top.empty:
                print(f"\n[기관 순매수 TOP {n}] (단위: 백만원)")
                print("-" * 70)
                for _, row in inst_top.iterrows():
                    print(f"  {row['코드명']:<15} 일간: {row['기관_일간']:>8.1f}  5일: {row.get('기관_5일', 0):>8.1f}  20일: {row.get('기관_20일', 0):>8.1f}")

        # RVOL TOP
        if 'RVOL(20)' in df.columns:
            vol_top = df.dropna(subset=['RVOL(20)']).nlargest(n, 'RVOL(20)')
            if not vol_top.empty:
                print(f"\n[RVOL(20) TOP {n}]")
                print("-" * 70)
                for _, row in vol_top.iterrows():
                    print(f"  {row['코드명']:<15} RVOL: {row['RVOL(20)']:>5.2f}x  RSI: {row.get('RSI(14)', 0):>5.1f}")

        print("=" * 80)

    # =========================================================================
    # 전체 실행
    # =========================================================================
    def run(self) -> pd.DataFrame:
        """전체 모니터링 프로세스 실행"""
        self.load_ipo_universe()
        self.load_supply_data()
        self.fetch_bloomberg_data()
        self.merge_data()
        self.print_summary()
        self.save_results()
        return self._result_df


# =============================================================================
# Main
# =============================================================================
if __name__ == "__main__":
    print("run.py를 통해 실행해주세요: python run.py")
