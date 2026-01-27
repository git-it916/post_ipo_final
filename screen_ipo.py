#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
IPO Screener
============
2년 이내 신규상장 종목 스크리닝

Bloomberg API를 통해 가격, 거래량, RSI, 변동성 데이터 수집
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
from dataclasses import dataclass
import re


@dataclass
class ScreenerConfig:
    """스크리너 설정"""
    IPO_DAYS_LIMIT: int = 730  # 2년
    TOP_N_RESULTS: int = 20
    BATCH_SIZE: int = 50


class IPOScreener:
    """IPO 종목 스크리너"""

    EXCLUDE_KEYWORDS = [
        'KODEX', 'TIGER', 'ACE', 'RISE', 'SOL', 'KBSTAR', 'HANARO',
        'ARIRANG', 'KOSEF', 'PLUS', 'KoAct', 'WON', 'ITF', 'TREX',
        'ETN', '스팩', 'SPAC', '호스팩', '기업인수',
        '리츠', 'REIT', 'REITs',
        'TIME', 'TRUSTONE', 'KIWOOM', 'UNICORN',
        'DAISHIN', 'BNK', '액티브', '밸류업'
    ]

    def __init__(self, config: ScreenerConfig = None, file_path: str = None):
        self.config = config or ScreenerConfig()

        # 기본 파일 경로
        if file_path is None:
            file_path = r"C:\Users\Bloomberg\Documents\ssh_project\최초상장일.xlsx"

        self.file_path = Path(file_path)
        self.output_dir = Path(r"C:\Users\Bloomberg\Documents\ssh_project\post_ipo\output")
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self._ipo_df = None
        self._bloomberg_df = None
        self._result_df = None

    def load_ipo_data(self) -> pd.DataFrame:
        """IPO 데이터 로드"""
        print("\n[Step 1] IPO 데이터 로드")
        print("=" * 50)

        df = pd.read_excel(self.file_path, skiprows=5)
        df.columns = ['코드', '코드명', '최초상장일', '상장일']

        df['최초상장일_dt'] = pd.to_datetime(
            df['최초상장일'].fillna(0).astype(int).astype(str),
            format='%Y%m%d',
            errors='coerce'
        )

        # 필터링
        cutoff = datetime.now() - timedelta(days=self.config.IPO_DAYS_LIMIT)
        recent = df[df['최초상장일_dt'] >= cutoff].copy()
        recent = recent.dropna(subset=['코드'])

        # 정규 종목코드만
        def is_regular_code(code):
            return bool(re.match(r'^A\d{6}$', str(code)))

        regular = recent[recent['코드'].apply(is_regular_code)].copy()

        # ETF/스팩 등 제외
        def is_regular_stock(name):
            if not isinstance(name, str):
                return True
            name_upper = name.upper()
            for kw in self.EXCLUDE_KEYWORDS:
                if kw.upper() in name_upper:
                    return False
            return True

        stocks = regular[regular['코드명'].apply(is_regular_stock)].copy()

        stocks['상장일'] = stocks['최초상장일_dt'].dt.strftime('%Y-%m-%d')
        stocks['days_since_ipo'] = (datetime.now() - stocks['최초상장일_dt']).dt.days
        stocks['ticker_ks'] = stocks['코드'].str[1:] + ' KS Equity'
        stocks['ticker_kq'] = stocks['코드'].str[1:] + ' KQ Equity'
        stocks = stocks.sort_values('최초상장일_dt', ascending=False)
        stocks = stocks.reset_index(drop=True)

        self._ipo_df = stocks
        print(f"대상 종목: {len(stocks)}개")

        return stocks

    def fetch_bloomberg_data(self) -> pd.DataFrame:
        """Bloomberg 데이터 수집"""
        from xbbg import blp

        print("\n[Step 2] Bloomberg 데이터 수집")
        print("=" * 50)

        if self._ipo_df is None:
            print("오류: IPO 데이터를 먼저 로드하세요.")
            return pd.DataFrame()

        tickers = self._ipo_df['ticker_ks'].tolist()

        fields = [
            'PX_LAST',
            'PX_VOLUME',
            'VOLUME_AVG_20D',
            'VOLATILITY_30D',
            'RSI_14D',
            'CHG_PCT_1D',
            'CHG_PCT_20D',
        ]

        print(f"대상: {len(tickers)}개 종목")

        results = []
        batch_size = self.config.BATCH_SIZE

        for i in range(0, len(tickers), batch_size):
            batch = tickers[i:i + batch_size]
            batch_num = i // batch_size + 1
            total_batches = (len(tickers) + batch_size - 1) // batch_size
            print(f"  배치 {batch_num}/{total_batches} 처리 중...")

            try:
                data = blp.bdp(batch, fields)
                if not data.empty:
                    results.append(data)
            except Exception as e:
                print(f"  배치 {batch_num} 오류: {e}")

        if not results:
            print("데이터 수집 실패")
            return pd.DataFrame()

        df = pd.concat(results)

        # KQ로 재시도
        missing = [t for t in tickers if t not in df.index]
        if missing:
            print(f"\nKOSDAQ 티커로 재시도: {len(missing)}개")
            kq_tickers = [t.replace(' KS ', ' KQ ') for t in missing]
            try:
                kq_data = blp.bdp(kq_tickers, fields)
                if not kq_data.empty:
                    df = pd.concat([df, kq_data])
            except Exception as e:
                print(f"  KQ 재시도 오류: {e}")

        df.columns = df.columns.str.lower()

        # 거래량 변동률
        if 'px_volume' in df.columns and 'volume_avg_20d' in df.columns:
            df['volume_change_pct'] = (
                (df['px_volume'] / df['volume_avg_20d'] - 1) * 100
            ).round(1)

        self._bloomberg_df = df.reset_index().rename(columns={'index': 'ticker'})
        print(f"\nBloomberg 데이터: {len(df)}개 종목")

        return self._bloomberg_df

    def calculate_results(self) -> pd.DataFrame:
        """결과 계산"""
        print("\n[Step 3] 결과 계산")
        print("=" * 50)

        if self._ipo_df is None:
            return pd.DataFrame()

        result = self._ipo_df[['코드', '코드명', '상장일', 'days_since_ipo', 'ticker_ks']].copy()

        if self._bloomberg_df is not None and not self._bloomberg_df.empty:
            bbg = self._bloomberg_df.copy()
            bbg['코드'] = 'A' + bbg['ticker'].str.split(' ').str[0]

            bbg_cols = ['코드', 'px_last', 'px_volume', 'volume_avg_20d', 'volume_change_pct',
                        'volatility_30d', 'rsi_14d', 'chg_pct_1d', 'chg_pct_20d']
            bbg_cols = [c for c in bbg_cols if c in bbg.columns]
            result = result.merge(bbg[bbg_cols], on='코드', how='left')

        rename_map = {
            'px_last': '현재가',
            'px_volume': '거래량',
            'volume_avg_20d': '20일평균거래량',
            'volume_change_pct': '거래량변동(%)',
            'volatility_30d': '변동성(30D)',
            'rsi_14d': 'RSI(14)',
            'chg_pct_1d': '1일수익률(%)',
            'chg_pct_20d': '20일수익률(%)'
        }
        result = result.rename(columns=rename_map)

        if 'ticker_ks' in result.columns:
            result = result.drop(columns=['ticker_ks'])

        self._result_df = result
        print(f"최종 데이터: {len(result)}개 종목")

        return result

    def print_rankings(self, n: int = 10):
        """랭킹 출력"""
        if self._result_df is None:
            return

        df = self._result_df.copy()

        print("\n" + "=" * 60)
        print("              IPO 스크리닝 결과")
        print("=" * 60)

        # 1일 수익률 TOP
        if '1일수익률(%)' in df.columns:
            top = df.dropna(subset=['1일수익률(%)']).nlargest(n, '1일수익률(%)')
            if not top.empty:
                print(f"\n[1일 수익률 TOP {n}]")
                for _, row in top.iterrows():
                    print(f"  {row['코드명']:<15} {row['1일수익률(%)']:>+7.2f}%")

        # RSI 과매도
        if 'RSI(14)' in df.columns:
            oversold = df[df['RSI(14)'] < 30].sort_values('RSI(14)')
            if not oversold.empty:
                print(f"\n[RSI 과매도 (< 30)] {len(oversold)}개")
                for _, row in oversold.head(n).iterrows():
                    print(f"  {row['코드명']:<15} RSI: {row['RSI(14)']:>5.1f}")

    def save_results(self, filename: str = None) -> Path:
        """결과 저장"""
        print("\n[Step 4] 결과 저장")
        print("=" * 50)

        if self._result_df is None:
            return None

        df = self._result_df.copy()

        if filename is None:
            today = datetime.now().strftime('%Y%m%d')
            filename = f"ipo_screening_{today}.xlsx"

        filepath = self.output_dir / filename

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='전체', index=False)

            if '1일수익률(%)' in df.columns:
                top = df.dropna(subset=['1일수익률(%)']).nlargest(20, '1일수익률(%)')
                if not top.empty:
                    top.to_excel(writer, sheet_name='1D수익률TOP', index=False)

            if '20일수익률(%)' in df.columns:
                top = df.dropna(subset=['20일수익률(%)']).nlargest(20, '20일수익률(%)')
                if not top.empty:
                    top.to_excel(writer, sheet_name='20D수익률TOP', index=False)

            if '거래량변동(%)' in df.columns:
                top = df.dropna(subset=['거래량변동(%)']).nlargest(20, '거래량변동(%)')
                if not top.empty:
                    top.to_excel(writer, sheet_name='거래량변동TOP', index=False)

            if '변동성(30D)' in df.columns:
                top = df.dropna(subset=['변동성(30D)']).nlargest(20, '변동성(30D)')
                if not top.empty:
                    top.to_excel(writer, sheet_name='변동성TOP', index=False)

            if 'RSI(14)' in df.columns:
                oversold = df[df['RSI(14)'] < 30].sort_values('RSI(14)')
                if not oversold.empty:
                    oversold.to_excel(writer, sheet_name='RSI과매도', index=False)

                overbought = df[df['RSI(14)'] > 70].sort_values('RSI(14)', ascending=False)
                if not overbought.empty:
                    overbought.to_excel(writer, sheet_name='RSI과매수', index=False)

        print(f"저장 완료: {filepath}")
        return filepath

    def run(self) -> pd.DataFrame:
        """전체 실행"""
        self.load_ipo_data()
        self.fetch_bloomberg_data()
        self.calculate_results()
        self.print_rankings()
        self.save_results()
        return self._result_df


# =============================================================================
# Main Execution (직접 실행 시)
# =============================================================================
if __name__ == "__main__":
    print("run.py를 통해 실행해주세요: python run.py")
