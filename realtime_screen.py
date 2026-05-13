#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Post IPO 실시간 모멘텀 스크리너
================================
__post ipo univ.xlsx 유니버스 내 종목을 장중 실시간으로 모니터링.
모멘텀·수급·거래량 기반 스코어링으로 괜찮은 종목을 선별.

실행:
    python realtime_screen.py              # 기본 60초 간격
    python realtime_screen.py --interval 30  # 30초 간격
    python realtime_screen.py --top 30       # 상위 30개 표시
    python realtime_screen.py --export       # 스냅샷 Excel 저장
"""
import sys
import time
import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from post_ipo_daily import Config


# ─────────────────────────────────────────────────────────────────────────────
# 유니버스 로드
# ─────────────────────────────────────────────────────────────────────────────
def load_universe(config: Config) -> pd.DataFrame:
    """__post ipo univ.xlsx에서 종목 유니버스 로드"""
    import re

    df = pd.read_excel(config.UNIV_FILE, header=1)

    if 'Symbol' not in df.columns:
        print("오류: __post ipo univ.xlsx에 'Symbol' 컬럼이 없습니다.")
        sys.exit(1)

    df = df[['Symbol', 'Name']].copy()
    df.columns = ['코드', '코드명']
    df = df.dropna(subset=['코드'])

    # A+6자리 숫자만
    mask = df['코드'].apply(lambda x: bool(re.match(r'^A\d{6}$', str(x))))
    df = df[mask].copy()
    df['ticker_ks'] = df['코드'].str[1:] + ' KS Equity'
    df['ticker_kq'] = df['코드'].str[1:] + ' KQ Equity'
    df = df.reset_index(drop=True)

    print(f"유니버스 로드: {len(df)}개 종목")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 수급 데이터 로드 (수급.xlsx — 전일 기준)
# ─────────────────────────────────────────────────────────────────────────────
def load_supply_data(config: Config) -> pd.DataFrame:
    """수급.xlsx에서 기관/외국인 순매수 데이터 로드 (최신 날짜 기준)"""
    supply_file = config.SUPPLY_FILE
    if not supply_file.exists():
        print(f"  수급.xlsx 없음 ({supply_file}) → 수급 데이터 없이 진행")
        return pd.DataFrame()

    try:
        raw = pd.read_excel(supply_file, header=None)
        codes = raw.iloc[8, :].values
        item_codes = raw.iloc[11, :].values
        data = raw.iloc[14:, :].copy()
        data.columns = range(len(data.columns))
        data[0] = pd.to_datetime(data[0], errors='coerce')
        data = data.dropna(subset=[0])
        latest_date = data[0].max()
        latest = data[data[0] == latest_date].iloc[0]

        records = []
        i = 1
        while i < len(codes):
            code = codes[i]
            if pd.isna(code) or not str(code).startswith('A'):
                i += 1
                continue
            rec = {'코드': code}
            j = i
            while j < len(codes) and codes[j] == code:
                ic = item_codes[j]
                val = latest[j] if j < len(latest) else None
                if ic == 'CI20003020':
                    rec['기관_일간'] = val
                elif ic == 'CI20003021':
                    rec['기관_5일'] = val
                elif ic == 'CI20003022':
                    rec['기관_20일'] = val
                elif ic == 'CI20113020':
                    rec['외국인_일간'] = val
                elif ic == 'CI20113021':
                    rec['외국인_5일'] = val
                elif ic == 'CI20113022':
                    rec['외국인_20일'] = val
                j += 1
            records.append(rec)
            i = j

        supply_df = pd.DataFrame(records).drop_duplicates(subset=['코드'], keep='first')
        for col in ['기관_일간', '기관_5일', '기관_20일', '외국인_일간', '외국인_5일', '외국인_20일']:
            if col in supply_df.columns:
                supply_df[col] = pd.to_numeric(supply_df[col], errors='coerce') / 1000  # 백만원

        print(f"  수급 데이터: {len(supply_df)}개 종목 (기준: {latest_date.strftime('%Y-%m-%d')})")
        return supply_df

    except Exception as e:
        print(f"  수급 데이터 로드 오류: {e}")
        return pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# Bloomberg 실시간 스냅샷
# ─────────────────────────────────────────────────────────────────────────────
def fetch_realtime_snapshot(tickers: list, batch_size: int = 50) -> pd.DataFrame:
    """Bloomberg BDP로 현재 시점 스냅샷 수집"""
    from xbbg import blp

    fields = [
        'LAST_PRICE',           # 현재가
        'CHG_PCT_1D',           # 일간 등락률(%)
        'VOLUME',               # 당일 거래량
        'VOLUME_AVG_20D',       # 20일 평균 거래량
        'RSI_14D',              # RSI(14)
        'MOV_AVG_10D',          # 10일 이동평균
        'EQY_WEIGHTED_AVG_PX',  # VWAP
        'CUR_MKT_CAP',          # 시가총액(백만원)
        'HIGH_1D',              # 당일 고가
        'LOW_1D',               # 당일 저가
        'OPEN_PX',              # 시가
    ]

    results = []
    for i in range(0, len(tickers), batch_size):
        batch = tickers[i:i + batch_size]
        try:
            df = blp.bdp(batch, fields)
            if not df.empty:
                results.append(df)
        except Exception as e:
            print(f"  BDP 오류 (배치 {i // batch_size + 1}): {e}")

    if not results:
        return pd.DataFrame()

    return pd.concat(results)


# ─────────────────────────────────────────────────────────────────────────────
# 실시간 스코어링
# ─────────────────────────────────────────────────────────────────────────────
def calculate_realtime_scores(
    univ: pd.DataFrame,
    bbg: pd.DataFrame,
    supply: pd.DataFrame,
) -> pd.DataFrame:
    """모멘텀·수급·거래량 기반 실시간 스코어 계산"""

    # Bloomberg 데이터 정리
    bbg = bbg.copy()
    bbg.columns = bbg.columns.str.lower()
    bbg['ticker'] = bbg.index
    bbg['코드'] = 'A' + bbg['ticker'].str.split(' ').str[0]

    # 유니버스와 병합
    result = univ[['코드', '코드명']].merge(bbg, on='코드', how='inner')

    # 시가총액 필터 (1000억 미만 제외)
    if 'cur_mkt_cap' in result.columns:
        result['시가총액(억)'] = (pd.to_numeric(result['cur_mkt_cap'], errors='coerce') / 100).round(0)
        result = result[result['시가총액(억)'].fillna(0) >= 1000].copy()

    # RVOL 계산
    if 'volume' in result.columns and 'volume_avg_20d' in result.columns:
        result['rvol'] = (
            pd.to_numeric(result['volume'], errors='coerce')
            / pd.to_numeric(result['volume_avg_20d'], errors='coerce')
        ).round(2)

    # 10일선 돌파 여부
    if 'last_price' in result.columns and 'mov_avg_10d' in result.columns:
        price = pd.to_numeric(result['last_price'], errors='coerce')
        ma10 = pd.to_numeric(result['mov_avg_10d'], errors='coerce')
        result['above_ma10'] = (price > ma10).astype(int)

    # VWAP 대비 위치 (현재가 > VWAP → 매수세 우위)
    if 'last_price' in result.columns and 'eqy_weighted_avg_px' in result.columns:
        price = pd.to_numeric(result['last_price'], errors='coerce')
        vwap = pd.to_numeric(result['eqy_weighted_avg_px'], errors='coerce')
        result['above_vwap'] = (price > vwap).astype(int)
        result['vwap_dist(%)'] = ((price - vwap) / vwap * 100).round(2)

    # 수급 데이터 병합
    if not supply.empty:
        supply_cols = [c for c in ['코드', '기관_일간', '기관_5일', '기관_20일',
                                    '외국인_일간', '외국인_5일', '외국인_20일']
                       if c in supply.columns]
        result = result.merge(supply[supply_cols], on='코드', how='left')

    # ── 스코어링 ──
    def pct_rank(s, higher_better=True):
        r = s.rank(pct=True, na_option='keep')
        return (r * 100).round(1) if higher_better else ((1 - r) * 100).round(1)

    # 1. 모멘텀 (40%)
    parts, weights = [], []
    if 'chg_pct_1d' in result.columns:
        result['_m1'] = pct_rank(pd.to_numeric(result['chg_pct_1d'], errors='coerce'))
        parts.append('_m1'); weights.append(0.35)
    if 'rsi_14d' in result.columns:
        result['_m2'] = pct_rank(pd.to_numeric(result['rsi_14d'], errors='coerce'))
        parts.append('_m2'); weights.append(0.25)
    if 'above_ma10' in result.columns:
        result['_m3'] = result['above_ma10'] * 100
        parts.append('_m3'); weights.append(0.20)
    if 'above_vwap' in result.columns:
        result['_m4'] = result['above_vwap'] * 100
        parts.append('_m4'); weights.append(0.20)

    if parts:
        w_sum = sum(weights)
        result['모멘텀'] = sum(result[c] * w for c, w in zip(parts, weights))
        result['모멘텀'] = (result['모멘텀'] / w_sum).round(1)
    else:
        result['모멘텀'] = 50

    # 2. 수급 (35%)
    sup_parts, sup_w = [], []
    for col, w in [('기관_일간', 0.20), ('기관_5일', 0.15), ('기관_20일', 0.10),
                   ('외국인_일간', 0.25), ('외국인_5일', 0.18), ('외국인_20일', 0.12)]:
        if col in result.columns:
            tag = f'_s_{col}'
            result[tag] = pct_rank(pd.to_numeric(result[col], errors='coerce'))
            sup_parts.append(tag); sup_w.append(w)

    if sup_parts:
        sw = sum(sup_w)
        result['수급'] = sum(result[c] * w for c, w in zip(sup_parts, sup_w))
        result['수급'] = (result['수급'] / sw).round(1)
    else:
        result['수급'] = 50

    # 3. 거래량 (25%)
    if 'rvol' in result.columns:
        result['거래량'] = pct_rank(pd.to_numeric(result['rvol'], errors='coerce'))
    else:
        result['거래량'] = 50

    # 종합
    result['종합'] = (
        result['모멘텀'] * 0.40 +
        result['수급'] * 0.35 +
        result['거래량'] * 0.25
    ).round(1)

    # 등급
    result['등급'] = pd.cut(
        result['종합'],
        bins=[0, 30, 50, 65, 80, 100],
        labels=['F', 'D', 'C', 'B', 'A'],
        include_lowest=True,
    ).astype(str)
    result.loc[result['종합'].isna(), '등급'] = '-'

    # 임시 컬럼 정리
    temp = [c for c in result.columns if c.startswith('_')]
    result = result.drop(columns=temp, errors='ignore')

    return result


# ─────────────────────────────────────────────────────────────────────────────
# 콘솔 출력
# ─────────────────────────────────────────────────────────────────────────────
def display_dashboard(df: pd.DataFrame, top_n: int = 20, cycle: int = 0):
    """콘솔 대시보드 출력"""
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 등급 분포
    grade_dist = df['등급'].value_counts()
    a_cnt = grade_dist.get('A', 0)
    b_cnt = grade_dist.get('B', 0)
    above_50 = (df['종합'] >= 50).sum()

    # 화면 지우기
    print('\033[2J\033[H', end='')

    print('=' * 100)
    print(f'  POST IPO 실시간 모멘텀 스크리너  |  {now}  |  갱신 #{cycle}')
    print('=' * 100)
    print(f'  전체: {len(df)}개  |  A등급: {a_cnt}  |  B등급: {b_cnt}  |  50점↑: {above_50}')
    print('-' * 100)

    # 상위 종목 테이블
    show = df.nlargest(top_n, '종합').reset_index(drop=True)

    # 표시 컬럼 선택
    header = f"{'#':>3} {'종목명':<14} {'등급':>4} {'종합':>5} {'모멘텀':>6} {'수급':>5} {'거래량':>5}"
    header += f" {'현재가':>9} {'등락%':>6} {'RSI':>5} {'RVOL':>5}"

    has_supply = '기관_일간' in show.columns
    if has_supply:
        header += f" {'기관(일)':>8} {'외인(일)':>8}"
    header += f" {'VWAP%':>6} {'MA10':>4}"

    print(header)
    print('-' * 100)

    for i, (_, row) in enumerate(show.iterrows(), 1):
        grade = row.get('등급', '-')
        # 등급별 색상 (ANSI)
        g_color = {'A': '\033[93m', 'B': '\033[92m', 'C': '\033[0m',
                   'D': '\033[33m', 'F': '\033[91m'}.get(grade, '\033[0m')
        reset = '\033[0m'

        line = (
            f"{i:>3} {str(row.get('코드명', '')):<14} "
            f"{g_color}{grade:>4}{reset} "
            f"{row.get('종합', 0):>5.1f} {row.get('모멘텀', 0):>6.1f} "
            f"{row.get('수급', 0):>5.1f} {row.get('거래량', 0):>5.1f}"
        )

        price = row.get('last_price', 0)
        chg = row.get('chg_pct_1d', 0)
        rsi = row.get('rsi_14d', 0)
        rvol = row.get('rvol', 0)

        # 등락률 색상
        chg_val = float(chg) if pd.notna(chg) else 0
        chg_color = '\033[91m' if chg_val > 0 else ('\033[94m' if chg_val < 0 else '\033[0m')

        line += (
            f" {float(price) if pd.notna(price) else 0:>9,.0f}"
            f" {chg_color}{chg_val:>+5.1f}%{reset}"
            f" {float(rsi) if pd.notna(rsi) else 0:>5.1f}"
            f" {float(rvol) if pd.notna(rvol) else 0:>5.2f}"
        )

        if has_supply:
            inst = row.get('기관_일간', 0)
            frgn = row.get('외국인_일간', 0)
            inst_val = float(inst) if pd.notna(inst) else 0
            frgn_val = float(frgn) if pd.notna(frgn) else 0
            line += f" {inst_val:>8.1f} {frgn_val:>8.1f}"

        vwap_d = row.get('vwap_dist(%)', 0)
        ma10 = row.get('above_ma10', 0)
        line += f" {float(vwap_d) if pd.notna(vwap_d) else 0:>+5.1f}%"
        line += f" {'○' if ma10 == 1 else '×':>4}"

        print(line)

    print('-' * 100)

    # 모멘텀 급등 감지 (등락률 상위 5, RVOL > 1.5)
    if 'chg_pct_1d' in df.columns and 'rvol' in df.columns:
        surge = df[
            (pd.to_numeric(df['chg_pct_1d'], errors='coerce') > 2) &
            (pd.to_numeric(df['rvol'], errors='coerce') > 1.5)
        ].nlargest(5, 'chg_pct_1d')

        if not surge.empty:
            print(f"\n  ★ 모멘텀 급등 (등락>2% & RVOL>1.5x)")
            for _, r in surge.iterrows():
                print(f"    {r['코드명']:<14} {float(r['chg_pct_1d']):>+5.1f}%  "
                      f"RVOL: {float(r['rvol']):>.2f}x  RSI: {float(r.get('rsi_14d', 0)):.1f}")

    # 수급 양호 (기관+외인 동시 순매수)
    if has_supply and '외국인_일간' in df.columns:
        both_buy = df[
            (pd.to_numeric(df['기관_일간'], errors='coerce') > 0) &
            (pd.to_numeric(df['외국인_일간'], errors='coerce') > 0)
        ].nlargest(5, '종합')

        if not both_buy.empty:
            print(f"\n  ★ 기관+외인 동시 순매수 (상위 5)")
            for _, r in both_buy.iterrows():
                print(f"    {r['코드명']:<14} 기관: {float(r['기관_일간']):>+.1f}  "
                      f"외인: {float(r['외국인_일간']):>+.1f}  종합: {r['종합']:.1f}")

    print(f"\n  Ctrl+C 로 종료  |  다음 갱신 대기 중...")


# ─────────────────────────────────────────────────────────────────────────────
# Excel 스냅샷 저장
# ─────────────────────────────────────────────────────────────────────────────
def export_snapshot(df: pd.DataFrame, config: Config):
    """현재 스냅샷을 Excel로 저장"""
    now_str = datetime.now().strftime('%Y%m%d_%H%M')
    output_dir = config.OUTPUT_DIR / "realtime"
    output_dir.mkdir(parents=True, exist_ok=True)
    filepath = output_dir / f"realtime_snapshot_{now_str}.xlsx"

    # 표시용 컬럼 정리
    display_cols = [
        '코드', '코드명', '등급', '종합', '모멘텀', '수급', '거래량',
        'last_price', 'chg_pct_1d', 'rsi_14d', 'rvol', 'above_ma10', 'vwap_dist(%)',
    ]
    supply_cols = ['기관_일간', '기관_5일', '기관_20일', '외국인_일간', '외국인_5일', '외국인_20일']
    display_cols += [c for c in supply_cols if c in df.columns]
    display_cols += ['시가총액(억)']
    display_cols = [c for c in display_cols if c in df.columns]

    out = df[display_cols].sort_values('종합', ascending=False).reset_index(drop=True)

    rename = {
        'last_price': '현재가', 'chg_pct_1d': '등락률(%)', 'rsi_14d': 'RSI(14)',
        'rvol': 'RVOL(20)', 'above_ma10': '10일선돌파', 'vwap_dist(%)': 'VWAP괴리(%)',
    }
    out = out.rename(columns=rename)

    out.to_excel(filepath, index=False, sheet_name='실시간스냅샷')
    print(f"\n  스냅샷 저장: {filepath}")


# ─────────────────────────────────────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description='Post IPO 실시간 모멘텀 스크리너')
    parser.add_argument('--interval', type=int, default=60, help='갱신 주기(초), 기본 60')
    parser.add_argument('--top', type=int, default=20, help='상위 N개 표시, 기본 20')
    parser.add_argument('--export', action='store_true', help='매 갱신마다 Excel 스냅샷 저장')
    parser.add_argument('--once', action='store_true', help='1회만 실행 후 종료')
    args = parser.parse_args()

    config = Config()
    config.ensure_directories()

    print('=' * 60)
    print('  Post IPO 실시간 모멘텀 스크리너')
    print('=' * 60)

    # 유니버스 로드
    univ = load_universe(config)

    # 수급 데이터 로드 (전일 기준, 1회만)
    print("\n수급 데이터 로드 중...")
    supply = load_supply_data(config)

    # KS/KQ 티커 준비
    tickers_ks = univ['ticker_ks'].tolist()

    cycle = 0
    try:
        while True:
            cycle += 1

            # Bloomberg 스냅샷 수집
            bbg = fetch_realtime_snapshot(tickers_ks, batch_size=config.BATCH_SIZE)

            if bbg.empty:
                print(f"\n  [{datetime.now().strftime('%H:%M:%S')}] Bloomberg 데이터 수집 실패. {args.interval}초 후 재시도...")
                time.sleep(args.interval)
                continue

            # KS에서 안 잡힌 종목 KQ 재시도
            fetched_codes = set('A' + idx.split(' ')[0] for idx in bbg.index)
            missing = univ[~univ['코드'].isin(fetched_codes)]
            if not missing.empty:
                kq_tickers = missing['ticker_kq'].tolist()
                kq_bbg = fetch_realtime_snapshot(kq_tickers, batch_size=config.BATCH_SIZE)
                if not kq_bbg.empty:
                    bbg = pd.concat([bbg, kq_bbg])

            # 스코어 계산
            scored = calculate_realtime_scores(univ, bbg, supply)

            if scored.empty:
                print(f"\n  [{datetime.now().strftime('%H:%M:%S')}] 스코어 계산 결과 없음.")
                time.sleep(args.interval)
                continue

            # 대시보드 출력
            display_dashboard(scored, top_n=args.top, cycle=cycle)

            # Excel 저장
            if args.export:
                export_snapshot(scored, config)

            if args.once:
                break

            time.sleep(args.interval)

    except KeyboardInterrupt:
        print('\n\n  스크리너 종료.')


if __name__ == '__main__':
    main()
