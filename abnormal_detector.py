"""
abnormal_detector.py

Usage:
    python abnormal_detector.py

- 엑셀 파일은 같은 폴더에 abnormalTransactionDetection.xlsx 로 둔다.
- 결과:
    - suspicious_groups.csv : 의심 그룹 리스트와 점수
    - suspicious_profits.csv : 의심 계정별 수익 요약
    - console 출력으로 상위 의심 그룹 요약 표시
"""

import os
import pandas as pd
import numpy as np
import networkx as nx
from collections import defaultdict
def load_sheets(file_name):
    """
    Excel 파일에서 시트별로 데이터를 불러옴
    Trade, Funding, Reward, IP, Spec 시트를 가져옴
    """
    if not os.path.exists(file_name):
        raise FileNotFoundError(f"{file_name} not found in cwd={os.getcwd()}")
    sheets = pd.read_excel(file_name, sheet_name=None)
    trade = sheets.get('Trade')
    funding = sheets.get('Funding')
    reward = sheets.get('Reward')
    ip = sheets.get('IP')
    spec = sheets.get('Spec')
    return trade, funding, reward, ip, spec

def load_sheets(file_name):
    """
    Excel 파일에서 시트별로 데이터를 불러옴
    Trade, Funding, Reward, IP, Spec 시트를 가져옴
    """
    if not os.path.exists(file_name):
        raise FileNotFoundError(f"{file_name} not found in cwd={os.getcwd()}")
    sheets = pd.read_excel(file_name, sheet_name=None)
    trade = sheets.get('Trade')
    funding = sheets.get('Funding')
    reward = sheets.get('Reward')
    ip = sheets.get('IP')
    spec = sheets.get('Spec')

# from abnormal_detector import load_sheets, calc_account_profits, find_time_synced_groups
FILE_NAME = "abnormalTransactionDetection.xlsx"

from datetime import timedelta


FILE_NAME = "./abnormalTransactionDetection.xlsx"
TIME_WINDOW_SECONDS = 60      # 같은 심볼/같은 방향으로 묶을 시점 차이(초)
MIN_GROUP_SIZE = 2            # 그룹으로 판단할 최소 계정 수
SCORE_WEIGHTS = {
    "size": 0.4,              # 그룹 계정수(클수록 의심)
    "same_ip": 0.2,           # 동일 IP 비중
    "funding_opposite": 0.2,  # 펀딩 부호 반대성 (내부 이전 의심)
    "reward_sync": 0.2        # 보상 동시성
}
# --------------------------

def load_sheets(file_name):
    if not os.path.exists(file_name):
        raise FileNotFoundError(f"{file_name} not found in cwd={os.getcwd()}")
    sheets = pd.read_excel(file_name, sheet_name=None)
    # 키가 'Trade','Funding','Reward','IP','Spec' 인 것을 기대
    trade = sheets.get('Trade')
    funding = sheets.get('Funding')
    reward = sheets.get('Reward')
    ip = sheets.get('IP')
    spec = sheets.get('Spec')
    return trade, funding, reward, ip, spec

def preprocess(trade, funding, reward, ip):
    # timestamps -> datetime
    trade['ts'] = pd.to_datetime(trade['ts'])
    funding['ts'] = pd.to_datetime(funding['ts'])
    reward['ts'] = pd.to_datetime(reward['ts'])
    # normalize column names if needed
    # ensure columns we will use exist
    return trade, funding, reward, ip

def find_time_synced_groups(trade, time_window_seconds=60):
    # group candidate windows: for each trade row, find others with same symbol & side within time window
    tw = pd.Timedelta(seconds=time_window_seconds)
    trade_sorted = trade.sort_values('ts').reset_index(drop=True)
    groups = []   # list of dicts {accounts:set, symbol, side, center_time}
    # Build index by symbol+side for speed
    for (symbol, side), df in trade_sorted.groupby(['symbol','side']):
        df = df.reset_index(drop=True)
        start = 0
        n = len(df)
        for end in range(n):
            while df.iloc[end]['ts'] - df.iloc[start]['ts'] > tw:
                start += 1
            if end - start >= 1:
                window_slice = df.iloc[start:end+1]
                accounts = set(window_slice['account_id'].astype(str).unique().tolist())
                if len(accounts) >= MIN_GROUP_SIZE:
                    groups.append({
                        'accounts': accounts,
                        'symbol': symbol,
                        'side': side,
                        'start_time': df.iloc[start]['ts'],
                        'end_time': df.iloc[end]['ts'],
                        'count': len(window_slice)
                    })
    # dedupe groups by accounts+symbol+time-range (create canonical key)
    seen = set()
    dedup = []
    for g in groups:
        key = (tuple(sorted(g['accounts'])), g['symbol'], g['side'], g['start_time'], g['end_time'])
        if key not in seen:
            seen.add(key)
            dedup.append(g)
    return dedup

def calc_account_profits(trade, funding, reward):
    # Trade profit: using amount_close - amount_open for same position_id & account
    trade['position_id'] = trade['position_id'].astype(str)
    open_trades = trade[trade['openclose'].str.upper() == 'OPEN'].copy()
    close_trades = trade[trade['openclose'].str.upper() == 'CLOSE'].copy()
    # align columns name possibilities: amount, price, qty
    # merge close with open on account_id + position_id + symbol
    merged = pd.merge(
        close_trades,
        open_trades,
        on=['account_id','position_id','symbol'],
        suffixes=('_close','_open'),
        how='inner'
    )
    # if amount exists, use amount diff; else use (price_close - price_open)*qty_open
    def compute_trade_profit(row):
        if 'amount_close' in row and not pd.isna(row.get('amount_close')) and 'amount_open' in row:
            try:
                return float(row['amount_close']) - float(row['amount_open'])
            except:
                pass
        # fallback
        try:
            return (float(row['price_close']) - float(row['price_open'])) * float(row['qty_open'])
        except:
            return 0.0
    if not merged.empty:
        merged['trade_profit'] = merged.apply(compute_trade_profit, axis=1)
        trade_profit_summary = merged.groupby('account_id')['trade_profit'].sum().reset_index()
    else:
        trade_profit_summary = pd.DataFrame(columns=['account_id','trade_profit'])
    # funding summary
    if 'funding_fee' in funding.columns:
        funding_summary = funding.groupby('account_id')['funding_fee'].sum().reset_index().rename(columns={'funding_fee':'funding_profit'})
    else:
        funding_summary = pd.DataFrame(columns=['account_id','funding_profit'])
    # reward summary
    if 'reward_amount' in reward.columns:
        reward_summary = reward.groupby('account_id')['reward_amount'].sum().reset_index().rename(columns={'reward_amount':'reward_profit'})
    else:
        reward_summary = pd.DataFrame(columns=['account_id','reward_profit'])
    # merge all
    df = trade_profit_summary.merge(funding_summary, on='account_id', how='outer')
    df = df.merge(reward_summary, on='account_id', how='outer')
    df = df.fillna(0)
    if 'trade_profit' not in df.columns:
        df['trade_profit'] = 0.0
    if 'funding_profit' not in df.columns:
        df['funding_profit'] = 0.0
    if 'reward_profit' not in df.columns:
        df['reward_profit'] = 0.0
    df['total_profit'] = df['trade_profit'] + df['funding_profit'] + df['reward_profit']
    return df

def ip_same_score(accounts, ip_df):
    # compute fraction of accounts sharing at least one IP
    if ip_df is None or ip_df.empty:
        return 0.0, {}
    acc_ips = {}
    for aid in accounts:
        ips = ip_df.loc[ip_df['account_id'].astype(str)==str(aid),'ip'].dropna().unique().tolist()
        acc_ips[aid] = ips
    # count most common IP occurance among accounts
    ip_counts = defaultdict(int)
    for ips in acc_ips.values():
        for ip in ips:
            ip_counts[ip] += 1
    if not ip_counts:
        return 0.0, acc_ips
    max_share = max(ip_counts.values()) / len(accounts)
    return float(max_share), acc_ips

def funding_opposite_score(accounts, funding_df, time_range=None):
    # For accounts in group, look at funding fees around group's time_range.
    # We compute proportion of accounts whose funding fees sign contrasts with group median sign.
    if funding_df is None or funding_df.empty:
        return 0.0
    subset = funding_df[funding_df['account_id'].isin(accounts)].copy()
    if subset.empty:
        return 0.0
    # optionally restrict by time_range if provided (start,end)
    if time_range:
        start,end = time_range
        subset = subset[(subset['ts'] >= start) & (subset['ts'] <= end)]
    if subset.empty:
        subset = funding_df[funding_df['account_id'].isin(accounts)].copy()
    # group by account and sum funding
    acc_sum = subset.groupby('account_id')['funding_fee'].sum()
    if acc_sum.empty:
        return 0.0
    median = acc_sum.median()
    # If median ~0, look at signs relative to mean
    sign_median = np.sign(median) if median!=0 else 0
    if sign_median == 0:
        sign_median = np.sign(acc_sum.mean())
    if sign_median == 0:
        # no signal
        return 0.0
    # proportion of accounts with opposite sign
    opposite = sum(1 for v in acc_sum if np.sign(v) != sign_median and v != 0)
    score = opposite / len(accounts)
    return float(score)

def reward_sync_score(accounts, reward_df, time_range=None):
    # calculate whether multiple accounts received reward at same minute/time
    if reward_df is None or reward_df.empty:
        return 0.0
    df = reward_df[reward_df['account_id'].isin(accounts)].copy()
    if df.empty:
        return 0.0
    df['tmin'] = df['ts'].dt.floor('min')  # minute resolution
    # find minutes where count >1
    grouped = df.groupby('tmin')['account_id'].nunique()
    if grouped.empty:
        return 0.0
    max_same = grouped.max()
    return float(max_same / len(accounts))

def score_group(g, ip_df, funding_df, reward_df):
    accounts = list(g['accounts'])
    # size score: normalized by log
    size_score = np.log1p(len(accounts)) / np.log1p(20)  # scale to ~0..1 (assumes groups up to ~20)
    same_ip_frac, acc_ips = ip_same_score(accounts, ip_df)
    funding_score = funding_opposite_score(accounts, funding_df, time_range=(g['start_time'], g['end_time']))
    reward_score = reward_sync_score(accounts, reward_df, time_range=(g['start_time'], g['end_time']))
    combined = (SCORE_WEIGHTS['size']*size_score +
                SCORE_WEIGHTS['same_ip']*same_ip_frac +
                SCORE_WEIGHTS['funding_opposite']*funding_score +
                SCORE_WEIGHTS['reward_sync']*reward_score)
    return {
        'size_score': size_score,
        'same_ip_frac': same_ip_frac,
        'funding_score': funding_score,
        'reward_score': reward_score,
        'combined_score': combined,
        'acc_ips': acc_ips
    }

def build_network_and_cluster(groups):
    # Build graph where nodes = account_id, edges between co-occurring in any group.
    G = nx.Graph()
    for g in groups:
        accts = list(g['accounts'])
        for a in accts:
            if a not in G:
                G.add_node(a)
        for i in range(len(accts)):
            for j in range(i+1, len(accts)):
                a,b = accts[i], accts[j]
                if G.has_edge(a,b):
                    G[a][b]['weight'] += 1
                else:
                    G.add_edge(a,b,weight=1)
    # Find connected components as clusters
    clusters = [set(c) for c in nx.connected_components(G) if len(c) > 1]
    return G, clusters

def main():
    print("Loading sheets from", FILE_NAME)
    trade, funding, reward, ip, spec = load_sheets(FILE_NAME)
    trade, funding, reward, ip = preprocess(trade, funding, reward, ip)

    print("Finding time-synced groups ...")
    groups = find_time_synced_groups(trade, time_window_seconds=TIME_WINDOW_SECONDS)
    print(f"Found {len(groups)} raw candidate groups (time-synced).")

    print("Calculating account profits (trade + funding + reward) ...")
    profit_df = calc_account_profits(trade, funding, reward)
    profit_df.to_csv("all_account_profits.csv", index=False)

    print("Scoring groups ...")
    scored_groups = []
    for g in groups:
        score_info = score_group(g, ip, funding, reward)
        gcopy = g.copy()
        gcopy.update(score_info)

        gcopy['group_trade_amount'] = sum(
            trade[trade['account_id'].isin(g['accounts'])]['amount'].fillna(0).astype(float)
        )
        gcopy['group_total_profit'] = profit_df[
            profit_df['account_id'].isin(g['accounts'])
        ]['total_profit'].sum() if not profit_df.empty else 0.0

        # 거래유형 자동 분류
        pos_sum = gcopy['group_total_profit']
        ip_frac = gcopy['same_ip_frac']
        fund_score = gcopy['funding_score']
        reward_score = gcopy['reward_score']

        if reward_score >= 0.5:
            trade_type = "Reward Abuse"
        elif fund_score >= 0.4:
            trade_type = "Funding Manipulation"
        elif abs(pos_sum) < 1000 and ip_frac < 0.3:
            trade_type = "Wash Trading"
        elif pos_sum > 1e6 and ip_frac < 0.2:
            trade_type = "Pump & Dump"
        else:
            trade_type = "Mixed/Other"

        gcopy['trade_type'] = trade_type
        scored_groups.append(gcopy)

    scored_groups_sorted = sorted(
        scored_groups,
        key=lambda x: (x['combined_score'], x['group_total_profit']),
        reverse=True
    )

    sg_df = pd.DataFrame([{
        'accounts': ','.join(sorted(list(g['accounts']))),
        'symbol': g['symbol'],
        'side': g['side'],
        'combined_score': g['combined_score'],
        'same_ip_frac': g['same_ip_frac'],
        'funding_score': g['funding_score'],
        'reward_score': g['reward_score'],
        'group_total_profit': g['group_total_profit'],
        'trade_type': g.get('trade_type', 'Unknown')
    } for g in scored_groups_sorted])

    sg_df.to_csv("suspicious_groups.csv", index=False)

    # --------------------------
    # 하이퍼파라미터 (기업별 경고 기준 조정용)
    # --------------------------
    # 0.35는 “의심 시작” 기준
    # 0.65는 “확정 조직적 패턴” 기준
    SUSPICIOUS_THRESHOLD = 0.65  #기업 평균 기준(0.3~0.35)


    # 의심 그룹만 필터링
    suspicious_groups = [g for g in scored_groups_sorted if g['combined_score'] >= SUSPICIOUS_THRESHOLD]

    # CSV로 전체/의심 그룹 분리 저장
    all_groups_df = pd.DataFrame(scored_groups_sorted)
    suspicious_df = pd.DataFrame(suspicious_groups)

    all_groups_df.to_csv("all_groups.csv", index=False)
    suspicious_df.to_csv("suspicious_groups_filtered.csv", index=False)

    print(f"\n✅ 전체 그룹 {len(all_groups_df)}개 중 의심 그룹 {len(suspicious_df)}개 탐지됨 (threshold={SUSPICIOUS_THRESHOLD})")



    # 의심 계정별 수익
    suspicious_accounts = set()
    for g in scored_groups_sorted[:200]:
        suspicious_accounts.update(g['accounts'])
    prof_susp = profit_df[profit_df['account_id'].isin(suspicious_accounts)].copy()
    prof_susp.to_csv("suspicious_profits.csv", index=False)

    # 수익 요약
    suspect_profit = prof_susp.copy().sort_values('total_profit', ascending=False)
    total_suspicious_profit = suspect_profit['total_profit'].sum()
    print(f"\n⚠️ 조직적 거래로 의심되는 계정들의 총 누적 수익: {total_suspicious_profit:,.2f}")
    suspect_profit.to_csv("suspicious_account_summary.csv", index=False)


    # --- Build network and clusters (🔍 조직 단위 탐지용) ---
    print("\nBuilding network graph and clusters ...")
    G, clusters = build_network_and_cluster(groups)
    print(f"Found {len(clusters)} connected clusters (size>1).")

    # --- ✅ Cluster-based Organized Trading Profit Summary (새 버전) ---
    group_summary_rows = []
    member_profit_rows = []

    for i, c in enumerate(clusters):
        members = list(c)
        member_df = profit_df[profit_df['account_id'].isin(members)][['account_id', 'total_profit']]

        if not member_df.empty:
            total_profit = member_df['total_profit'].sum()
            avg_profit = member_df['total_profit'].mean()
            max_profit = member_df['total_profit'].max()
            min_profit = member_df['total_profit'].min()
            std_profit = member_df['total_profit'].std()
            top_member = member_df.loc[member_df['total_profit'].idxmax(), 'account_id']
            top_profit = member_df['total_profit'].max()
        else:
            total_profit = avg_profit = max_profit = min_profit = std_profit = top_profit = 0
            top_member = "N/A"

        # 그룹 요약 (조직 단위)
        group_summary_rows.append({
            "Detection_Type": "Abnormal Detection - Organized Trading",
            "Group_ID": i + 1,
            "Suspect_Accounts": ', '.join(sorted(members)),
            "Symbol": "Multiple",
            "Side": "Mixed",
            "Trade_Type": "Organized Cluster",
            "Member_Count": len(members),
            "Group_Total_Profit": total_profit,
            "Avg_Member_Profit": avg_profit,
            "Top_Member": top_member,
            "Top_Member_Profit": top_profit,
            "Profit_StdDev": std_profit
        })

        # 멤버별 상세 기록
        for _, row in member_df.iterrows():
            member_profit_rows.append({
                "Group_ID": i + 1,
                "Account_ID": row['account_id'],
                "Member_Profit": row['total_profit'],
                "Profit_Ratio(%)": (row['total_profit'] / total_profit * 100) if total_profit != 0 else 0
            })

    # DataFrame 변환 및 저장
    group_summary = pd.DataFrame(group_summary_rows)
    member_profit_detail = pd.DataFrame(member_profit_rows)

    group_summary.to_csv("suspicious_group_summary.csv", index=False)
    member_profit_detail.to_csv("suspicious_member_profit.csv", index=False)


    # 콘솔 출력
    print("\n[Organized Trading Suspicious Group Summary]")
    print(group_summary.head(10).to_string(index=False))
    print(f"\n✅ Detected {len(group_summary)} organized trading suspicious clusters.")
    print("→ Results saved to suspicious_group_summary.csv and suspicious_member_profit.csv")

    # --- Additional notes file ---
with open("../analysis_readme.txt", "w", encoding="utf-8") as f:
    f.write("Generated files:\n")
    f.write(" - suspicious_groups.csv : Detected cooperative trading groups with scores\n")
    f.write(" - suspicious_group_summary.csv : Summary of each group’s total and average profit\n")
    f.write(" - suspicious_member_profit.csv : Profit details of individual members within each group\n")

    f.write("\nInterpretation tips:\n")
    f.write(" - combined_score near 1.0 => stronger evidence: multiple anomaly signals aligned (e.g., same time, same IP, mirrored positions)\n")
    f.write(" - High Profit_StdDev or Top_Member_Ratio may indicate central control or coordinated profit distribution.\n")
    f.write(" - Inspect raw trades for top groups to confirm actual cooperative manipulation patterns.\n")


    print("\nDone. Output files: suspicious_groups.csv, suspicious_profits.csv, all_account_profits.csv, suspicious_group_summary.csv, suspicious_member_profit.csv, analysis_readme.txt")

    print("\n📁 Files saved in folder:", os.getcwd())

if __name__ == "__main__":
    main()