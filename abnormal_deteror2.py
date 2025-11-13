# 📊 EDA for Funding Fee (정상 범위 탐색 + 이상치 시각화)
# -------------------------------------------------------
# 필요 라이브러리
import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np



# 1️⃣ 엑셀 불러오기
file_path = "abnormalTransactionDetection.xlsx"   # 같은 폴더에 있어야 함
sheets = pd.read_excel(file_path, sheet_name=None)

# 'Funding' 시트 가져오기
funding = sheets['Funding']
funding['funding_fee'] = funding['funding_fee'].astype(float)

print("✅ Funding 시트 로드 완료")
print(f"데이터 개수: {len(funding)}개")
print(funding.head())


# 2️⃣ 기본 통계값 계산
mean = funding['funding_fee'].mean()
std = funding['funding_fee'].std()
Q1 = funding['funding_fee'].quantile(0.25)
Q3 = funding['funding_fee'].quantile(0.75)
IQR = Q3 - Q1

# 정상 범위 계산 (Z-score, IQR 두 방식)
z_lower, z_upper = mean - 3 * std, mean + 3 * std
iqr_lower, iqr_upper = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR

print("\n📏 [통계 요약]")
print(f"평균 (Mean): {mean:.8f}")
print(f"표준편차 (Std): {std:.8f}")
print(f"Z-Score 기준 정상 범위: [{z_lower:.8f}, {z_upper:.8f}]")
print(f"IQR 기준 정상 범위: [{iqr_lower:.8f}, {iqr_upper:.8f}]")

# -------------------------------------------------------

# 3️⃣ 히스토그램 (Z-score 기준 정상 범위 표시)
plt.figure(figsize=(10,6))
plt.hist(funding['funding_fee'], bins=60, color='lightsteelblue', edgecolor='gray')
plt.axvline(mean, color='red', linestyle='--', linewidth=2, label='Mean')
plt.axvline(z_lower, color='orange', linestyle='--', label='Lower Bound (-3σ)')
plt.axvline(z_upper, color='orange', linestyle='--', label='Upper Bound (+3σ)')
plt.title("Funding Fee Distribution (Z-score 기준)", fontsize=14)
plt.xlabel("Funding Fee Value")
plt.ylabel("Count")
plt.legend()
plt.grid(True, alpha=0.3)
plt.show()

# -------------------------------------------------------

# 4️⃣ Boxplot (IQR 기준 이상치 확인)
plt.figure(figsize=(8,5))
plt.boxplot(funding['funding_fee'], vert=False, patch_artist=True,
            boxprops=dict(facecolor='lightyellow', color='gray'),
            medianprops=dict(color='red'))
plt.title("Funding Fee Boxplot (IQR 기준)", fontsize=14)
plt.xlabel("Funding Fee")
plt.grid(True, axis='x', alpha=0.3)
plt.show()

# -------------------------------------------------------

# 5️⃣ (선택) 시계열 변화 추이
if 'ts' in funding.columns:
    funding['ts'] = pd.to_datetime(funding['ts'])
    funding_sorted = funding.sort_values('ts')

    plt.figure(figsize=(12,6))
    plt.plot(funding_sorted['ts'], funding_sorted['funding_fee'], color='navy', alpha=0.7)
    plt.axhline(mean, color='red', linestyle='--', label='Mean')
    plt.axhline(z_lower, color='orange', linestyle='--', label='Lower Bound (-3σ)')
    plt.axhline(z_upper, color='orange', linestyle='--', label='Upper Bound (+3σ)')
    plt.title("Funding Fee Over Time", fontsize=14)
    plt.xlabel("Timestamp")
    plt.ylabel("Funding Fee")
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.show()

# -------------------------------------------------------

# 6️⃣ 이상치 탐지 및 출력
funding['is_abnormal_z'] = (funding['funding_fee'] < z_lower) | (funding['funding_fee'] > z_upper)
abnormal_df = funding[funding['is_abnormal_z']]

print(f"\n🚨 이상 거래 건수(Z-score 기준): {len(abnormal_df)}건")
print(funding.head())

# =========================================
# 7️⃣ Funding 이상치 기반 조직적 거래 의심 구간 탐지 준비
# =========================================

# Funding 이상치에서 계정과 시간대 추출
abnormal_accounts = abnormal_df['account_id'].unique()
abnormal_times = pd.to_datetime(abnormal_df['ts']).dt.floor('min').unique()

print("\n🔍 Funding 기반 이상 의심 계정 수:", len(abnormal_accounts))
print("🔍 이상 Funding 발생 시간대 수:", len(abnormal_times))

# -----------------------------------------
# Trade/Reward/IP/Spec 시트 불러오기
# -----------------------------------------
trade = sheets['Trade']
reward = sheets['Reward']
ip = sheets['IP']
spec = sheets['Spec']

# 타입 정리
trade['ts'] = pd.to_datetime(trade['ts'])
reward['ts'] = pd.to_datetime(reward['ts'])
funding['ts'] = pd.to_datetime(funding['ts'])

# -----------------------------------------
# 1) 이상치 시간 ±1분 안에 거래한 Trade 필터링
# -----------------------------------------
TIME_WINDOW = pd.Timedelta(seconds=60)

candidate_trades = trade[
    trade['ts'].dt.floor('min').isin(abnormal_times)
]

print("⚠️ Funding 이상치 시간대에 거래 발생한 Trade 수:", len(candidate_trades))

# -----------------------------------------
# 2) 이전에 정의한 그룹탐색 알고리즘: 같은 심볼 + 같은 방향 + 1분 이내
# -----------------------------------------
groups = {}

for symbol in candidate_trades['symbol'].unique():
    df_sym = candidate_trades[candidate_trades['symbol'] == symbol]
    df_sym = df_sym.sort_values('ts')

    for t_min in df_sym['ts'].dt.floor('min').unique():
        group_df = df_sym[df_sym['ts'].dt.floor('min') == t_min]

        # 같은 방향(side)별로 그룹화
        for side in ['LONG', 'SHORT']:
            g = group_df[group_df['side'] == side]
            if len(g) >= 2:
                key = f"{symbol}_{side}_{t_min}"
                groups[key] = g['account_id'].unique().tolist()

print("🔗 발견된 후보 그룹 수:", len(groups))

# -----------------------------------------
# 3) 그룹별 점수 계산 템플릿 (size, ip, funding-sign, reward)
# -----------------------------------------

import numpy as np

def calculate_group_scores(account_list):
    # size_score
    size_score = np.log(1 + len(account_list))

    # same_ip score
    df_ip = ip[ip['account_id'].isin(account_list)]
    same_ip_frac = df_ip['ip'].nunique() / len(account_list)
    same_ip_frac = 1 - same_ip_frac  # IP가 적게 다양할수록 조직성 ↑

    # funding score
    df_f = funding[funding['account_id'].isin(account_list)]
    funding_sign = np.sign(df_f.groupby('account_id')['funding_fee'].sum())
    median_sign = np.sign(funding_sign.median())
    funding_score = np.mean(funding_sign != median_sign)

    # reward score
    df_r = reward[reward['account_id'].isin(account_list)]
    df_r['t_min'] = df_r['ts'].dt.floor('min')
    reward_score = df_r['t_min'].value_counts().max() / len(account_list) if len(df_r) > 0 else 0

    combined = (
            0.4 * size_score +
            0.2 * same_ip_frac +
            0.2 * funding_score +
            0.2 * reward_score
    )
    return size_score, same_ip_frac, funding_score, reward_score, combined


# 실제 그룹 점수 계산
group_scores = []

for key, accounts in groups.items():
    s, ip_s, f_s, r_s, c = calculate_group_scores(accounts)
    group_scores.append([key, accounts, s, ip_s, f_s, r_s, c])

group_scores_df = pd.DataFrame(group_scores,
                               columns=['group_key', 'accounts', 'size_score', 'ip_score', 'funding_score', 'reward_score', 'combined'])

print("\n📊 그룹 점수 계산 결과:")
print(group_scores_df.sort_values('combined', ascending=False).head())

threshold = group_scores_df['combined'].quantile(0.90)
print(f"\n🚨 조직판정 Threshold (상위 10%): {threshold:.4f}")

suspected_groups = group_scores_df[group_scores_df['combined'] >= threshold]
print(f"🔍 조직 의심 그룹 수: {len(suspected_groups)}")

suspected_accounts_top10 = (
    suspected_groups
    .explode("accounts")['accounts']
    .unique()
)

print(f"🚨 최종 조직 의심 계정 수 (상위 10% 기준): {len(suspected_accounts_top10)}")
print(suspected_accounts_top10)


# =========================================
# 8️⃣ CSV 저장: Funding outliers
# =========================================

funding_outliers_path = "funding_outliers.csv"
abnormal_df.to_csv(funding_outliers_path, index=False)
print(f"📁 Funding 이상치 계정 저장 완료 → {funding_outliers_path}")

# =========================================
# 9️⃣ CSV 저장: 조직 의심 그룹 정보
# =========================================

group_scores_path = "suspected_groups_from_funding.csv"
group_scores_df.to_csv(group_scores_path, index=False)
print(f"📁 조직 의심 그룹 저장 완료 → {group_scores_path}")

# =========================================
# 🔟 CSV 저장: 최종 용의 계정 리스트
# =========================================

suspected_accounts = (
    group_scores_df
    .explode("accounts")  # 계정 목록을 행으로 펼침
    .groupby("accounts")['combined']
    .mean()
    .reset_index()
    .rename(columns={'accounts': 'account_id', 'combined': 'avg_combined_score'})
)

suspected_accounts_path = "suspected_accounts.csv"
suspected_accounts.to_csv(suspected_accounts_path, index=False)
print(f"📁 최종 용의 계정 저장 완료 → {suspected_accounts_path}")
# =========================================
# 11️⃣ 상위 10% 조직 의심 계정 저장 (최종 범인)
# =========================================

pd.DataFrame({'account_id': suspected_accounts_top10}).to_csv(
    "suspected_accounts_top10.csv",
    index=False
)
print("📁 최종 범인(상위 10%) 저장 완료 → suspected_accounts_top10.csv")


# =========================================
# 12️⃣ 수익 계산 준비
# =========================================

# Trade profit (LONG = +, SHORT = -)
if 'qty' in trade.columns and 'price' in trade.columns:
    trade['profit'] = trade.apply(
        lambda row: row['qty'] * row['price'] * (1 if row['side'] == 'LONG' else -1),
        axis=1
    )
else:
    trade['profit'] = 0

funding_profit_all = funding.groupby('account_id')['funding_fee'].sum()
reward_profit_all = reward.groupby('account_id')['reward_amount'].sum()


# =========================================
# 13️⃣ 전체 용의자 수익 요약 (suspected_accounts_profit.csv)
# =========================================

suspected_accounts['trade_profit'] = suspected_accounts['account_id'].map(trade.groupby('account_id')['profit'].sum()).fillna(0)
suspected_accounts['funding_profit'] = suspected_accounts['account_id'].map(funding_profit_all).fillna(0)
suspected_accounts['reward_profit'] = suspected_accounts['account_id'].map(reward_profit_all).fillna(0)

suspected_accounts['total_profit'] = (
        suspected_accounts['trade_profit'] +
        suspected_accounts['funding_profit'] +
        suspected_accounts['reward_profit']
)

suspected_accounts.to_csv("suspected_accounts_profit.csv", index=False)
print("📁 전체 용의자 수익 요약 저장 완료 → suspected_accounts_profit.csv")

# =========================================
# 14️⃣ 상위 10% 최종 범인 수익 요약 (suspected_accounts_top10_profit.csv)
# =========================================

sus_top10_df = pd.DataFrame({'account_id': suspected_accounts_top10})

sus_top10_df['trade_profit'] = sus_top10_df['account_id'].map(trade.groupby('account_id')['profit'].sum()).fillna(0)
sus_top10_df['funding_profit'] = sus_top10_df['account_id'].map(funding_profit_all).fillna(0)
sus_top10_df['reward_profit'] = sus_top10_df['account_id'].map(reward_profit_all).fillna(0)

sus_top10_df['total_profit'] = (
        sus_top10_df['trade_profit'] +
        sus_top10_df['funding_profit'] +
        sus_top10_df['reward_profit']
)

sus_top10_df.to_csv("suspected_accounts_top10_profit.csv", index=False)
print("📁 최종 범인(상위10%) 수익 요약 저장 완료 → suspected_accounts_top10_profit.csv")

