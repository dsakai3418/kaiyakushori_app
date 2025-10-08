import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import locale
from dateutil.relativedelta import relativedelta

# ロケールを日本語に設定 (st.date_inputの表示には影響しませんが、日付フォーマット指定で対応)
try:
    locale.setlocale(locale.LC_ALL, 'ja_JP.UTF-8')
except locale.Error:
    st.warning("日本語ロケールの設定に失敗しました。日付の書式などが英語になる場合があります。")

st.set_page_config(layout="wide")

st.title("Camel 契約管理・請求計算システム")
st.markdown("---")

# --- 1. ファイルアップロードセクション ---
st.header("CSVファイルアップロード")
upload_col1, upload_col2 = st.columns(2)

np_df = None
bakuraku_df = None

with upload_col1:
    st.subheader("NP CSV")
    np_csv_file = st.file_uploader("NPからの請求CSVをここにドラッグ＆ドロップ、またはファイルを選択", type=["csv"], key="np_csv")
    if np_csv_file:
        try:
            np_df = pd.read_csv(np_csv_file)
            st.success("NP CSVを正常に読み込みました。")
            if '請求金額' not in np_df.columns:
                st.warning("NP CSVに'請求金額'カラムが見つかりません。計算に影響する可能性があります。")
            st.expander("NP CSVプレビュー").dataframe(np_df.head())
        except Exception as e:
            st.error(f"NP CSVの読み込み中にエラーが発生しました: {e}")

with upload_col2:
    st.subheader("バクラク CSV")
    bakuraku_csv_file = st.file_uploader("バクラクからの請求CSVをここにドラッグ＆ドロップ、またはファイルを選択", type=["csv"], key="bakuraku_csv")
    if bakuraku_csv_file:
        try:
            bakuraku_df = pd.read_csv(bakuraku_csv_file)
            st.success("バクラク CSVを正常に読み込みました。")
            if '金額' not in bakuraku_df.columns:
                st.warning("バクラク CSVに'金額'カラムが見つかりません。計算に影響する可能性があります。")
            st.expander("バクラク CSVプレビュー").dataframe(bakuraku_df.head())
        except Exception as e:
            st.error(f"バクラク CSVの読み込み中にエラーが発生しました: {e}")
st.markdown("---")


# --- 2. 未入金金額計算セクション ---
st.header("未入金金額計算")
total_billed_amount = 0

np_billed_amount = 0
if np_df is not None and '請求金額' in np_df.columns:
    np_billed_amount = np_df['請求金額'].sum()
    st.info(f"**NPからの請求金額合計:** {np_billed_amount:,.0f}円")
    total_billed_amount += np_billed_amount
else:
    st.info("NP CSVが未アップロード、または請求金額カラムが見つかりません。")

bakuraku_billed_amount = 0
if bakuraku_df is not None and '金額' in bakuraku_df.columns: # 必要に応じてカラム名を調整してください
    bakuraku_billed_amount = bakuraku_df['金額'].sum()
    st.info(f"**バクラクからの請求金額合計:** {bakuraku_billed_amount:,.0f}円")
    total_billed_amount += bakuraku_billed_amount
else:
    st.info("バクラク CSVが未アップロード、または金額カラムが見つかりません。")

st.subheader("入金状況入力")
paid_amount = st.number_input("入金額を入力してください", min_value=0, value=0, step=1000, key="paid_amount", help="手入力で入金された金額を入力します。")

unpaid_amount = total_billed_amount - paid_amount
payment_status_text = "未入金" if unpaid_amount > 0 else "入金済み" if unpaid_amount == 0 else "過払い"

st.metric(label="現在の未入金金額", value=f"{unpaid_amount:,.0f}円", delta_color="inverse")
st.markdown(f"**支払い状況:** {payment_status_text}")
st.markdown("---")


# --- 3. 契約残存期間計算セクション ---
st.header("契約情報入力")

st.subheader("Camel契約情報")
contract_start_date = st.date_input("Camel契約開始日を選択してください", value=datetime.today(), key="contract_start_date", format="YYYY/MM/DD")

st.subheader("休業期間設定")
if 'holiday_periods' not in st.session_state:
    st.session_state.holiday_periods = []

col_h_start, col_h_end, col_h_add = st.columns([0.4, 0.4, 0.2])
with col_h_start:
    # value=None で初期化し、ユーザーが選択していない状態とする
    new_holiday_start = st.date_input("新しい休業期間開始日", value=None, key="new_holiday_start_input", format="YYYY/MM/DD")
with col_h_end:
    # value=None で初期化し、ユーザーが選択していない状態とする
    new_holiday_end = st.date_input("新しい休業期間終了日", value=None, key="new_holiday_end_input", format="YYYY/MM/DD")
with col_h_add:
    st.write("") # スペーサー
    st.write("") # スペーサー
    if st.button("入力期間で確定", key="add_holiday_btn"): # ボタン名を変更
        if new_holiday_start and new_holiday_end and new_holiday_start <= new_holiday_end:
            st.session_state.holiday_periods.append((new_holiday_start, new_holiday_end))
            
            # --- 休業期間の結合ロジック ---
            temp_holidays = sorted(st.session_state.holiday_periods)
            merged_holidays = []
            if temp_holidays:
                current_start, current_end = temp_holidays[0]
                for next_start, next_end in temp_holidays[1:]:
                    if next_start <= current_end + timedelta(days=1): # 期間が連続または重複
                        current_end = max(current_end, next_end)
                    else:
                        merged_holidays.append((current_start, current_end))
                        current_start, current_end = next_start, next_end
                merged_holidays.append((current_start, current_end))
            st.session_state.holiday_periods = merged_holidays

            st.success(f"休業期間 {new_holiday_start.strftime('%Y/%m/%d')}〜{new_holiday_end.strftime('%Y/%m/%d')} を確定しました。")
        else:
            st.error("有効な休業期間を入力してください（開始日≦終了日）。")
        
        # エラー回避のため、明示的にvalueをNoneにするのではなく、re-runでフォームをリセット
        st.experimental_rerun() # 画面を再描画してリストとフォームを更新

# 登録済みの休業期間を表示
if st.session_state.holiday_periods:
    st.markdown("**現在の登録済み休業期間:**")
    for i, (h_start, h_end) in enumerate(st.session_state.holiday_periods):
        st.write(f"- {h_start.strftime('%Y/%m/%d')} 〜 {h_end.strftime('%Y/%m/%d')}")
    if st.button("全ての休業期間をクリア", key="clear_holidays_btn"):
        st.session_state.holiday_periods = []
        st.experimental_rerun()

st.subheader("解約情報入力")
col_cancel_year, col_cancel_month = st.columns(2)
with col_cancel_year:
    cancel_year = st.number_input("解約希望年", min_value=datetime.today().year, value=datetime.today().year, key="cancel_year")
with col_cancel_month:
    cancel_month = st.number_input("解約希望月", min_value=1, max_value=12, value=datetime.today().month, key="cancel_month")
billing_unit_price = st.number_input("請求単価", min_value=0, value=10000, step=1000, key="billing_unit_price", help="残存期間分の請求金額計算に使用する単価です。")

st.markdown("---")


# --- 日付計算ヘルパー関数 ---
def get_last_day_of_month(year, month):
    """指定された年月の最終日を返す"""
    if not 1 <= month <= 12:
        raise ValueError("月は1から12の間で指定してください。")
    return (datetime(year, month, 1) + relativedelta(months=1, days=-1)).date()

def is_holiday(date: datetime.date, holidays: list):
    """指定された日付が休業期間内にあるかを判定する"""
    for h_start, h_end in holidays:
        if h_start <= date <= h_end:
            return True
    return False

# この関数は現在使用されていませんが、念のため残しています。
# def add_business_days_with_holidays(start_date: datetime.date, days_to_add: int, holidays: list):
#     """
#     指定された日数（営業日換算）をstart_dateに加算し、休業日をスキップする。
#     """
#     current_date = start_date
#     while days_to_add > 0:
#         current_date += timedelta(days=1)
#         if not is_holiday(current_date, holidays):
#             days_to_add -= 1
    
#     # 最後に、結果の日付が休業日であれば、休業日を過ぎるまで進める
#     while is_holiday(current_date, holidays):
#         current_date += timedelta(days=1)
    
#     return current_date


def calculate_min_cancel_date_contract_logic(start_date: datetime.date, holidays: list):
    """
    最短解約日（契約期間）を計算する。
    契約開始から6ヶ月ごとの自動更新で、休業期間を考慮する。
    """
    today = datetime.today().date()
    
    current_cycle_start = start_date
    
    while True:
        # まずは休業日を考慮しない6ヶ月後の日付（サイクル終了日）を計算
        next_period_end_base = current_cycle_start + relativedelta(months=6, days=-1)
        
        # current_cycle_start から next_period_end_base までの間に休業日があるかチェック
        # その休業日数分だけ next_period_end_base を延伸する
        
        extension_days = 0
        temp_check_date = current_cycle_start
        
        while temp_check_date <= next_period_end_base: # 期間全体をチェック
            if is_holiday(temp_check_date, holidays):
                extension_days += 1
            temp_check_date += timedelta(days=1)
            
        # 延伸された次の契約終了日
        next_contract_end = next_period_end_base + timedelta(days=extension_days)

        # 最終的に次の契約終了日が休業日であれば、さらに休業日を過ぎるまで延伸
        while is_holiday(next_contract_end, holidays):
            next_contract_end += timedelta(days=1)

        if next_contract_end >= today:
            return next_contract_end
        
        current_cycle_start = next_contract_end + timedelta(days=1) # 次のサイクルの開始日

def calculate_min_cancel_date_declared_logic(cancel_year: int, cancel_month: int, holidays: list):
    """
    最短解約日（申告日）を計算する。
    解約希望月をもとに、休業期間を考慮して日付を延伸する。
    """
    try:
        # 解約希望月の最終日を基本とする
        declared_cancel_date_base = get_last_day_of_month(cancel_year, cancel_month)

        # 解約希望日が休業日であれば、休業日を過ぎるまで延伸
        extended_cancel_date = declared_cancel_date_base
        while is_holiday(extended_cancel_date, holidays):
            extended_cancel_date += timedelta(days=1)

        return extended_cancel_date

    except ValueError:
        return None

def calculate_remaining_months_logic(today: datetime.date, final_cancel_date: datetime.date, holidays: list):
    """休業期間を考慮した残存月数を計算する"""
    if not final_cancel_date or final_cancel_date < today:
        return 0

    # relativedelta で年と月の差分を計算
    delta = relativedelta(final_cancel_date, today)
    remaining_months = delta.years * 12 + delta.months
    
    # 日付部分の調整：
    # final_cancel_dateがtodayの月内で、かつfinal_cancel_date.day < today.day の場合、
    # 実際には1ヶ月未満なので0ヶ月、または繰り上げ/繰り下げのルールを適用
    if remaining_months == 0 and final_cancel_date >= today:
        # 今日と同じ月で、今日以降の日付なら1ヶ月と数える
        return 1
    elif remaining_months > 0 and final_cancel_date.day < today.day:
        # 例: 今日が1/20、解約日が2/15の場合、1ヶ月とせず0ヶ月とするか、繰り上げるか
        # ここでは「未満なら繰り下げ」のルールを採用 (業務要件に合わせて調整)
        remaining_months -= 1
        
    return max(0, remaining_months)


# --- 4. 最終出力フォーマット表示セクション ---
st.header("計算結果")

# 計算ボタン
if st.button("全ての計算を実行", key="execute_calculation_btn", type="primary"):
    with st.spinner('計算中...しばらくお待ちください。'):
        # 各計算結果を保持する変数
        formatted_contract_start_date = contract_start_date.strftime('%Y/%m/%d')
        
        holiday_periods_str_list = [f"{s.strftime('%Y/%m/%d')}〜{e.strftime('%Y/%m/%d')}" for s, e in st.session_state.holiday_periods]
        # 休業期間リストが空の場合に "設定なし" と表示
        formatted_holiday_periods = ", ".join(holiday_periods_str_list) if holiday_periods_str_list else "設定なし"
        
        # 契約残存期間の計算を実行
        calculated_min_cancel_date_contract = calculate_min_cancel_date_contract_logic(contract_start_date, st.session_state.holiday_periods)
        calculated_min_cancel_date_declared = calculate_min_cancel_date_declared_logic(cancel_year, cancel_month, st.session_state.holiday_periods)
        
        # 残存期間と請求金額の計算
        today_date = datetime.today().date()
        remaining_months = calculate_remaining_months_logic(today_date, calculated_min_cancel_date_declared, st.session_state.holiday_periods)
        
        payment_plan_amount_for_remaining = billing_unit_price * remaining_months
        
        # 支払計画の合計金額は、残存期間分の請求額 + 既存未入金金額
        total_payment_plan_amount = payment_plan_amount_for_remaining + unpaid_amount
            
        # 結果表示を改行して出力
        st.markdown(f"◆ Camel契約開始日：**{formatted_contract_start_date}**")
        st.markdown(f"◆ 休業期間：**{formatted_holiday_periods}**")
        st.markdown(f"◆ 最短解約日（契約期間）：**{calculated_min_cancel_date_contract.strftime('%Y/%m/%d') if calculated_min_cancel_date_contract else '計算不可'}**")
        st.markdown(f"◆ 最短解約日（申告日）：**{calculated_min_cancel_date_declared.strftime('%Y/%m/%d') if calculated_min_cancel_date_declared else '計算不可'}**")
        st.markdown(f"◆ 支払い状況：**{unpaid_amount:,.0f}円** （**{payment_status_text}**）")
        st.markdown(f"◆ 支払計画：**{total_payment_plan_amount:,.0f}円** （残存**{remaining_months}ヶ月** + 既存未入金）")

else:
    st.info("「全ての計算を実行」ボタンを押すと、結果が表示されます。")