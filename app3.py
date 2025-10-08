import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import locale
from dateutil.relativedelta import relativedelta # 月単位の正確な加算・減算

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
if bakuraku_df is not None and '金額' in bakuraku_df.columns: # 必要に応じてカラム名を調整
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
st.header("契約残存期間計算")

st.subheader("契約情報入力")
contract_start_date = st.date_input("Camel契約開始日を選択してください", value=datetime.today(), key="contract_start_date", format="YYYY/MM/DD")

st.subheader("休業期間設定")
if 'holiday_periods' not in st.session_state:
    st.session_state.holiday_periods = []

col_h_start, col_h_end, col_h_add = st.columns([0.4, 0.4, 0.2])
with col_h_start:
    new_holiday_start = st.date_input("新しい休業期間開始日", value=None, key="new_holiday_start", format="YYYY/MM/DD")
with col_h_end:
    new_holiday_end = st.date_input("新しい休業期間終了日", value=None, key="new_holiday_end", format="YYYY/MM/DD")
with col_h_add:
    st.write("") # スペーサー
    st.write("") # スペーサー
    if st.button("休業期間を追加", key="add_holiday_btn"):
        if new_holiday_start and new_holiday_end and new_holiday_start <= new_holiday_end:
            st.session_state.holiday_periods.append((new_holiday_start, new_holiday_end))
            # 重複期間を結合し、日付順にソートするロジック（オプション）
            st.session_state.holiday_periods = sorted(list(set(st.session_state.holiday_periods))) # 重複排除とソート
            
            # ここで休業期間を結合するロジックを追加
            # 例: [(2023/1/1, 2023/1/5), (2023/1/4, 2023/1/8)] -> [(2023/1/1, 2023/1/8)]
            merged_holidays = []
            if st.session_state.holiday_periods:
                current_start, current_end = st.session_state.holiday_periods[0]
                for next_start, next_end in st.session_state.holiday_periods[1:]:
                    if next_start <= current_end + timedelta(days=1): # 期間が連続または重複
                        current_end = max(current_end, next_end)
                    else:
                        merged_holidays.append((current_start, current_end))
                        current_start, current_end = next_start, next_end
                merged_holidays.append((current_start, current_end))
            st.session_state.holiday_periods = merged_holidays

            st.success(f"休業期間 {new_holiday_start.strftime('%Y/%m/%d')}〜{new_holiday_end.strftime('%Y/%m/%d')} を追加しました。")
        else:
            st.error("有効な休業期間を入力してください（開始日≦終了日）。")
        # 追加後にフォームをクリアして再入力しやすくする
        st.session_state.new_holiday_start = None
        st.session_state.new_holiday_end = None
        st.experimental_rerun() # 画面を再描画してリストを更新

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

def _add_delta_with_holidays(start_date: datetime.date, delta: relativedelta, holidays: list):
    """
    指定されたrelativedeltaをstart_dateに加算し、その間に休業日があればその分日付を延伸する。
    月単位の加算はrelativedeltaで行い、日単位の延伸は休業日をスキップする。
    """
    current_date = start_date
    target_date_raw = start_date + delta # まず休業日を無視した目標日

    # 日付を1日ずつ進めながら休業日をスキップする
    # deltaが月の加算を含む場合、まず月を加算した目標日を計算し、そこから日単位の調整を行う
    
    # ここでは、まず月単位で進め、その後に日単位のオフセットと休業日考慮を行う
    if delta.years or delta.months:
        current_date = start_date + relativedelta(years=delta.years, months=delta.months)
    
    days_to_add = delta.days # 日単位のオフセット
    
    while days_to_add > 0:
        current_date += timedelta(days=1)
        if not is_holiday(current_date, holidays):
            days_to_add -= 1
    
    # 最後に、目標日が休業日であれば、休業日を過ぎるまで進める
    while is_holiday(current_date, holidays):
        current_date += timedelta(days=1)

    return current_date


def calculate_min_cancel_date_contract_logic(start_date: datetime.date, holidays: list):
    """最短解約日（契約期間）を計算する"""
    today = datetime.today().date()
    
    # 現在の契約サイクルを特定
    # 契約開始から6ヶ月ごとのサイクルで、最も近い未来の契約終了日を見つける
    current_cycle_start = start_date
    
    while True:
        # 6ヶ月後の日付の、契約開始日と同じ日付、またはその月の最終日を基準とする
        # 契約期間は6ヶ月ごと、前日までが期間
        # 例: 1/1開始 -> 6/30終了、7/1開始 -> 12/31終了
        
        # 契約開始日のdayを基準に6ヶ月後を計算
        next_contract_end_base = current_cycle_start + relativedelta(months=6, days=-1)
        
        # 休業日を考慮して日付を延伸
        final_contract_end = _add_delta_with_holidays(current_cycle_start, relativedelta(months=6, days=-1), holidays)

        if final_contract_end >= today:
            return final_contract_end
        
        current_cycle_start = final_contract_end + timedelta(days=1) # 次のサイクルの開始日

def calculate_min_cancel_date_declared_logic(cancel_year: int, cancel_month: int, holidays: list):
    """最短解約日（申告日）を計算する"""
    try:
        # 解約希望月の最終日を基本とする
        declared_cancel_date_base = get_last_day_of_month(cancel_year, cancel_month)

        # 休業期間を考慮して、解約日を延伸
        # 解約希望日が休業日であれば、休業日を過ぎるまで延伸
        extended_cancel_date = declared_cancel_date_base
        while is_holiday(extended_cancel_date, holidays):
            extended_cancel_date += timedelta(days=1)

        return extended_cancel_date

    except ValueError:
        return None

def calculate_remaining_months_logic(today: datetime.date, final_cancel_date: datetime.date, holidays: list):
    """休業期間を考慮した残存月数を計算する"""
    if final_cancel_date < today:
        return 0

    # 今日から最終解約日までの「営業日」をカウントし、月数に換算する
    # ここでは簡易的に、最終解約日が決まっているので、その日付と今日の日付の差を月換算する
    
    # 今日から最終解約日までの営業日をカウントし、それを月数に換算することも可能
    # しかし、ここではシンプルに月数差分から休業日数を調整するアプローチを取る
    
    # relativedelta を使って、今日から最終解約日までの期間を取得
    delta = relativedelta(final_cancel_date, today)
    remaining_months = delta.years * 12 + delta.months

    # もし最終解約日が今日より後の同じ月であれば1ヶ月とカウント
    if remaining_months == 0 and final_cancel_date > today:
        remaining_months = 1
    elif final_cancel_date.day < today.day and remaining_months > 0:
        # 例えば、今日が1/20、解約日が2/15の場合、1ヶ月とせず0ヶ月とするか、繰り上げるか
        # 業務要件によるが、ここでは単純に月数が減る調整
        pass
    
    return max(0, remaining_months)


# --- 4. 最終出力フォーマット表示セクション ---
st.header("計算結果")

# 計算ボタン
if st.button("全ての計算を実行", key="execute_calculation_btn", type="primary"):
    with st.spinner('計算中...しばらくお待ちください。'):
        # 各計算結果を保持する変数
        formatted_contract_start_date = contract_start_date.strftime('%Y/%m/%d')
        
        holiday_periods_str_list = [f"{s.strftime('%Y/%m/%d')}〜{e.strftime('%Y/%m/%d')}" for s, e in st.session_state.holiday_periods]
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
            
        result_text = f"""
◆ Camel契約開始日：**{formatted_contract_start_date}**
◆ 休業期間：**{formatted_holiday_periods}**
◆ 最短解約日（契約期間）：**{calculated_min_cancel_date_contract.strftime('%Y/%m/%d') if calculated_min_cancel_date_contract else '計算不可'}**
◆ 最短解約日（申告日）：**{calculated_min_cancel_date_declared.strftime('%Y/%m/%d') if calculated_min_cancel_date_declared else '計算不可'}**
◆ 支払い状況：**{unpaid_amount:,.0f}円** （**{payment_status_text}**）
◆ 支払計画：**{total_payment_plan_amount:,.0f}円** （残存**{remaining_months}ヶ月** + 既存未入金）
"""
        st.markdown(result_text)

else:
    st.info("「全ての計算を実行」ボタンを押すと、結果が表示されます。")