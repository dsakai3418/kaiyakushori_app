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
            bakuraku_df = pd.read_csv(bakuraku_csv_file) # bakuraku_file -> bakuraku_csv_file
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
            st.session_state.holiday_periods.sort() # 日付順にソート

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

def get_holiday_days_in_period(start_period: datetime.date, end_period: datetime.date, holidays: list):
    """指定期間内の休業期間の日数を計算する"""
    total_holiday_days = 0
    for h_start, h_end in holidays:
        # 休業期間が指定期間と重複する部分を計算
        overlap_start = max(start_period, h_start)
        overlap_end = min(end_period, h_end)
        if overlap_start <= overlap_end:
            total_holiday_days += (overlap_end - overlap_start).days + 1
    return total_holiday_days

def calculate_min_cancel_date_contract_logic(start_date: datetime.date, holidays: list):
    """最短解約日（契約期間）を計算する"""
    current_contract_start = start_date
    today = datetime.today().date()

    # 現在の契約サイクルを特定
    # 契約開始から6ヶ月ごとのサイクルで、最も近い未来の契約終了日を見つける
    while True:
        # 6ヶ月後の日付の、契約開始日と同じ日付、またはその月の最終日
        next_contract_end_base_month = (current_contract_start + relativedelta(months=6)).month
        next_contract_end_base_year = (current_contract_start + relativedelta(months=6)).year
        
        # 契約開始日のdayを基準にするが、その月の最終日を超えないように調整
        target_day = min(current_contract_start.day, get_last_day_of_month(next_contract_end_base_year, next_contract_end_base_month).day)
        
        next_contract_end_base = datetime(next_contract_end_base_year, next_contract_end_base_month, target_day).date()
        
        # 契約開始日の前日までが期間とするため、-1日
        next_contract_end_base -= timedelta(days=1)

        if next_contract_end_base >= today:
            break
        current_contract_start = next_contract_end_base + timedelta(days=1)

    # 契約終了日を休業期間で延伸
    extended_end_date = next_contract_end_base
    
    # 契約サイクル期間内の休業期間日数を加算
    holiday_days_in_cycle = get_holiday_days_in_period(current_contract_start, next_contract_end_base, holidays)
    extended_end_date += timedelta(days=holiday_days_in_cycle)

    # 最終的な解約日は、その月の月末に調整（例：2023/07/15が延伸後の日付なら、2023/07/31）
    return get_last_day_of_month(extended_end_date.year, extended_end_date.month)


def calculate_min_cancel_date_declared_logic(cancel_year: int, cancel_month: int, holidays: list):
    """最短解約日（申告日）を計算する"""
    try:
        # 解約希望月の最終日を基本とする
        declared_cancel_date_base = get_last_day_of_month(cancel_year, cancel_month)

        # 申告日以降の休業期間を考慮して、解約日を延伸
        extended_cancel_date = declared_cancel_date_base
        
        # 申告月以降の休業期間の日数を加算
        # ここでは、申告月の最終日を起点として、それ以降の休業期間を加算
        # 実際に契約が有効な期間と重なる休業期間のみを考慮
        
        # 仮に、解約希望月から次の6ヶ月間の休業期間を考慮する
        look_ahead_end_date = declared_cancel_date_base + relativedelta(months=6) 
        
        holiday_days_after_declared = get_holiday_days_in_period(declared_cancel_date_base + timedelta(days=1), look_ahead_end_date, holidays) 
        extended_cancel_date += timedelta(days=holiday_days_after_declared)

        # 最終的な解約日は、延伸後の日付の月末に調整
        return get_last_day_of_month(extended_cancel_date.year, extended_cancel_date.month)

    except ValueError:
        return None

def calculate_remaining_months_logic(today: datetime.date, final_cancel_date: datetime.date, holidays: list):
    """休業期間を考慮した残存月数を計算する"""
    if final_cancel_date < today:
        return 0

    # 今日から最終解約日までの月数をカウント
    # 休業期間による契約期間の延伸は最終解約日に既に反映されていると仮定
    remaining_months = (final_cancel_date.year - today.year) * 12 + final_cancel_date.month - today.month

    # もし最終解約日の日数が今日の日数よりも前であれば、1ヶ月減らす
    # 例: 今日が1/15で解約日が2/10の場合、1ヶ月とカウント
    # 例: 今日が1/15で解約日が2/20の場合、1ヶ月とカウント
    # 月の途中の日数の扱いについては業務要件による。ここでは月の途中の日付でも1ヶ月とカウントする。
    # 例: 1/15 -> 2/14 なら1ヶ月
    # 例: 1/15 -> 2/15 なら2ヶ月 (単純な月数差分 +1)
    # ここでは、最終月の丸々残っている月を数える
    if final_cancel_date.day < today.day:
        # 例えば、今日が1/20、最終解約日が2/15の場合、1ヶ月未満とみなし、0ヶ月とする
        # 1/20 -> 2/19が1ヶ月
        # 最終解約日が今日以降であれば、少なくとも1ヶ月と数える
        pass
    
    # 今日が月初で、最終解約日が月末の場合などを考慮して、正確な月数を出す
    r_delta = relativedelta(final_cancel_date, today)
    calculated_months = r_delta.years * 12 + r_delta.months

    # 日数まで考慮し、残り日数がある程度あれば1ヶ月追加 (業務要件による)
    if r_delta.days > 0: # 例えば、残りが1日でもあれば、1ヶ月とカウントするかなど
        # ここでは、単純に月数差分のみとする
        pass

    return max(0, calculated_months)


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

        # 計算結果のコピー機能は削除
        # st_copy_to_clipboard(result_text, "計算結果をコピー")
else:
    st.info("「全ての計算を実行」ボタンを押すと、結果が表示されます。")