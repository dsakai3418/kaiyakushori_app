import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import locale

# ロケールを日本語に設定
try:
    locale.setlocale(locale.LC_ALL, 'ja_JP.UTF-8')
except locale.Error:
    st.warning("日本語ロケールの設定に失敗しました。日付の書式などが英語になる場合があります。")

st.set_page_config(layout="wide")

st.title("Camel 契約管理・請求計算システム")
st.markdown("---")

# --- 初期化 ---
if 'initialized' not in st.session_state:
    st.session_state.holiday_periods = []
    st.session_state.holiday_input_key = 0 # st.date_inputをリセットするためのキーカウンター
    st.session_state.initialized = True # 初期化フラグ

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
            if '金額' not in bakuraku_df.columns: # ここは必要に応じてカラム名を調整してください
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
if bakuraku_df is not None and '金額' in bakuraku_df.columns: # ここは必要に応じてカラム名を調整してください
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

# シミュレーション用の「今日の日付」入力フィールドを削除
# st.subheader("シミュレーション基準日")
# simulation_today_date = st.date_input(
#     "残存期間計算の基準日（今日の日付）",
#     value=datetime.today().date(), # デフォルトは実際の今日の日付
#     key="simulation_today_date",
#     format="YYYY/MM/DD",
#     help="残存期間は、ここで指定した日付を基準に計算されます。"
# )


st.subheader("休業期間設定")

col_h_start, col_h_end, col_h_add = st.columns([0.4, 0.4, 0.2])
with col_h_start:
    new_holiday_start = st.date_input(
        "新しい休業期間開始日",
        value=None,
        key=f"new_holiday_start_{st.session_state.holiday_input_key}",
        format="YYYY/MM/DD"
    )
with col_h_end:
    new_holiday_end = st.date_input(
        "新しい休業期間終了日",
        value=None,
        key=f"new_holiday_end_{st.session_state.holiday_input_key}",
        format="YYYY/MM/DD"
    )
with col_h_add:
    st.write("") # スペーサー
    st.write("") # スペーサー
    if st.button("入力期間で確定", key="confirm_holiday_btn"):
        if new_holiday_start and new_holiday_end and new_holiday_start <= new_holiday_end:
            st.session_state.holiday_periods.append((new_holiday_start, new_holiday_end))
            st.success(f"休業期間 {new_holiday_start.strftime('%Y/%m/%d')}〜{new_holiday_end.strftime('%Y/%m/%d')} を追加しました。")
            st.session_state.holiday_input_key += 1 # キーを更新してst.date_inputをリセット
            st.rerun() # UIを更新するために再実行
        else:
            st.error("有効な休業期間を入力してください（開始日≦終了日）。")


# 登録済みの休業期間を表示
if st.session_state.holiday_periods:
    st.markdown("**現在の登録済み休業期間:**")
    for i, (h_start, h_end) in enumerate(st.session_state.holiday_periods):
        st.write(f"- {h_start.strftime('%Y/%m/%d')} 〜 {h_end.strftime('%Y/%m/%d')}")
    if st.button("全ての休業期間をクリア", key="clear_holidays_btn"):
        st.session_state.holiday_periods = []
        st.session_state.holiday_input_key += 1 # キーを更新
        st.rerun() # UIを更新

st.subheader("解約情報入力")
col_cancel_year, col_cancel_month = st.columns(2)
with col_cancel_year:
    cancel_year = st.number_input("解約希望年", min_value=datetime.today().year, value=datetime.today().year, key="cancel_year")
with col_cancel_month:
    cancel_month = st.number_input("解約希望月", min_value=1, max_value=12, value=datetime.today().month, key="cancel_month")

# チェックボックスを追加
apply_cancellation_rule = st.checkbox(
    "「更新月の1ヶ月前までの申し出で解約可能」ルールを適用する",
    value=True, # デフォルトで適用する場合
    help="チェックを外すと、解約希望月をそのまま解約月として計算します（ただし、契約開始日より過去の場合は契約開始日）。"
)

billing_unit_price = st.number_input("請求単価", min_value=0, value=10000, step=1000, key="billing_unit_price", help="残存期間分の請求金額計算に使用する単価です。")

st.markdown("---")

# --- 4. 最終出力フォーマット表示セクション ---
st.header("計算結果")

# --- 計算ロジック本体 ---

# 指定期間内の休業日数を取得するヘルパー関数
def get_holiday_days_in_period(start_dt: pd.Timestamp, end_dt: pd.Timestamp, holiday_periods: list):
    total_holiday_days = 0
    for h_start, h_end in holiday_periods:
        h_start_dt = pd.to_datetime(h_start)
        h_end_dt = pd.to_datetime(h_end)
        
        overlap_start = max(start_dt, h_start_dt)
        overlap_end = min(end_dt, h_end_dt)
        
        if overlap_start <= overlap_end:
            total_holiday_days += (overlap_end - overlap_start).days + 1
    return total_holiday_days

# 次の更新日を計算するヘルパー関数
# current_reference_date: 計算の基準となる日付 (contract_start_dt など)
def find_next_renewal_date(contract_start_dt: pd.Timestamp, current_reference_date: pd.Timestamp, holiday_periods: list):
    cycle_start_dt = contract_start_dt
    
    # 契約開始日が current_reference_date より後の場合、最初の更新日を返すべき
    if contract_start_dt > current_reference_date:
        # 最初の6ヶ月サイクルを計算
        projected_renewal_dt_initial = contract_start_dt + pd.DateOffset(months=6)
        holidays_initial = get_holiday_days_in_period(contract_start_dt, projected_renewal_dt_initial - pd.Timedelta(days=1), holiday_periods)
        return projected_renewal_dt_initial + pd.Timedelta(days=holidays_initial)

    # current_reference_date 以前のサイクルをスキップし、current_reference_date を超える最初の更新日を探す
    while True:
        projected_renewal_dt = cycle_start_dt + pd.DateOffset(months=6)
        
        holidays_in_cycle = get_holiday_days_in_period(cycle_start_dt, projected_renewal_dt - pd.Timedelta(days=1), holiday_periods)
        actual_renewal_dt = projected_renewal_dt + pd.Timedelta(days=holidays_in_cycle)
        
        if actual_renewal_dt > current_reference_date:
            return actual_renewal_dt
        
        cycle_start_dt = actual_renewal_dt

# 最短解約日（契約期間）の計算ロジック
# `simulation_today_date` を `contract_start_date` に置き換え
def calculate_min_cancel_date_contract_logic(start_date: datetime.date, holiday_periods: list):
    contract_start_dt = pd.to_datetime(start_date)

    # 現在の契約サイクルを超えて、次に到来する更新日を探す (基準日は contract_start_dt)
    # ここでの find_next_renewal_date は、契約開始日以降で最初の更新日を見つけることを意図
    next_renewal_dt = find_next_renewal_date(contract_start_dt, contract_start_dt, holiday_periods)
    
    # 「更新月の1ヶ月前までの申し出で解約可能」ルールを適用
    min_cancel_dt = next_renewal_dt - pd.DateOffset(months=1)
    min_cancel_dt = min_cancel_dt.to_period('M').end_time # 月末日を取得

    # 最短解約日（契約期間）は、契約開始日より過去になることはない
    return max(min_cancel_dt.date(), contract_start_dt)

# 最短解約日（申告日）の計算ロジック
# `simulation_today_date` を `contract_start_date` に置き換え
def calculate_min_cancel_date_declared_logic(contract_start_date: datetime.date, cancel_year: int, cancel_month: int, holiday_periods: list, apply_cancellation_rule: bool):
    try:
        contract_start_dt = pd.to_datetime(contract_start_date)
        
        # ユーザーの解約希望月の月末日
        requested_cancel_dt_eom = pd.to_datetime(datetime(cancel_year, cancel_month, 1)).to_period('M').end_time
        
        # 計算の基準日 (契約開始日)
        current_reference_dt = contract_start_dt

        if not apply_cancellation_rule:
            # ルールを適用しない場合、ユーザーの希望月をそのまま解約月とする
            # ただし、契約開始日よりも過去の場合は契約開始日とする
            return max(requested_cancel_dt_eom.date(), current_reference_dt.date())

        # ルールを適用する場合
        # ユーザーの解約希望が属する更新サイクルを特定
        # `current_reference_dt` (契約開始日) を起点に更新日を見つける
        # ユーザーの希望月が契約開始日より前の場合、contract_start_dt を基準とする
        actual_reference_for_renewal = max(current_reference_dt, requested_cancel_dt_eom)
        
        next_contract_renewal_dt = find_next_renewal_date(contract_start_dt, actual_reference_for_renewal, holiday_periods)

        # 「更新月の1ヶ月前までの申し出」締め切り日
        deadline_for_current_renewal = (next_contract_renewal_dt - pd.DateOffset(months=1)).to_period('M').end_time

        # ユーザーの希望が締め切りに間に合うか？
        if requested_cancel_dt_eom <= deadline_for_current_renewal:
            # 希望が間に合う場合、希望月の月末が解約日となる
            # ただし、契約開始日より過去なら契約開始日以降で最も早い解約日
            return max(requested_cancel_dt_eom.date(), current_reference_dt.date())
        else:
            # ユーザーの希望が締め切りに間に合わない場合、次の更新サイクルまで契約が継続される
            # その次の更新日を探す (基準日は next_contract_renewal_dt)
            next_next_renewal_dt = find_next_renewal_date(contract_start_dt, next_contract_renewal_dt, holiday_periods)
            
            # その次の更新サイクルでの締め切り日
            deadline_for_next_renewal = (next_next_renewal_dt - pd.DateOffset(months=1)).to_period('M').end_time
            return deadline_for_next_renewal.date()

    except ValueError:
        return None

# 計算ボタン
if st.button("全ての計算を実行", key="execute_calculation_btn", type="primary"):
    with st.spinner('計算中...しばらくお待ちください。'):
        # 各計算結果を保持する変数
        formatted_contract_start_date = contract_start_date.strftime('%Y/%m/%d')
        
        holiday_periods_str_list = [f"{s.strftime('%Y/%m/%d')}〜{e.strftime('%Y/%m/%d')}" for s, e in st.session_state.holiday_periods]
        formatted_holiday_periods = ", ".join(holiday_periods_str_list) if holiday_periods_str_list else "設定なし"
        
        # 契約残存期間の計算を実行
        # 基準日は contract_start_date を使用
        calculated_min_cancel_date_contract = calculate_min_cancel_date_contract_logic(contract_start_date, st.session_state.holiday_periods)
        calculated_min_cancel_date_declared = calculate_min_cancel_date_declared_logic(contract_start_date, cancel_year, cancel_month, st.session_state.holiday_periods, apply_cancellation_rule)
        
        # 残存期間（月）と請求金額の計算
        remaining_months_rough = 0
        payment_plan_amount = 0
        
        if calculated_min_cancel_date_declared:
            # `simulation_today_date` を `contract_start_date` に置き換え
            # ここでは、契約開始日から最短解約日（申告日）までの月数を計算
            if calculated_min_cancel_date_declared >= contract_start_date:
                # pandas Period を使って正確な月数を計算
                start_period = pd.Period(contract_start_date, 'M')
                end_period = pd.Period(calculated_min_cancel_date_declared, 'M')
                remaining_months_rough = (end_period - start_period).n + 1 # 期間の両端を含む
            else:
                remaining_months_rough = 0 # 既に解約日を過ぎている
            
            payment_plan_amount = billing_unit_price * remaining_months_rough
            
        st.markdown(f"◆ Camel契約開始日：**{formatted_contract_start_date}**")
        st.markdown(f"◆ 休業期間：**{formatted_holiday_periods}**")
        st.markdown(f"◆ 最短解約日（契約期間）：**{calculated_min_cancel_date_contract.strftime('%Y/%m/%d') if calculated_min_cancel_date_contract else '計算不可'}**")
        st.markdown(f"◆ 最短解約日（申告日）：**{calculated_min_cancel_date_declared.strftime('%Y/%m/%d') if calculated_min_cancel_date_declared else '計算不可'}**")
        
        # 支払い状況の表示
        payment_detail = f"総請求額: {total_billed_amount:,.0f}円 "
        payment_detail += f"(内訳: NP {np_billed_amount:,.0f}円, バクラク {bakuraku_billed_amount:,.0f}円)"
        payment_detail += f", 入金額: {paid_amount:,.0f}円"
        payment_detail += f", 未入金: {unpaid_amount:,.0f}円 ({payment_status_text})"
        st.markdown(f"◆ 支払い状況：**{payment_detail}**")
        
        # --- 支払計画に発行元を追加するロジック（最終調整版） ---
        payment_plan_label = f"**{payment_plan_amount:,.0f}円** （残存**{remaining_months_rough}ヶ月**）"
        billing_entity_info = ""

        if unpaid_amount == 0:
            payment_plan_label += "（未入金なし）"
        else: # 未入金がある場合 (unpaid_amount > 0)
            if np_df is not None and bakuraku_df is not None:
                billing_entity_info = "（NP/バクラク）" # 両方ある場合は両方を記載
            elif np_df is not None:
                billing_entity_info = "（NP）"
            elif bakuraku_df is not None:
                billing_entity_info = "（バクラク）"
            else: # CSVファイルがどちらもアップロードされていないが、未入金がある場合
                billing_entity_info = "（請求元不明）"

            payment_plan_label += billing_entity_info # 未入金がある場合のみ発行元を付加

        st.markdown(f"◆ 支払計画：{payment_plan_label}")
else:
    st.info("「全ての計算を実行」ボタンを押すと、結果が表示されます。")