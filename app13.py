import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import locale

# ロケールを日本語に設定 (ただし、st.date_inputの表示には影響しないことが多い)
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
billing_unit_price = st.number_input("請求単価", min_value=0, value=10000, step=1000, key="billing_unit_price", help="残存期間分の請求金額計算に使用する単価です。")

st.markdown("---")

# --- 4. 最終出力フォーマット表示セクション ---
st.header("計算結果")

# --- 計算ロジック本体 ---
# 最短解約日（契約期間）の計算ロジック
def calculate_min_cancel_date_contract_logic(start_date: datetime.date, holiday_periods: list):
    contract_start_dt = pd.to_datetime(start_date)
    today_dt = pd.to_datetime(datetime.today().date())

    # 実質的な契約期間の長さを計算（休業期間による延伸）
    # この関数は契約期間内の休業日数合計を返す
    def get_effective_contract_days(base_start_dt, base_end_dt, holidays):
        total_effective_days = (base_end_dt - base_start_dt).days + 1
        for h_start, h_end in holidays:
            h_start_dt = pd.to_datetime(h_start)
            h_end_dt = pd.to_datetime(h_end)
            
            # 休業期間が契約期間と重複する部分を計算
            overlap_start = max(base_start_dt, h_start_dt)
            overlap_end = min(base_end_dt, h_end_dt)
            
            if overlap_start <= overlap_end:
                total_effective_days -= (overlap_end - overlap_start).days + 1 # 重複日数を差し引く
        return total_effective_days

    # 次の6ヶ月更新日を探す
    current_contract_cycle_end = contract_start_dt
    
    # 契約開始日から現在までの「実質的な」期間を計算
    total_elapsed_days = (today_dt - contract_start_dt).days + 1 # 今日までの経過日数
    
    # 休業期間による実質的な経過日数の調整（契約開始日から今日までの休業日数）
    elapsed_holiday_days = 0
    for h_start, h_end in holiday_periods:
        h_start_dt = pd.to_datetime(h_start)
        h_end_dt = pd.to_datetime(h_end)
        
        overlap_start = max(contract_start_dt, h_start_dt)
        overlap_end = min(today_dt, h_end_dt)
        
        if overlap_start <= overlap_end:
            elapsed_holiday_days += (overlap_end - overlap_start).days + 1
            
    effective_elapsed_days = total_elapsed_days - elapsed_holiday_days

    # 6ヶ月 (約182.5日) ごとに契約更新されるとして、何サイクル経過したか
    cycle_days = (pd.DateOffset(months=6) - pd.DateOffset(days=1)).days + 1 # 6ヶ月の日数 (例: 184日)
    
    cycles_passed = effective_elapsed_days // cycle_days
    
    # 現在の契約サイクルの開始日を計算
    # (契約開始日 + (経過サイクル数 * 6ヶ月))
    current_cycle_start_dt = contract_start_dt + pd.DateOffset(months=cycles_passed * 6)
    
    # 休業期間による契約期間の延伸を加味した次の更新日を計算
    # 現在の契約サイクル開始日から6ヶ月後の更新日を探し、休業期間分ずらす
    target_renewal_dt = current_cycle_start_dt + pd.DateOffset(months=6)

    # 契約終了日が休業期間によりどれだけ遅れるか計算
    offset_days = 0
    for h_start, h_end in holiday_periods:
        h_start_dt = pd.to_datetime(h_start)
        h_end_dt = pd.to_datetime(h_end)
        
        # 休業期間が現在のサイクル終了日を過ぎていたら、その日数を加算
        # 厳密には、更新日までの期間に影響する休業期間のみを考慮すべき
        if h_start_dt > current_cycle_start_dt and h_start_dt < target_renewal_dt:
            # 更新日までの期間にある休業期間の日数を加算
            overlap_start = max(current_cycle_start_dt, h_start_dt)
            overlap_end = min(target_renewal_dt, h_end_dt)
            if overlap_start <= overlap_end:
                offset_days += (overlap_end - overlap_start).days + 1
        elif h_start_dt >= target_renewal_dt:
            # 更新日以降に始まる休業期間は、この更新日の計算には影響しない
            pass

    # 最終的な次の更新日
    final_renewal_dt = target_renewal_dt + timedelta(days=offset_days)
    
    # 最短解約日（契約期間）：更新月の1ヶ月前までの申し出で解約可能
    # 例: 7月更新なら6月末日が解約可能日
    min_cancel_dt = final_renewal_dt - pd.DateOffset(months=1)
    min_cancel_dt = min_cancel_dt.to_period('M').end_time # 月末日を取得

    # もし計算結果が今日よりも過去なら、次の更新サイクルまで待つ
    # これはあくまで「現在の契約期間」での最短解約日なので、今日よりも未来であるべき
    if min_cancel_dt.date() < today_dt.date():
        # 次の更新サイクルで再度計算（現在のロジックでは自動的に未来になるはず）
        min_cancel_dt = final_renewal_dt + pd.DateOffset(months=5) # 6ヶ月後の更新日の1ヶ月前
        min_cancel_dt = min_cancel_dt.to_period('M').end_time
        
    return min_cancel_dt.date()

# 最短解約日（申告日）の計算ロジック
# ユーザーが指定した「解約希望月」が、上記の「更新月の1ヶ月前ルール」に則っていつになるかを計算
def calculate_min_cancel_date_declared_logic(contract_start_date: datetime.date, cancel_year: int, cancel_month: int, holiday_periods: list):
    try:
        requested_cancel_date = datetime(cancel_year, cancel_month, 1).date() # ユーザーの解約希望月の1日
        
        # 契約開始日からユーザーの解約希望月までの「実質的な」期間を計算
        # これにより、どの更新サイクルに解約希望日が該当するかを判断する
        contract_start_dt = pd.to_datetime(contract_start_date)
        requested_cancel_dt = pd.to_datetime(requested_cancel_date)
        
        # 契約開始日から解約希望月までの実質的な期間（休業期間を除く）
        total_days_to_requested_cancel = (requested_cancel_dt - contract_start_dt).days + 1
        
        effective_days_to_requested_cancel = total_days_to_requested_cancel
        for h_start, h_end in holiday_periods:
            h_start_dt = pd.to_datetime(h_start)
            h_end_dt = pd.to_datetime(h_end)
            
            overlap_start = max(contract_start_dt, h_start_dt)
            overlap_end = min(requested_cancel_dt, h_end_dt)
            
            if overlap_start <= overlap_end:
                effective_days_to_requested_cancel -= (overlap_end - overlap_start).days + 1

        cycle_days = (pd.DateOffset(months=6) - pd.DateOffset(days=1)).days + 1 # 6ヶ月の日数 (例: 184日)
        
        # ユーザーの解約希望日が属する更新サイクルを特定
        # 契約開始日から何サイクル目の更新月になるか
        # effective_days_to_requested_cancel が0以下になることはないが、念のためmax
        cycles_at_requested = max(0, (effective_days_to_requested_cancel -1) // cycle_days) # -1は1サイクル未満を0とするため
        
        # そのサイクル終了後の更新月
        actual_renewal_dt_after_requested = contract_start_dt + pd.DateOffset(months=(cycles_at_requested + 1) * 6)

        # ここに休業期間による実際の更新日のずれを加える
        offset_days = 0
        for h_start, h_end in holiday_periods:
            h_start_dt = pd.to_datetime(h_start)
            h_end_dt = pd.to_datetime(h_end)
            
            # 契約開始日からこの更新日までの間の休業期間
            overlap_start = max(contract_start_dt, h_start_dt)
            overlap_end = min(actual_renewal_dt_after_requested, h_end_dt)
            
            if overlap_start <= overlap_end:
                offset_days += (overlap_end - overlap_start).days + 1
        
        actual_renewal_dt_after_requested_adjusted = actual_renewal_dt_after_requested + timedelta(days=offset_days)


        # 「更新月の1ヶ月前までの申し出で解約可能」ルールを適用
        # 例: 7月更新なら6月末日が解約可能日
        min_cancel_declared_dt = actual_renewal_dt_after_requested_adjusted - pd.DateOffset(months=1)
        min_cancel_declared_dt = min_cancel_declared_dt.to_period('M').end_time # 月末日を取得
        
        return min_cancel_declared_dt.date()
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
        calculated_min_cancel_date_contract = calculate_min_cancel_date_contract_logic(contract_start_date, st.session_state.holiday_periods)
        calculated_min_cancel_date_declared = calculate_min_cancel_date_declared_logic(contract_start_date, cancel_year, cancel_month, st.session_state.holiday_periods)
        
        # 残存期間（月）と請求金額の計算
        remaining_months_rough = 0
        payment_plan_amount = 0
        
        if calculated_min_cancel_date_declared:
            today_date = datetime.today().date()
            
            # 最短解約日（申告日）が今日よりも未来の場合のみ残存期間を計算
            if calculated_min_cancel_date_declared >= today_date:
                # `pd.to_datetime` で正確な月差を計算
                # 例: 8/1に10/31解約なら3ヶ月 (8,9,10月)
                months_diff = (pd.to_datetime(calculated_min_cancel_date_declared).to_period('M') - pd.to_datetime(today_date).to_period('M')).n
                remaining_months_rough = max(0, months_diff + 1) # 現在月を含めて計算

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