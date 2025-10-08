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
    st.session_state.holiday_input_key = 0
    st.session_state.initialized = True

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
            if '請求書発行日' not in np_df.columns:
                st.warning("NP CSVに'請求書発行日'カラムが見つかりません。対象期間の計算に影響する可能性があります。")
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
            if '日付' not in bakuraku_df.columns:
                st.warning("バクラク CSVに'日付'カラムが見つかりません。対象期間の計算に影響する可能性があります。")
            st.expander("バクラク CSVプレビュー").dataframe(bakuraku_df.head())
        except Exception as e:
            st.error(f"バクラク CSVの読み込み中にエラーが発生しました: {e}")
st.markdown("---")


# --- 2. 未入金金額計算セクション ---
st.header("未入金金額計算")
total_billed_amount = 0

np_billed_amount = 0
np_billing_months = set()
if np_df is not None and '請求金額' in np_df.columns and '請求書発行日' in np_df.columns:
    np_billed_amount = np_df['請求金額'].sum()
    try:
        np_df['請求書発行日_dt'] = pd.to_datetime(np_df['請求書発行日'])
        np_billing_months = set((d - pd.DateOffset(months=1)).strftime('%Y年%m月') for d in np_df['請求書発行日_dt'])
    except Exception as e:
        st.warning(f"NP CSVの「請求書発行日」の解析中にエラーが発生しました: {e}")
    st.info(f"**NPからの請求金額合計:** {np_billed_amount:,.0f}円")
    total_billed_amount += np_billed_amount
else:
    st.info("NP CSVが未アップロード、または請求金額/請求書発行日カラムが見つかりません。")

bakuraku_billed_amount = 0
bakuraku_billing_months = set()
if bakuraku_df is not None and '金額' in bakuraku_df.columns and '日付' in bakuraku_df.columns:
    bakuraku_billed_amount = bakuraku_df['金額'].sum()
    try:
        bakuraku_df['日付_dt'] = pd.to_datetime(bakuraku_df['日付'])
        bakuraku_billing_months = set((d - pd.DateOffset(months=1)).strftime('%Y年%m月') for d in bakuraku_df['日付_dt'])
    except Exception as e:
        st.warning(f"バクラク CSVの「日付」の解析中にエラーが発生しました: {e}")
    st.info(f"**バクラクからの請求金額合計:** {bakuraku_billed_amount:,.0f}円")
    total_billed_amount += bakuraku_billed_amount
else:
    st.info("バクラク CSVが未アップロード、または金額/日付カラムが見つかりません。")

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
    st.write("")
    st.write("")
    if st.button("入力期間で確定", key="confirm_holiday_btn"):
        if new_holiday_start and new_holiday_end and new_holiday_start <= new_holiday_end:
            st.session_state.holiday_periods.append((new_holiday_start, new_holiday_end))
            st.success(f"休業期間 {new_holiday_start.strftime('%Y/%m/%d')}〜{new_holiday_end.strftime('%Y/%m/%d')} を追加しました。")
            st.session_state.holiday_input_key += 1
            st.rerun()
        else:
            st.error("有効な休業期間を入力してください（開始日≦終了日）。")


# 登録済みの休業期間を表示
if st.session_state.holiday_periods:
    st.markdown("**現在の登録済み休業期間:**")
    for i, (h_start, h_end) in enumerate(st.session_state.holiday_periods):
        st.write(f"- {h_start.strftime('%Y/%m/%d')} 〜 {h_end.strftime('%Y/%m/%d')}")
    if st.button("全ての休業期間をクリア", key="clear_holidays_btn"):
        st.session_state.holiday_periods = []
        st.session_state.holiday_input_key += 1
        st.rerun()

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

# --- 計算ロジック ---

# 休業期間の合計日数を計算するヘルパー関数
def calculate_holiday_offset_days(start_date: datetime.date, end_date: datetime.date, holiday_periods: list) -> int:
    total_offset_days = 0
    for h_start, h_end in holiday_periods:
        # 重複期間を計算
        overlap_start = max(start_date, h_start)
        overlap_end = min(end_date, h_end)
        
        if overlap_start <= overlap_end:
            total_offset_days += (overlap_end - overlap_start).days + 1
    return total_offset_days


# 最短解約日（契約期間）の計算ロジック
def calculate_min_cancel_date_contract_logic(start_date: datetime.date, holiday_periods: list) -> datetime.date:
    contract_start_dt = pd.to_datetime(start_date)
    today_dt = pd.to_datetime(datetime.today().date())

    current_contract_period_start = contract_start_dt
    
    # 今日以降で最初の6ヶ月更新日を見つける
    while True:
        # 現在のサイクル期間の終了日 (休業期間未考慮)
        nominal_period_end = current_contract_period_start + pd.DateOffset(months=6) - pd.DateOffset(days=1)
        
        # この期間内の休業日数を計算
        offset_days = calculate_holiday_offset_days(current_contract_period_start.date(), nominal_period_end.date(), holiday_periods)
        
        # 休業期間で延長された実質の契約終了日
        effective_period_end = nominal_period_end + timedelta(days=offset_days)
        
        # もし実質的な終了日が今日以降であれば、それが最短解約日（契約期間）
        if effective_period_end.date() >= today_dt.date():
            return effective_period_end.date()
        
        # 次の6ヶ月サイクルへ
        current_contract_period_start = nominal_period_end + timedelta(days=1) + timedelta(days=offset_days) # 前回の実質終了日の翌日から開始
        
        # 無限ループ防止のための安全装置 (万が一の場合)
        if current_contract_period_start.year > today_dt.year + 10: # 例: 10年先まで見ても見つからなければ異常
             return today_dt.date() # またはエラーを返す

# 最短解約日（申告日）の計算ロジック
def calculate_min_cancel_date_declared_logic(cancel_year: int, cancel_month: int, holiday_periods: list) -> datetime.date:
    try:
        # 解約希望月の月末日を仮の基準日とする
        nominal_cancel_date = (pd.to_datetime(f"{cancel_year}-{cancel_month}-01") + pd.DateOffset(months=1) - pd.DateOffset(days=1)).date()
        
        # 契約開始日からnominal_cancel_dateまでの期間にある休業日数を考慮
        # この計算は、契約の「開始」からずれる休業期間と、解約希望月の影響を総合的に見る必要があるため、
        # より複雑になるが、ここではシンプルに申告希望月までの休業期間延長を考慮
        
        # ただし、解約希望日自体が休業期間によってずれることは通常ない（申し出た日が基準）
        # ここは「契約期間のずれ」によって、その月で解約できるか否かが決まる
        # そのため、ここでは申告希望月の月末を返す
        
        # もし、解約希望月が休業期間と重なる場合、その分だけ後ろにずれるという要件であれば、
        # ここに`calculate_holiday_offset_days`のようなロジックを入れる必要がありますが、
        # 通常、「解約希望月」はユーザーが指定した月であり、それが休業によって強制的に後ろ倒しになることは稀なので、
        # まずは指定された月の月末を返すようにします。
        return nominal_cancel_date
    except ValueError:
        return None

# --- 年月期間を整形するヘルパー関数 ---
def format_month_ranges(month_strings):
    if not month_strings:
        return "N/A"

    # 'YYYY年MM月'形式の文字列をdatetimeオブジェクトに変換しソート
    months_dt = sorted([datetime.strptime(m, '%Y年%m月') for m in month_strings])
    
    if not months_dt:
        return "N/A"

    ranges = []
    current_start = months_dt[0]
    current_end = months_dt[0]

    for i in range(1, len(months_dt)):
        # 前の月と現在の月が連続しているかチェック
        # pd.DateOffset(months=1)で1ヶ月進めて、その年と月が現在の月と一致するか
        next_month_from_current_end = current_end + pd.DateOffset(months=1)
        if next_month_from_current_end.year == months_dt[i].year and \
           next_month_from_current_end.month == months_dt[i].month:
            current_end = months_dt[i]
        else:
            # 連続が途切れたので、現在の範囲を追加
            if current_start == current_end:
                ranges.append(current_start.strftime('%Y年%m月'))
            else:
                ranges.append(f"{current_start.strftime('%Y年%m月')}～{current_end.strftime('%Y年%m月')}")
            current_start = months_dt[i]
            current_end = months_dt[i]
    
    # 最後の範囲を追加
    if current_start == current_end:
        ranges.append(current_start.strftime('%Y年%m月'))
    else:
        ranges.append(f"{current_start.strftime('%Y年%m月')}～{current_end.strftime('%Y年%m月')}")
    
    return "、".join(ranges)


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
        
        # --- 残存期間と請求金額の計算（修正版） ---
        remaining_months = 0
        payment_plan_amount = 0
        
        if calculated_min_cancel_date_declared:
            today_date_pd = pd.to_datetime(datetime.today().date())
            declared_cancel_date_pd = pd.to_datetime(calculated_min_cancel_date_declared)
            
            # 最短解約日（契約期間）もpd.Timestampに変換
            min_cancel_contract_date_pd = pd.to_datetime(calculated_min_cancel_date_contract)

            # 「更新月」のルールを考慮して、実際に解約される月を特定
            # 解約希望日が、最短解約日（契約期間）よりも前の場合、次の更新月まで契約が継続される
            actual_end_date_for_billing = declared_cancel_date_pd
            if declared_cancel_date_pd < min_cancel_contract_date_pd:
                # 解約希望月が契約の更新月ではない、または最短解約日より前の場合は、最短解約日を適用
                actual_end_date_for_billing = min_cancel_contract_date_pd
            
            # 今日から実際の契約終了までの月数を計算
            if actual_end_date_for_billing >= today_date_pd:
                # 月差を計算（月末日基準で計算）
                # today_date_pdの月末
                today_month_end = today_date_pd + pd.offsets.MonthEnd(0)
                # actual_end_date_for_billingの月末
                actual_end_month_end = actual_end_date_for_billing + pd.offsets.MonthEnd(0)
                
                # 月の差を計算（例: 2024/08/15 -> 2024/08/31, 2024/10/31 -> 2ヶ月）
                # today_date_pd が 2024-08-15, actual_end_date_for_billing が 2024-10-31 の場合
                # today_month_end は 2024-08-31, actual_end_month_end は 2024-10-31
                # (actual_end_month_end.year - today_month_end.year) * 12 + (actual_end_month_end.month - today_month_end.month) + 1
                # (2024-2024)*12 + (10-8) + 1 = 3ヶ月
                remaining_months = (actual_end_month_end.year - today_month_end.year) * 12 + \
                                   (actual_end_month_end.month - today_month_end.month) + 1
                remaining_months = max(0, remaining_months)
            else:
                remaining_months = 0
            
            payment_plan_amount = billing_unit_price * remaining_months
            
        st.markdown(f"◆ Camel契約開始日：**{formatted_contract_start_date}**")
        st.markdown(f"◆ 休業期間：**{formatted_holiday_periods}**")
        st.markdown(f"◆ 最短解約日（契約期間）：**{calculated_min_cancel_date_contract.strftime('%Y/%m/%d') if calculated_min_cancel_date_contract else '計算不可'}**")
        st.markdown(f"◆ 最短解約日（申告日）：**{calculated_min_cancel_date_declared.strftime('%Y/%m/%d') if calculated_min_cancel_date_declared else '計算不可'}**")
        
        # --- 支払い状況の表示を更新 (請求対象年月期間を追加) ---
        payment_detail = f"総請求額: {total_billed_amount:,.0f}円 "
        payment_detail += f"(内訳: NP {np_billed_amount:,.0f}円, バクラク {bakuraku_billed_amount:,.0f}円)"
        
        # NPとバクラクの請求対象年月期間を整形
        formatted_np_months = format_month_ranges(np_billing_months)
        formatted_bakuraku_months = format_month_ranges(bakuraku_billing_months)

        billing_month_info = []
        if formatted_np_months != "N/A":
            billing_month_info.append(f"NP対象期間: {formatted_np_months}")
        if formatted_bakuraku_months != "N/A":
            billing_month_info.append(f"バクラク対象期間: {formatted_bakuraku_months}")
        
        if billing_month_info:
            payment_detail += " " + " / ".join(billing_month_info)

        payment_detail += f", 入金額: {paid_amount:,.0f}円"
        payment_detail += f", 未入金: {unpaid_amount:,.0f}円 ({payment_status_text})"
        st.markdown(f"◆ 支払い状況：**{payment_detail}**")
        
        # --- 支払計画に発行元を追加するロジック（最終調整版） ---
        payment_plan_label = f"**{payment_plan_amount:,.0f}円** （残存**{remaining_months}ヶ月**）"
        billing_entity_info = ""

        if unpaid_amount == 0:
            payment_plan_label += "（未入金なし）"
        else: # 未入金がある場合 (unpaid_amount > 0)
            if np_df is not None and bakuraku_df is not None:
                billing_entity_info = "（バクラク）"
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