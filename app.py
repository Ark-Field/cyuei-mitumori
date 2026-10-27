import streamlit as st
import openpyxl
import io
import json
import os
import requests

# ページ基本設定
st.set_page_config(page_title="汎用エクセル帳票ジェネレーター", layout="wide")
st.title("📊 汎用エクセル帳票発行システム")

# ── 【安全なSecrets取得処理（エラー回避版）】 ──
if "pocket" in st.secrets:
    POCKET_API_KEY = st.secrets["pocket"].get("api_key", "XZtsgp27LemoHhCkwkoEiS2MZMyCU4Ay7qyYF8PS")
    POCKET_BASE_URL = st.secrets["pocket"].get("base_url", "https://app070.at-pocket.com/arkf05_bb")
else:
    # Secretsが未設定の場合でもエラー落ちしないためのフォールバック
    POCKET_API_KEY = "XZtsgp27LemoHhCkwkoEiS2MZMyCU4Ay7qyYF8PS"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"

APP_ID_ESTIMATE = 32 # 見積アプリID
MAPPING_FILE = "template_mappings.json"

# デフォルトのマッピング設定（初期データ）
DEFAULT_MAPPINGS = {
    "mitumoriV1": {
        "見積番号": "A2",
        "顧客名": "A3",
        "発行日": "F3",
        "合計金額": "F19"
    },
    "mitumoriV2": {
        "見積番号": "E1",
        "顧客名": "B2",
        "発行日": "E2",
        "合計金額": "F25"
    }
}

# マッピング設定の読み込み
def load_mappings():
    if os.path.exists(MAPPING_FILE):
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return DEFAULT_MAPPINGS
    return DEFAULT_MAPPINGS

# マッピング設定の保存
def save_mappings(mappings):
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(mappings, f, ensure_ascii=False, indent=4)

mappings_data = load_mappings()

# ── @pocketからのパラメータ受け取り ──
query_params = st.query_params
target_record_id = query_params.get("record_id", None)
passed_estimate_no = query_params.get("estimate_no", "")

# @pocketのAPI（アプリID: 32）からデータを取得する関数
def fetch_pocket_data(record_id):
    if not record_id:
        return {}
    
    api_url = f"{POCKET_BASE_URL}/api/items/{record_id}"
    headers = {"Content-Type": "application/json"}
    payload = {
        "ApiVersion": 1.1,
        "ApiKey": POCKET_API_KEY
    }
    
    try:
        response = requests.post(api_url, json=payload, headers=headers)
        if response.status_code == 200:
            result = response.json()
            data = result.get("Response", {}).get("Data", {})
            return {
                "見積番号": passed_estimate_no or data.get("Title", ""),
                "顧客名": data.get("Body", "サンプル顧客"),
                "発行日": "2026-10-27",
                "合計金額": "150000"
            }
    except Exception as e:
        st.error(f"API通信エラー: {e}")
    return {}

# データの準備
pocket_data = {}
if target_record_id:
    st.sidebar.success(f"🔗 @pocket 連携中\n\n(ID: {target_record_id} / 見積番号: {passed_estimate_no})")
    pocket_data = fetch_pocket_data(target_record_id)
else:
    pocket_data = {
        "見積番号": passed_estimate_no or "EST-001",
        "顧客名": "テスト株式会社",
        "発行日": "2026-10-27",
        "合計金額": "100000"
    }

# ── タブ構成 ──
tab1, tab2 = st.tabs(["📥 帳票ダウンロード", "⚙️ テンプレート・項目マッピング設定"])

# ==========================================
# タブ1：ダウンロード画面
# ==========================================
with tab1:
    st.header("帳票の生成とダウンロード")
    
    selected_template = st.selectbox(
        "使用するテンプレートを選択",
        options=list(mappings_data.keys())
    )

    st.divider()
    st.subheader("📝 埋め込みデータの確認・編集")
    
    current_mapping = mappings_data.get(selected_template, {})
    input_values = {}
    
    for field_name, cell_pos in current_mapping.items():
        default_val = pocket_data.get(field_name, "サンプル値")
        input_values[field_name] = st.text_input(
            f"{field_name} (セル: {cell_pos})",
            value=default_val
        )

    st.divider()

    # エクセル生成ボタン
    if st.button("🚀 帳票エクセルを生成する", type="primary"):
        template_filename = f"{selected_template}.xlsx"
        
        if not os.path.exists(template_filename):
            st.error(f"テンプレートファイル '{template_filename}' がアプリのルートディレクトリに見つかりません。")
        else:
            try:
                # openpyxlでテンプレートのデザインを完全に維持したまま読み込み
                wb = openpyxl.load_workbook(template_filename)
                ws = wb.active
                
                # マッピング設定に従ってデータを指定セルに流し込む
                for field_name, cell_pos in current_mapping.items():
                    val = input_values.get(field_name, "")
                    if cell_pos in ws:
                        ws[cell_pos] = val
                
                # メモリ上にバイナリとして保存
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("✨ エクセルの生成に成功しました！")
                st.download_button(
                    label=f"📥 {selected_template}_出力結果.xlsx をダウンロード",
                    data=output,
                    file_name=f"{selected_template}_出力結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"エクセル生成エラー: {e}")

# ==========================================
# タブ2：マッピング設定管理
# ==========================================
with tab2:
    st.header("テンプレートごとのセル関連付け設定")
    st.write("テンプレートごとに、項目名とエクセルセル番地の対応関係を管理できます（変更は自動で保存されます）。")
    
    edit_template = st.selectbox(
        "編集するテンプレートを選択",
        options=list(mappings_data.keys()),
        key="edit_select"
    )
    current_rules = mappings_data.get(edit_template, {})
    
    updated_rules = {}
    cols_to_delete = []
    
    for field_name, cell_pos in list(current_rules.items()):
        c1, c2, c3 = st.columns([3, 3, 1])
        with c1:
            new_key = st.text_input("項目名", value=field_name, key=f"key_{edit_template}_{field_name}")
        with c2:
            new_val = st.text_input("セル番地", value=cell_pos, key=f"val_{edit_template}_{field_name}")
        with c3:
            if st.button("🗑️ 削除", key=f"del_{edit_template}_{field_name}"):
                cols_to_delete.append(field_name)
        if new_key and new_key not in cols_to_delete:
            updated_rules[new_key] = new_val

    for k in cols_to_delete:
        if k in updated_rules:
            del updated_rules[k]

    st.divider()
    st.write("➕ 新しいマッピング項目の追加")
    new_field = st.text_input("追加する項目名 (例: 担当者名)", key=f"new_field_{edit_template}")
    new_cell = st.text_input("対応するセル番地 (例: B4)", key=f"new_cell_{edit_template}")
    
    if st.button("追加 / 更新を保存", type="primary"):
        if new_field and new_cell:
            updated_rules[new_field] = new_cell
        mappings_data[edit_template] = updated_rules
        save_mappings(mappings_data)
        st.success("✅ マッピング設定を保存しました！")
        st.rerun()