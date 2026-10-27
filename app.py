import streamlit as st
import openpyxl
import io
import json
import os
import requests
import pandas as pd
import copy

# ページ基本設定
st.set_page_config(page_title="汎用エクセル帳票ジェネレーター", layout="wide")
st.title("📊 汎用エクセル帳票発行システム（エクセルテンプレート連動型）")

# ── 【安全なSecrets取得処理】 ──
if "pocket" in st.secrets:
    POCKET_API_KEY = st.secrets["pocket"].get("api_key", "XZtsgp27LemoHhCkwkoEiS2MZMyCU4Ay7qyYF8PS")
    POCKET_BASE_URL = st.secrets["pocket"].get("base_url", "https://app070.at-pocket.com/arkf05_bb")
else:
    POCKET_API_KEY = "XZtsgp27LemoHhCkwkoEiS2MZMyCU4Ay7qyYF8PS"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"

MAPPING_FILE = "template_mappings.json"

# ── 【デフォルト設定（プレースホルダーと@pocketフィールドIDの紐づけ）】 ──
DEFAULT_MAPPINGS = {
    "mitumori_template": {
        "header": {
            "#field-2": "Title",      # 見積番号
            "#field-36": "ClassA",    # 顧客名
            "#field-15": "ClassB",    # 品名
            "#field-34": "NumA",      # 金額
            "#field-37": "ClassC"     # 担当者
        },
        "details": {
            "template_keyword": "#field-31_2", # このキーワードが含まれる行を明細の雛形とする
            "columns": {
                "#field-31_2": "ItemName",   # 品名・詳細仕様
                "#field-31_4": "Quantity",   # 数量
                "#field-31_3": "UnitPrice",  # 単価
                "#field-31_6": "Amount"      # 金額
            }
        }
    }
}

# マッピング設定の読み込み・保存
def load_mappings():
    if os.path.exists(MAPPING_FILE):
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return DEFAULT_MAPPINGS
    return DEFAULT_MAPPINGS

def save_mappings(mappings):
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(mappings, f, ensure_ascii=False, indent=4)

mappings_data = load_mappings()

# パラメータ受け取り
query_params = st.query_params
target_record_id = query_params.get("record_id", None)
passed_estimate_no = query_params.get("estimate_no", "")

# @pocketからデータを取得する関数
def fetch_pocket_record(record_id):
    if not record_id:
        return {}, []
    
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
            
            # @pocketの子テーブル（明細）データを取り出し（なければサンプル）
            raw_details = data.get("Results", [])
            if not raw_details:
                raw_details = [
                    {"ItemName": "特選 生ウニ（100g）ソート品", "Quantity": 10, "UnitPrice": 5000, "Amount": 50000},
                    {"ItemName": "厳選 イクラ醤油漬け（500g）", "Quantity": 5, "UnitPrice": 8000, "Amount": 40000},
                ]
            return data, raw_details
            
    except Exception as e:
        st.error(f"API通信エラー: {e}")
    return {}, []

raw_pocket_data, pocket_details_raw = fetch_pocket_record(target_record_id)

if not target_record_id:
    raw_pocket_data = {
        "Title": passed_estimate_no or "EST-001",
        "ClassA": "株式会社 小林断截",
        "ClassB": "特殊断截加工および用紙セット",
        "NumA": "150000",
        "ClassC": "鈴木 太郎"
    }
    pocket_details_raw = [
        {"ItemName": "上質紙 70kg 四六判 裁断加工", "Quantity": 1000, "UnitPrice": 50, "Amount": 50000},
        {"ItemName": "コート紙 135kg 梱包・納品費用", "Quantity": 1, "Loss": 0, "UnitPrice": 100000, "Amount": 100000}
    ]

# ── タブ構成 ──
tab1, tab2 = st.tabs(["📥 帳票ダウンロード", "⚙️ テンプレート・フィールド設定"])

# ==========================================
# タブ1：ダウンロード画面
# ==========================================
with tab1:
    st.header("帳票の生成とダウンロード")
    
    # ルートディレクトリにある .xlsx ファイルを自動検出
    template_files = [f.replace(".xlsx", "") for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~")]
    if not template_files:
        template_files = list(mappings_data.keys())

    selected_template = st.selectbox(
        "使用するエクセルテンプレートを選択",
        options=template_files
    )
    
    current_config = mappings_data.get(selected_template, {"header": {}, "details": {}})
    header_rules = current_config.get("header", {})
    detail_rules = current_config.get("details", {})

    st.divider()
    st.subheader("📝 1. ヘッダー情報の確認・編集")
    
    input_headers = {}
    cols = st.columns(2)
    i = 0
    for placeholder, pocket_key in header_rules.items():
        default_val = passed_estimate_no if placeholder == "#field-2" else raw_pocket_data.get(pocket_key, "")
        with cols[i % 2]:
            input_headers[placeholder] = st.text_input(
                f"{placeholder} (← @pocketキー: {pocket_key})",
                value=str(default_val)
            )
        i += 1

    st.divider()
    st.subheader("📝 2. 明細行の確認・編集")
    
    col_mappings = detail_rules.get("columns", {})
    editable_rows = []
    for item in pocket_details_raw:
        row_dict = {}
        for placeholder, pocket_key in col_mappings.items():
            row_dict[placeholder] = item.get(pocket_key, "")
        editable_rows.append(row_dict)
        
    df_details_input = pd.DataFrame(editable_rows) if editable_rows else pd.DataFrame(columns=list(col_mappings.keys()))
    
    edited_df = st.data_editor(
        df_details_input,
        num_rows="dynamic",
        use_container_width=True,
        key="detail_editor"
    )

    st.divider()

    # 🚀 エクセル生成ボタン
    if st.button("🚀 帳票エクセルを生成する", type="primary"):
        template_filename = f"{selected_template}.xlsx"
        
        if not os.path.exists(template_filename):
            st.error(f"テンプレートファイル '{template_filename}' がアプリのルートディレクトリに見つかりません。エクセルファイルをアップロードまたは配置してください。")
        else:
            try:
                wb = openpyxl.load_workbook(template_filename)
                ws = wb.active
                
                # ① ヘッダープレースホルダーの置換処理（全セル走査）
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            for placeholder, val_key in header_rules.items():
                                if placeholder in cell.value:
                                    replacement_val = input_headers.get(placeholder, "")
                                    cell.value = cell.value.replace(placeholder, str(replacement_val))

                # ② 明細行の動的展開処理
                template_keyword = detail_rules.get("template_keyword", "#field-31_2")
                target_row_idx = None
                
                # テンプレート内でキーワードが含まれる行を検索
                for r in range(1, ws.max_row + 1):
                    for c in range(1, ws.max_column + 1):
                        val = ws.cell(row=r, column=c).value
                        if val and isinstance(val, str) and template_keyword in val:
                            target_row_idx = r
                            break
                    if target_row_idx:
                        break
                
                if target_row_idx and not edited_df.empty:
                    # テンプレート行のスタイル・結合情報をコピー用に保持
                    template_row_cells = [ws.cell(row=target_row_idx, column=c).value for c in range(1, ws.max_column + 1)]
                    
                    # 最初の1行目は既存の行にそのまま書き込み、2行目以降は下へ挿入
                    for index, row in edited_df.iterrows():
                        current_row = target_row_idx + index
                        if index > 0:
                            ws.insert_rows(current_row)
                            # 書式・高さをコピー
                            ws.row_dimensions[current_row].height = ws.row_dimensions[target_row_idx].height
                        
                        # 各列のプレースホルダーを置換
                        for c in range(1, ws.max_column + 1):
                            cell = ws.cell(row=current_row, column=c)
                            orig_val = ws.cell(row=target_row_idx, column=c).value if index == 0 else template_row_cells[c-1]
                            
                            if orig_val and isinstance(orig_val, str):
                                new_val = orig_val
                                for placeholder, pocket_key in col_mappings.items():
                                    if placeholder in new_val:
                                        row_val = row.get(placeholder, "")
                                        new_val = new_val.replace(placeholder, str(row_val))
                                cell.value = new_val

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("✨ エクセルの生成に成功しました！デザインも完全維持されています！")
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
    st.header("テンプレートごとのプレースホルダー紐づけ設定")
    st.write("エクセルに書き込んだプレースホルダー（例: `#field-36`）と @pocket のフィールドキーの対応を管理できます。")
    
    edit_template = st.selectbox(
        "編集するテンプレートを選択",
        options=template_files,
        key="edit_select"
    )
    
    target_conf = mappings_data.get(edit_template, {"header": {}, "details": {"template_keyword": "#field-31_2", "columns": {}}})
    
    st.subheader("📌 ヘッダーのプレースホルダー設定")
    current_header = target_conf.get("header", {})
    updated_header = {}
    headers_to_delete = []
    
    for placeholder, pocket_key in list(current_header.items()):
        c1, c2, c3 = st.columns([3, 3, 1])
        with c1:
            new_p = st.text_input("プレースホルダー", value=placeholder, key=f"hp_{edit_template}_{placeholder}")
        with c2:
            new_k = st.text_input("@pocketキー", value=pocket_key, key=f"hk_{edit_template}_{placeholder}")
        with c3:
            if st.button("🗑️ 削除", key=f"hdel_{edit_template}_{placeholder}"):
                headers_to_delete.append(placeholder)
        if new_p and new_p not in headers_to_delete:
            updated_header[new_p] = new_k

    for k in headers_to_delete:
        if k in updated_header: del updated_header[k]
        
    st.write("➕ ヘッダー項目の追加")
    np_field = st.text_input("プレースホルダー (例: #field-38)", key=f"np_{edit_template}")
    nk_field = st.text_input("@pocketキー (例: ClassD)", key=f"nk_{edit_template}")
    if np_field:
        updated_header[np_field] = nk_field

    st.divider()
    st.subheader("📌 明細行のプレースホルダー設定")
    current_details = target_conf.get("details", {"template_keyword": "#field-31_2", "columns": {}})
    
    template_kw = st.text_input("明細行検出キーワード (例: #field-31_2)", value=current_details.get("template_keyword", "#field-31_2"))
    
    st.write("明細列の紐づけ（プレースホルダー と @pocketキー）")
    current_cols = current_details.get("columns", {})
    updated_cols = {}
    cols_to_delete_d = []
    
    for placeholder, pocket_key in list(current_cols.items()):
        c1, c2, c3 = st.columns([3, 3, 1])
        with c1:
            dp = st.text_input("明細プレースホルダー", value=placeholder, key=f"dp_{edit_template}_{placeholder}")
        with c2:
            dk = st.text_input("@pocketキー", value=pocket_key, key=f"dk_{edit_template}_{placeholder}")
        with c3:
            if st.button("🗑️ 削除", key=f"ddel_{edit_template}_{placeholder}"):
                cols_to_delete_d.append(placeholder)
        if dp and dp not in cols_to_delete_d:
            updated_cols[dp] = dk

    for k in cols_to_delete_d:
        if k in updated_cols: del updated_cols[k]

    st.write("➕ 明細列の追加")
    nd_p = st.text_input("明細プレースホルダー (例: #field-31_5)", key=f"ndp_{edit_template}")
    nd_k = st.text_input("@pocketキー (例: Spec)", key=f"ndk_{edit_template}")
    if nd_p:
        updated_cols[nd_p] = nd_k

    st.divider()
    if st.button("💾 すべての設定を保存する", type="primary"):
        mappings_data[edit_template] = {
            "header": updated_header,
            "details": {
                "template_keyword": template_kw,
                "columns": updated_cols
            }
        }
        save_mappings(mappings_data)
        st.success("✅ マッピング設定を保存しました！")
        st.rerun()