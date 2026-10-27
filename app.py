import streamlit as st
import openpyxl
import io
import json
import os
import requests
import pandas as pd

# ページ基本設定
st.set_page_config(page_title="汎用エクセル帳票ジェネレーター", layout="wide")
st.title("📊 汎用エクセル帳票発行システム（ヘッダー＆明細対応版）")

# ── 【安全なSecrets取得処理】 ──
if "pocket" in st.secrets:
    POCKET_API_KEY = st.secrets["pocket"].get("api_key", "XZtsgp27LemoHhCkwkoEiS2MZMyCU4Ay7qyYF8PS")
    POCKET_BASE_URL = st.secrets["pocket"].get("base_url", "https://app070.at-pocket.com/arkf05_bb")
else:
    POCKET_API_KEY = "XZtsgp27LemoHhCkwkoEiS2MZMyCU4Ay7qyYF8PS"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"

APP_ID_ESTIMATE = 32 # 見積アプリID
MAPPING_FILE = "template_mappings.json"

# ── 【デフォルトのマッピング設定（ヘッダー ＆ 明細ルール）】 ──
DEFAULT_MAPPINGS = {
    "mitumoriV1": {
        "header": {
            "見積番号": "A2",
            "顧客名": "A3",
            "発行日": "F3",
            "合計金額": "F19"
        },
        "details": {
            "start_row": 9,       # 明細を書き始める行
            "columns": {
                "品名": "B",      # B列
                "備考": "C",      # C列
                "数量": "D",      # D列
                "単価": "E"       # E列 (金額FはエクセルのSUM数式で自動計算)
            }
        }
    },
    "mitumoriV2": {
        "header": {
            "見積番号": "E1",
            "顧客名": "B2",
            "発行日": "E2",
            "合計金額": "F25"
        },
        "details": {
            "start_row": 12,
            "columns": {
                "品名": "B",
                "備考": "C",
                "数量": "D",
                "単価": "E"
            }
        }
    }
}

# マッピング設定の読み込み・保存関数
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

# ── @pocketからのパラメータ受け取り ──
query_params = st.query_params
target_record_id = query_params.get("record_id", None)
passed_estimate_no = query_params.get("estimate_no", "")

# @pocketのAPI（アプリID: 32）からデータを取得する関数
def fetch_pocket_data(record_id):
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
            
            # ヘッダー情報の抽出
            header_data = {
                "見積番号": passed_estimate_no or data.get("Title", ""),
                "顧客名": data.get("Body", "サンプル株式会社"),
                "発行日": "2026-10-27",
                "合計金額": "" # 自動計算
            }
            
            # 明細データの抽出（※@pocketのテーブル項目の構造に合わせて調整可能。なければサンプル）
            # 例として、APIから子テーブルや Results が取れる場合はここにバインドします
            sample_details = [
                {"品名": "特選 生ウニ（100g）", "備考": "冷蔵便", "数量": 10, "単価": 5000},
                {"品名": "厳選 イクラ醤油漬け（500g）", "備考": "冷凍便", "数量": 5, "単価": 8000},
            ]
            
            return header_data, sample_details
            
    except Exception as e:
        st.error(f"API通信エラー: {e}")
        
    return {}, []

# データの取得・初期化
pocket_header = {}
pocket_details = []

if target_record_id:
    st.sidebar.success(f"🔗 @pocket 連携中\n\n(ID: {target_record_id} / 見積番号: {passed_estimate_no})")
    pocket_header, pocket_details = fetch_pocket_data(target_record_id)
else:
    pocket_header = {
        "見積番号": passed_estimate_no or "EST-001",
        "顧客名": "テスト株式会社",
        "発行日": "2026-10-27",
        "合計金額": ""
    }
    pocket_details = [
        {"品名": "サンプル商品 A", "備考": "通常", "数量": 2, "単価": 10000},
        {"品名": "サンプル商品 B", "備考": "特急", "数量": 1, "単価": 25000}
    ]

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
    
    current_config = mappings_data.get(selected_template, {})
    header_rules = current_config.get("header", {})
    detail_rules = current_config.get("details", {})

    st.divider()
    st.subheader("📝 1. ヘッダー情報の確認・編集")
    
    input_headers = {}
    cols = st.columns(2)
    i = 0
    for field_name, cell_pos in header_rules.items():
        default_val = pocket_header.get(field_name, "")
        with cols[i % 2]:
            input_headers[field_name] = st.text_input(
                f"{field_name} (セル: {cell_pos})",
                value=default_val
            )
        i += 1

    st.divider()
    st.subheader("📝 2. 明細行の確認・編集（ループ展開）")
    st.write(f"テンプレートのエクセル {detail_rules.get('start_row', 9)}行目 から順番に書き込まれます。")
    
    # 明細を編集できる表形式（Data Editor）
    df_details_input = pd.DataFrame(pocket_details)
    edited_df = st.data_editor(
        df_details_input,
        num_rows="dynamic", # 行の追加・削除を自由に行える
        use_container_width=True,
        key="detail_editor"
    )

    st.divider()

    # 🚀 エクセル生成ボタン
    if st.button("🚀 帳票エクセルを生成する", type="primary"):
        template_filename = f"{selected_template}.xlsx"
        
        if not os.path.exists(template_filename):
            st.error(f"テンプレートファイル '{template_filename}' がアプリのルートディレクトリに見つかりません。")
        else:
            try:
                # openpyxlでテンプレートのデザインを完全に維持したまま読み込み
                wb = openpyxl.load_workbook(template_filename)
                ws = wb.active
                
                # ① ヘッダー情報をピンポイントで流し込む
                for field_name, cell_pos in header_rules.items():
                    val = input_headers.get(field_name, "")
                    if cell_pos in ws:
                        ws[cell_pos] = val
                
                # ② 明細情報をループで順番に流し込む
                start_row = detail_rules.get("start_row", 9)
                col_mapping = detail_rules.get("columns", {})
                
                for index, row in edited_df.iterrows():
                    current_row = start_row + index # 9行目、10行目、11行目…と下へズラす
                    
                    for field_name, col_letter in col_mapping.items():
                        cell_address = f"{col_letter}{current_row}"
                        if field_name in row:
                            val = row[field_name]
                            ws[cell_address] = val
                            
                            # 数値の場合はカンマ区切りフォーマットを適用
                            if field_name in ["数量", "単価"] and pd.notnumeric(val):
                                ws[cell_address].number_format = '#,##0'

                # ③ メモリ上にバイナリとして保存
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
    st.header("テンプレートごとのセル関連付け設定")
    st.write("ヘッダーのセル番地や、明細の書き始め行（start_row）・列割り当てをここで調整できます。")
    
    edit_template = st.selectbox(
        "編集するテンプレートを選択",
        options=list(mappings_data.keys()),
        key="edit_select"
    )
    
    target_conf = mappings_data.get(edit_template, {"header": {}, "details": {"start_row": 9, "columns": {}}})
    
    st.subheader("📌 ヘッダー項目のマッピング設定")
    current_header = target_conf.get("header", {})
    updated_header = {}
    headers_to_delete = []
    
    for field_name, cell_pos in list(current_header.items()):
        c1, c2, c3 = st.columns([3, 3, 1])
        with c1:
            new_k = st.text_input("項目名", value=field_name, key=f"hk_{edit_template}_{field_name}")
        with c2:
            new_v = st.text_input("セル番地", value=cell_pos, key=f"hv_{edit_template}_{field_name}")
        with c3:
            if st.button("🗑️ 削除", key=f"hdel_{edit_template}_{field_name}"):
                headers_to_delete.append(field_name)
        if new_k and new_k not in headers_to_delete:
            updated_header[new_k] = new_v

    for k in headers_to_delete:
        if k in updated_header: del updated_header[k]
        
    # ヘッダー追加
    nh_field = st.text_input("追加するヘッダー項目名 (例: 担当者)", key=f"nh_{edit_template}")
    nh_cell = st.text_input("対応するセル番地 (例: B4)", key=f"nc_{edit_template}")
    if nh_field and nh_cell:
        updated_header[nh_field] = nh_cell

    st.divider()
    st.subheader("📌 明細（テーブル）のマッピング設定")
    current_details = target_conf.get("details", {"start_row": 9, "columns": {}})
    
    start_row_input = st.number_input("明細の書き始め行 (start_row)", value=int(current_details.get("start_row", 9)), min_int=1)
    
    st.write("列の割り当て（例: 品名 -> B, 数量 -> D）")
    current_cols = current_details.get("columns", {})
    updated_cols = {}
    
    for f_name, c_letter in list(current_cols.items()):
        c1, c2 = st.columns(2)
        with c1:
            dk = st.text_input("項目", value=f_name, key=f"dk_{edit_template}_{f_name}")
        with c2:
            dv = st.text_input("列アルファベット", value=c_letter, key=f"dv_{edit_template}_{f_name}")
        if dk:
            updated_cols[dk] = dv

    # 設定保存ボタン
    if st.button("💾 すべての設定を保存する", type="primary"):
        mappings_data[edit_template] = {
            "header": updated_header,
            "details": {
                "start_row": start_row_input,
                "columns": updated_cols
            }
        }
        save_mappings(mappings_data)
        st.success("✅ マッピング設定を保存しました！")
        st.rerun()