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
    if not