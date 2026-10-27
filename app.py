# ==========================================
# タブ1：ダウンロード画面
# ==========================================
with tab1:
  st.header("帳票の生成とダウンロード")

  # ── 【ここを追加】@pocketからの取得状況を常時見える化するデバッグボックス ──
  with st.expander(
      "🔍 【デバッグ】@pocketからのデータ取得状況を確認する", expanded=True
  ):
    st.write(f"- **受け取った レコードID (record_id):** `{target_record_id}`")
    st.write(f"- **受け取った 見積番号 (estimate_no):** `{passed_estimate_no}`")

    if target_record_id:
      st.success(
          "✅ @pocketからのレコードID検知：API通信を試行しています。"
      )
    else:
      st.warning(
          "⚠️ レコードIDが検知されません（※現在は「テスト用ダミーデータ」で動作しています）。"
          " @pocketのボタンから開いていない可能性があります。"
      )

    st.write("**▼ @pocketから返ってきた実際の生データ (raw_pocket_data):**")
    st.json(raw_pocket_data)

  template_files = [
      f.replace(".xlsx", "")
      for f in os.listdir(".")
      if f.endswith(".xlsx") and not f.startswith("~")
  ]
  if not template_files:
    template_files = list(mappings_data.keys())

  selected_template = st.selectbox(
      "使用するエクセルテンプレートを選択", options=template_files
  )

  current_config = mappings_data.get(
      selected_template, {"header": {}, "details": {}}
  )
  header_rules = current_config.get("header", {})
  detail_rules = current_config.get("details", {})

  st.divider()
  st.subheader("📝 1. ヘッダー情報の確認・編集")

  input_headers = {}
  cols = st.columns(2)
  i = 0
  for placeholder, pocket_key in header_rules.items():
    # データを正しく取得できているか確認
    fetched_val = raw_pocket_data.get(pocket_key, "")
    default_val = (
        passed_estimate_no
        if placeholder == "field-2"
        else (fetched_val if fetched_val != "" else f"({pocket_key}が未取得)")
    )

    with cols[i % 2]:
      input_headers[placeholder] = st.text_input(
          f"{placeholder} (← @pocketキー: {pocket_key})",
          value=str(default_val),
      )
    i += 1