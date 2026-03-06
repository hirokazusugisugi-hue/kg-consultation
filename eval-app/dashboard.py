"""
コンサルタント評価システム（管理者専用）

Usage:
    streamlit run eval-app/dashboard.py
"""

import json
import os
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from dotenv import load_dotenv

from config.settings import CATEGORIES, HUMAN_ITEMS, ITEM_NAMES, AI_MAX, TOTAL_MAX, NG_CATEGORIES

load_dotenv()

# ── App Config ──
st.set_page_config(
    page_title="コンサルタント評価システム",
    page_icon="📋",
    layout="wide",
)

RESULTS_DIR = Path("eval-app/results")
RESULTS_DIR.mkdir(parents=True, exist_ok=True)


# ── Shared Helpers ──

def load_results():
    files = sorted(RESULTS_DIR.glob("*.json"))
    data = []
    for f in files:
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            d["_file"] = f.name
            data.append(d)
        except Exception:
            pass
    return data


def render_radar_chart(result):
    """カテゴリ別レーダーチャートを描画"""
    cat_keys = ["c1", "c2", "c3", "c4", "c5", "c6"]
    cat_labels = [CATEGORIES[k]["name"] for k in cat_keys]
    scores = result.get("category_scores", {})
    cat_values = [scores.get(k, 0) for k in cat_keys]
    cat_max = [CATEGORIES[k]["max_raw"] for k in cat_keys]
    cat_pct = [v / m * 5 if m > 0 else 0 for v, m in zip(cat_values, cat_max)]

    fig = go.Figure(data=go.Scatterpolar(
        r=cat_pct + [cat_pct[0]],
        theta=cat_labels + [cat_labels[0]],
        fill="toself",
        fillcolor="rgba(15,35,80,0.15)",
        line=dict(color="#0F2350", width=2),
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
        showlegend=False,
        height=400,
        margin=dict(l=60, r=60, t=30, b=30),
    )
    return fig


def render_item_table(result):
    """22項目スコア表を描画"""
    item_data = []
    item_scores = result.get("item_scores", {})
    evidence = result.get("evidence", {})
    for num in range(1, 23):
        num_str = str(num)
        score = item_scores.get(num_str, item_scores.get(num, 0))
        ev = evidence.get(num_str, evidence.get(str(num), {}))
        # カテゴリ特定
        cat_name = ""
        for ck, cv in CATEGORIES.items():
            if num in cv["items"]:
                cat_name = cv["name"]
                break
        item_data.append({
            "No.": num,
            "カテゴリ": cat_name,
            "項目名": ITEM_NAMES.get(num, ""),
            "スコア": score,
            "根拠": ev.get("reasoning", "")[:80] if isinstance(ev, dict) else "",
        })
    return pd.DataFrame(item_data)


def render_result_detail(result):
    """評価結果の詳細表示"""
    ai_total = result.get("ai_total", 0)
    raw_total = result.get("raw_total", 0)

    # スコアサマリー
    col1, col2, col3 = st.columns(3)
    with col1:
        color = "normal" if ai_total >= 60 else "off"
        st.metric("AI評価スコア", f"{ai_total}/90", delta=None)
    with col2:
        st.metric("素点合計", f"{raw_total}/110")
    with col3:
        ng_count = len(result.get("ng_words", []))
        st.metric("NG語句検出数", ng_count)

    # レーダーチャート
    st.plotly_chart(render_radar_chart(result), use_container_width=True)

    # 22項目テーブル
    st.subheader("22項目スコア一覧")
    df = render_item_table(result)
    st.dataframe(df, use_container_width=True, hide_index=True)

    # NG語句
    ng_words = result.get("ng_words", [])
    if ng_words:
        st.subheader("NG語句検出")
        for ng in ng_words:
            cat = ng.get("category", "?")
            cat_desc = NG_CATEGORIES.get(cat, cat)
            st.warning(f"**[{cat}] {cat_desc}**: {ng.get('text', '')}  \n文脈: {ng.get('context', '')}")

    # メタデータ
    meta = result.get("metadata", {})
    if meta:
        with st.expander("メタデータ"):
            st.json(meta)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ページ: 評価実行
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def page_evaluate():
    st.title("評価実行")
    st.caption("文字起こしテキストを入力し、コンサルタントのAI評価を実行します")

    # 実行モード設定
    with st.sidebar:
        st.subheader("実行設定")
        mode = st.radio(
            "実行モード",
            ["ローカル（直接API呼出）", "Cloud Function経由"],
            help="ローカル: ANTHROPIC_API_KEYが必要。CF: デプロイ済みCF URLが必要。",
        )

        if mode == "Cloud Function経由":
            cf_url = st.text_input(
                "CF URL",
                value=os.environ.get("EVALUATION_CF_URL", ""),
                type="default",
            )
            cf_secret = st.text_input(
                "CF Secret",
                value=os.environ.get("EVALUATION_CF_SECRET", ""),
                type="password",
            )
        else:
            api_key = os.environ.get("ANTHROPIC_API_KEY", "")
            if api_key:
                st.success("ANTHROPIC_API_KEY: 設定済み")
            else:
                st.error("ANTHROPIC_API_KEY が未設定です。.envファイルに設定してください。")

    # メタデータ入力
    st.subheader("相談情報")
    col1, col2 = st.columns(2)
    with col1:
        consultant_name = st.text_input("コンサルタント名（リーダー）", placeholder="例: 山田太郎")
        company_name = st.text_input("相談企業名", placeholder="例: 株式会社サンプル")
    with col2:
        industry = st.text_input("業種", placeholder="例: 製造業")
        theme = st.text_input("相談テーマ", placeholder="例: 売上拡大策")

    consultation_type = st.selectbox("相談形式", ["対面", "Zoom", "その他"])

    # 入力データ
    st.subheader("入力データ")
    input_method = st.radio(
        "入力方法",
        ["音声ファイル", "テキスト直接入力", "テキストファイル"],
        horizontal=True,
    )

    transcript = ""
    audio_bytes = None
    audio_ext = ".wav"

    if input_method == "音声ファイル":
        st.info(
            "音声ファイルをアップロードすると、**自動で文字起こし→AI評価**を実行します。\n\n"
            "対応形式: WAV, MP3, M4A  |  推奨: 30-90分の相談録音  |  "
            "前提: `brew install ffmpeg` と `pip install faster-whisper`"
        )

        audio_tab, record_tab = st.tabs(["ファイルアップロード", "ブラウザ録音"])

        with audio_tab:
            audio_file = st.file_uploader(
                "音声ファイル (.wav, .mp3, .m4a)",
                type=["wav", "mp3", "m4a", "mp4", "ogg", "flac"],
            )
            if audio_file:
                file_size_mb = len(audio_file.getvalue()) / (1024 * 1024)
                st.caption(f"ファイル: {audio_file.name} ({file_size_mb:.1f} MB)")
                audio_bytes = audio_file.getvalue()
                audio_ext = "." + audio_file.name.rsplit(".", 1)[-1].lower()

        with record_tab:
            audio_recorded = st.audio_input(
                "録音（短時間テスト向け）",
            )
            if audio_recorded and not audio_bytes:
                audio_bytes = audio_recorded.getvalue()
                audio_ext = ".wav"
                st.caption(f"録音データ: {len(audio_bytes) / 1024:.0f} KB")

    elif input_method == "テキスト直接入力":
        transcript = st.text_area(
            "文字起こしテキスト",
            height=300,
            placeholder="相談の文字起こしテキストをここに貼り付けてください...\n\n"
                        "【話者1】こんにちは、本日はよろしくお願いします。\n"
                        "【話者2】よろしくお願いします。早速ですが...",
        )

    else:  # テキストファイル
        uploaded = st.file_uploader(
            "文字起こしファイル (.txt)",
            type=["txt"],
            help="UTF-8のテキストファイルをアップロードしてください",
        )
        if uploaded:
            transcript = uploaded.read().decode("utf-8")
            st.info(f"読み込み完了: {len(transcript):,}文字")

    # テキスト統計（テキスト入力時のみ）
    if transcript:
        word_count = len(transcript)
        st.caption(f"テキスト長: {word_count:,}文字 / 上限: {120,000:,}文字")
        if word_count > 120000:
            st.warning("テキストが上限を超えています。自動的に切り詰められます。")

    # ── 実行 ──
    st.divider()

    # 実行可否判定
    if input_method == "音声ファイル":
        can_run = audio_bytes is not None
    else:
        can_run = bool(transcript.strip())

    if mode == "ローカル（直接API呼出）":
        can_run = can_run and bool(os.environ.get("ANTHROPIC_API_KEY"))
    else:
        can_run = can_run and bool(cf_url)

    button_label = "文字起こし + AI評価を自動実行" if input_method == "音声ファイル" else "評価を実行"

    if st.button(button_label, type="primary", disabled=not can_run, use_container_width=True):
        metadata = {
            "consultant_name": consultant_name,
            "company_name": company_name,
            "industry": industry,
            "theme": theme,
            "consultation_type": consultation_type,
            "evaluated_at": datetime.now().isoformat(),
            "input_type": "audio" if input_method == "音声ファイル" else "text",
        }

        try:
            from modules.evaluator import evaluate_local, evaluate_via_cf

            # ── Step 1: 音声→文字起こし（音声入力時のみ）──
            if input_method == "音声ファイル":
                from modules.transcriber import transcribe_uploaded_bytes

                st.subheader("Step 1 / 2: 文字起こし")
                transcription_status = st.empty()
                transcription_progress = st.progress(0, text="Whisperモデルをロード中...")

                start_time = time.time()

                def on_transcribe_progress(message):
                    transcription_status.info(message)

                transcription_progress.progress(0.1, text="文字起こしを実行中...")

                transcript = transcribe_uploaded_bytes(
                    audio_bytes=audio_bytes,
                    file_extension=audio_ext,
                    progress_callback=on_transcribe_progress,
                )

                elapsed = time.time() - start_time
                transcription_progress.progress(1.0, text=f"文字起こし完了（{elapsed:.0f}秒）")
                st.success(f"文字起こし完了: {len(transcript):,}文字 / {elapsed:.0f}秒")

                with st.expander("文字起こし結果を確認"):
                    st.text_area("文字起こしテキスト", transcript, height=200, disabled=True)

                metadata["transcript_chars"] = len(transcript)
                metadata["transcription_time_sec"] = round(elapsed, 1)

                step_label = "Step 2 / 2: AI評価"
            else:
                step_label = "AI評価"

            # ── Step 2: AI評価 ──
            if input_method == "音声ファイル":
                st.subheader(step_label)

            if mode == "ローカル（直接API呼出）":
                progress_bar = st.progress(0, text="評価を開始しています...")

                def on_progress(current, total, message):
                    progress_bar.progress(current / total if total > 0 else 0, text=message)

                with st.spinner("Claude APIで評価中...（約2-3分）"):
                    result = evaluate_local(transcript, progress_callback=on_progress)

                progress_bar.progress(1.0, text="評価完了")

            else:
                with st.spinner("Cloud Functionで評価中...（約2-3分）"):
                    result = evaluate_via_cf(
                        transcript=transcript,
                        cf_url=cf_url,
                        cf_secret=cf_secret,
                        metadata=metadata,
                    )

            # メタデータを結果に追加
            result["metadata"] = metadata

            # 結果保存
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = (consultant_name or "unknown").replace(" ", "_")
            filename = f"{timestamp}_{safe_name}.json"
            result_path = RESULTS_DIR / filename
            result_path.write_text(
                json.dumps(result, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

            st.success(f"評価完了！結果を保存しました: {filename}")

            # 結果表示
            st.divider()
            render_result_detail(result)

        except ImportError as e:
            st.error(
                f"モジュールのインポートに失敗しました: {e}\n\n"
                "以下をインストールしてください:\n"
                "```\npip install faster-whisper\nbrew install ffmpeg\n```"
            )
        except Exception as e:
            st.error(f"実行エラー: {e}")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ページ: ダッシュボード
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def page_dashboard():
    st.title("評価ダッシュボード")
    st.caption("ICMCI CMC・Schein理論・SERVQUAL・MITIに基づく22項目+3項目の学術的評価")

    results = load_results()

    if not results:
        st.info("評価結果がありません。「評価実行」ページから評価を実行してください。")
        return

    # サマリー
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("総評価数", len(results))
    with col2:
        avg = sum(r.get("ai_total", 0) for r in results) / len(results)
        st.metric("平均AIスコア", f"{avg:.1f}/90")
    with col3:
        max_score = max(r.get("ai_total", 0) for r in results)
        st.metric("最高スコア", f"{max_score}/90")
    with col4:
        min_score = min(r.get("ai_total", 0) for r in results)
        st.metric("最低スコア", f"{min_score}/90")

    # 評価結果一覧
    st.subheader("評価結果一覧")
    table_data = []
    for r in reversed(results):
        meta = r.get("metadata", {})
        table_data.append({
            "日時": meta.get("evaluated_at", r.get("_file", ""))[:16],
            "コンサルタント": meta.get("consultant_name", "-"),
            "企業名": meta.get("company_name", "-"),
            "形式": meta.get("consultation_type", "-"),
            "AIスコア": r.get("ai_total", 0),
            "素点": r.get("raw_total", 0),
            "NG語句": len(r.get("ng_words", [])),
            "ファイル": r.get("_file", ""),
        })

    if table_data:
        df_list = pd.DataFrame(table_data)
        st.dataframe(df_list, use_container_width=True, hide_index=True)

    # 個別結果の詳細表示
    st.subheader("詳細表示")
    file_options = [r.get("_file", f"result_{i}") for i, r in enumerate(results)]
    selected_file = st.selectbox("評価結果を選択", reversed(file_options))

    selected = None
    for r in results:
        if r.get("_file") == selected_file:
            selected = r
            break

    if selected:
        render_result_detail(selected)

    # コンサルタント別比較
    consultants = {}
    for r in results:
        name = r.get("metadata", {}).get("consultant_name", "不明")
        if name not in consultants:
            consultants[name] = []
        consultants[name].append(r.get("ai_total", 0))

    if len(consultants) > 1:
        st.subheader("コンサルタント別平均スコア")
        comp_data = []
        for name, scores in consultants.items():
            comp_data.append({
                "コンサルタント": name,
                "評価回数": len(scores),
                "平均AIスコア": round(sum(scores) / len(scores), 1),
                "最高": max(scores),
                "最低": min(scores),
            })
        st.dataframe(pd.DataFrame(comp_data), use_container_width=True, hide_index=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ページ: 設定
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def page_settings():
    st.title("設定・セットアップ")

    st.subheader("現在の設定状況")

    # 環境変数チェック
    checks = {
        "ANTHROPIC_API_KEY": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "EVALUATION_CF_URL": bool(os.environ.get("EVALUATION_CF_URL")),
        "EVALUATION_CF_SECRET": bool(os.environ.get("EVALUATION_CF_SECRET")),
    }

    for key, ok in checks.items():
        if ok:
            st.success(f"{key}: 設定済み")
        else:
            st.warning(f"{key}: 未設定")

    # プロンプトファイルチェック
    st.subheader("プロンプトファイル")
    prompts_dir = Path(__file__).parent.parent / "cloud_functions" / "consultation_evaluation" / "prompts"
    prompt_files = ["system.txt"] + [f"call{i}_{n}.txt" for i, n in enumerate(
        ["problem", "solution", "communication", "time", "logic", "ethics"], 1
    )]
    for pf in prompt_files:
        path = prompts_dir / pf
        if path.exists():
            size = path.stat().st_size
            st.success(f"{pf}: {size:,} bytes")
        else:
            st.error(f"{pf}: 見つかりません")

    # 音声文字起こし
    st.subheader("音声文字起こし")
    try:
        import faster_whisper
        st.success(f"faster-whisper: インストール済み")
    except ImportError:
        st.warning("faster-whisper: 未インストール（`pip install faster-whisper`）")

    import subprocess
    try:
        ffmpeg_result = subprocess.run(
            ["ffmpeg", "-version"], capture_output=True, text=True, timeout=5
        )
        if ffmpeg_result.returncode == 0:
            version_line = ffmpeg_result.stdout.split("\n")[0]
            st.success(f"FFmpeg: {version_line}")
        else:
            st.warning("FFmpeg: 動作確認に失敗")
    except FileNotFoundError:
        st.error("FFmpeg: 未インストール（`brew install ffmpeg`）")
    except Exception:
        st.warning("FFmpeg: 確認できません")

    whisper_model = os.environ.get("WHISPER_MODEL", "large-v3")
    st.info(f"Whisperモデル: {whisper_model}（WHISPER_MODEL環境変数で変更可能）")

    # 結果ディレクトリ
    st.subheader("結果データ")
    result_files = list(RESULTS_DIR.glob("*.json"))
    st.info(f"保存済み評価結果: {len(result_files)}件 ({RESULTS_DIR})")

    # .envテンプレート
    st.subheader(".env 設定テンプレート")
    st.code(
        "# eval-app/.env に配置\n"
        "ANTHROPIC_API_KEY=sk-ant-xxxxx\n"
        "EVALUATION_CF_URL=https://xxx.cloudfunctions.net/consultation_evaluation\n"
        "EVALUATION_CF_SECRET=your-shared-secret\n\n"
        "# 音声文字起こし（large-v3 or medium）\n"
        "WHISPER_MODEL=large-v3\n",
        language="bash",
    )

    st.subheader("Cloud Functionデプロイコマンド")
    st.code(
        "cd cloud_functions/consultation_evaluation\n\n"
        "gcloud functions deploy consultation_evaluation \\\n"
        "  --gen2 --runtime python311 --trigger-http \\\n"
        "  --allow-unauthenticated \\\n"
        "  --timeout 600 --memory 1024MB \\\n"
        "  --set-env-vars SHARED_SECRET=xxx,ANTHROPIC_API_KEY=sk-ant-xxx",
        language="bash",
    )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# メインルーティング
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PAGES = {
    "評価実行": page_evaluate,
    "ダッシュボード": page_dashboard,
    "設定": page_settings,
}

with st.sidebar:
    st.title("コンサルタント評価")
    st.caption("管理者専用ツール")
    st.divider()
    page = st.radio("ページ", list(PAGES.keys()), label_visibility="collapsed")

PAGES[page]()
