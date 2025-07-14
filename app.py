import os
from pathlib import Path
import base64
from typing import List, Dict, Union, Any
# 非同期利用しないためコメントアウト
# import asyncio
import tempfile
import logging
import json

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.exceptions import CellCoordinatesException
from io import BytesIO

from langchain_openai import ChatOpenAI, AzureChatOpenAI
from langchain_core.messages import HumanMessage
import PyPDF2
from langchain_core.pydantic_v1 import BaseModel, Field, create_model

from excel_format_analyzer import ExcelFormatAnalyzer

# --------------------------------------------------------------------------------------
# 環境変数ロード
# --------------------------------------------------------------------------------------
load_dotenv()

# --------------------------------------------------------------------------------------
# ログ設定
# --------------------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# --------------------------------------------------------------------------------------
# 定数・パス
# --------------------------------------------------------------------------------------
DATA_DIR = Path("data")
BATCH_DIR = DATA_DIR / "batch"
FORMAT_DIR = DATA_DIR / "format"
TEMP_DIR = FORMAT_DIR / "temp"

MODEL_NAME = "gpt-4.1"
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2023-07-01-preview")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", MODEL_NAME)

# --------------------------------------------------------------------------------------
# 構造化出力用モデル
# --------------------------------------------------------------------------------------
class CellUpdate(BaseModel):
    """個別のセル更新情報を表すモデル"""
    cell_id: str = Field(description="更新対象のセル番号 (例: C5)")
    content: Union[str, int, float, None] = Field(description="セルに書き込む内容")

class CommonFieldSummary(BaseModel):
    """調書の共通項目に記載する内容のデータモデル"""
    updates: List[CellUpdate] = Field(description="共通項目に対する更新内容のリスト")


# --------------------------------------------------------------------------------------
# ユーティリティ関数
# --------------------------------------------------------------------------------------

def create_chat_model() -> ChatOpenAI:
    """OpenAI または Azure OpenAI のどちらかを初期化して返す"""
    if AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY:
        return AzureChatOpenAI(
            api_key=AZURE_OPENAI_API_KEY,
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            openai_api_version=AZURE_OPENAI_API_VERSION,
            azure_deployment=AZURE_OPENAI_DEPLOYMENT,
            temperature=0,
        )
    return ChatOpenAI(model=MODEL_NAME, api_key=OPENAI_API_KEY, temperature=0)

def list_batches():
    """data/batch 配下のバッチフォルダ名を取得"""
    if not BATCH_DIR.exists():
        return []
    return sorted([p.name for p in BATCH_DIR.iterdir() if p.is_dir()])


def list_templates():
    """data/format 配下のテンプレートファイルを取得"""
    if not FORMAT_DIR.exists():
        return []
    return sorted([p.name for p in FORMAT_DIR.iterdir() if p.is_file()])


def read_sample(file_path: Path) -> Union[str, Dict[str, str]]:
    """サンプルファイルを読み取りテキスト化、または画像をbase64エンコード"""
    suffix = file_path.suffix.lower()
    try:
        if suffix in [".txt", ".csv"]:
            return file_path.read_text(encoding="utf-8", errors="ignore")
        elif suffix == ".pdf" and PyPDF2:
            text = []
            with open(file_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)  # type: ignore
                for page in reader.pages:
                    text.append(page.extract_text() or "")  # type: ignore
            return "\n".join(text)
        elif suffix in [".png", ".jpg", ".jpeg", ".gif"]:
            image_data = file_path.read_bytes()
            base64_image = base64.b64encode(image_data).decode("utf-8")
            if suffix == ".jpg":
                mime_type = "image/jpeg"
            else:
                mime_type = f"image/{suffix[1:]}"
            return {"type": "image", "data": base64_image, "mime_type": mime_type}
        else:
            return f"[INFO] {file_path.name} は画像／非対応フォーマットのため OCR 未実装"
    except Exception as e:
        return f"[ERROR] {file_path.name} の読み取りに失敗: {e}"


def generate_summary_for_common_fields(model: "ChatOpenAI", procedure: str, results: List[Dict[str, Any]], common_fields: Dict[str, Any]) -> Dict[str, str]:
    """LLM を使って共通項目に記載するサマリーを生成"""
    
    if not common_fields:
        return {}
        
    st.info("共通項目に記載する内容の生成を開始します...")

    # `results` をテキスト形式に変換
    # 動的なキーに対応するため、各サンプルの結果(JSON)をそのまま文字列にする
    results_text_parts = []
    for r in results:
        sample_name = r.get('sample', 'N/A')
        # 'result'キーに格納されている辞書（LLMからの構造化出力）をJSON文字列に変換
        result_details_json = json.dumps(r.get('result', {}), ensure_ascii=False, indent=2)
        results_text_parts.append(f"--- サンプル: {sample_name} ---\n{result_details_json}\n")
    
    results_text = "\\n".join(results_text_parts)
    
    # `common_fields` をテキスト形式に変換
    fields_text = "\\n".join([
        f"- {v['description']} (セル: {k})" for k, v in common_fields.items()
    ])

    example_json = """
{
  "updates": [
    {
      "cell_id": "C5",
      "content": "全サンプルを確認した結果、手続きは概ね良好に実施されている。"
    },
    {
      "cell_id": "C6",
      "content": "一部のサンプルで軽微な不備が見られたが、全体として大きな問題はない。"
    }
  ]
}
"""

    prompt = f"""
以下の監査手続きと、サンプルごとのテスト結果（JSON形式）に基づき、調書の共通項目に記載する内容を生成してください。

# 監査手続き
{procedure}

# テスト結果詳細
{results_text}

# 記載が必要な共通項目
{fields_text}

# 指示
- 各共通項目の説明に沿った内容を、テスト結果全体を要約して記述してください。
- 記載に必要な情報が足りない場合はN/Aと記載してください。
- 出力は 'updates' というキーを持つJSONオブジェクトとしてください。
- 'updates' の値は、各セルへの更新情報を格納したオブジェクトのリストです。
- 各オブジェクトは 'cell_id' (セル番号) と 'content' (記載内容) のキーを持ちます。
- 以下の例のようなJSON形式で出力してください。

例:
```json
{example_json}
```
"""

    try:
        structured_llm = model.with_structured_output(CommonFieldSummary)
        response = structured_llm.invoke([HumanMessage(content=prompt)])
        # List[CellUpdate] から {"C5": "..."} 形式の辞書に変換
        summary_contents = {update.cell_id: update.content for update in response.updates}
        st.success("共通項目の内容生成が完了しました。")
        return summary_contents
    except Exception as e:
        st.error(f"共通項目の内容生成に失敗しました: {e}")
        import traceback
        st.error(traceback.format_exc())
        return {}

def fill_excel_with_results(
    template_path: str,
    output_path: str,
    analysis_result: Dict[str, Any],
    test_results: List[Dict[str, Any]],
    procedure: str,
    model: "ChatOpenAI"
) -> bool:
    """
    分析結果とテスト結果を元に調書(Excel)を更新する（ルールベース）
    """
    try:
        st.info("調書への結果転記を開始します...")
        workbook = openpyxl.load_workbook(template_path)
        # TODO: 複数シートに対応する場合は、どのシートを対象にするか決定するロジックが必要
        sheet: Worksheet = workbook.active
        
        common_fields = analysis_result.get("common_fields", {})
        sample_fields_info = analysis_result.get("sample_fields", {})
        
        # 1. 共通項目を更新（ここはLLMによる要約を利用）
        if common_fields:
            summary_contents = generate_summary_for_common_fields(model, procedure, test_results, common_fields)
            for cell_id, content in summary_contents.items():
                try:
                    sheet[cell_id] = content
                except (KeyError, ValueError, CellCoordinatesException) as e:
                    st.warning(f"共通項目のセル '{cell_id}' の更新に失敗しました: {e}")

        # 2. サンプル個別項目を更新（ルールベース）
        if sample_fields_info and "start_row" in sample_fields_info and "columns" in sample_fields_info:
            start_row = sample_fields_info["start_row"]
            # key_name をキー、col を値とするマップを作成
            columns_map = {key: info["col"] for key, info in sample_fields_info["columns"].items()}
               
            for i, result_data in enumerate(test_results):
                current_row = start_row + i
                # test_resultsの中身は 'sample' と 'result' (JSON文字列)
                # 'result' は pydantic model を dict に変換したもの
                test_result_json = result_data.get("result", {})
                
                for key_name, content_to_write in test_result_json.items():
                    if key_name in columns_map:
                        col_letter = columns_map[key_name]
                        cell_id = f"{col_letter}{current_row}"
                        try:
                            sheet[cell_id] = content_to_write
                        except (KeyError, ValueError, CellCoordinatesException) as e:
                            st.warning(f"サンプル個別のセル '{cell_id}' の更新に失敗しました: {e}")
                    else:
                        # sample_id のように結果JSONには含まれるが、調書の列にはない場合がある
                        # その逆（調書列はあるが、結果JSONにない）も考慮
                        logging.info(f"キー '{key_name}' に対応する列が調書フォーマットに見つかりません。転記をスキップします。")

        else:
            st.warning("サンプル個別項目の情報が不十分なため、更新をスキップしました。")


        workbook.save(output_path)
        st.success(f"調書の更新が完了しました: {output_path}")
        return True
    except Exception as e:
        st.error(f"調書の更新中にエラーが発生しました: {e}")
        import traceback
        st.error(traceback.format_exc())
        return False


def execute_audit_procedure(model: "ChatOpenAI", procedure: str, sample_contents: List[Union[str, Dict[str, Any]]], sample_format_keys: Dict[str, Any]) -> Dict[str, Any]:
    """LLM を使って監査手続きを実行し、構造化された結果を返す"""
    
    # sample_format_keys から動的にPydanticモデルを生成
    # 例: {"sample_id": (str, ...), "test_result": (str, ...)}
    fields_for_model = {
        key: (Union[str, int, float, None], Field(description=info.get("description", "")))
        for key, info in sample_format_keys.items()
    }
    DynamicAuditResult = create_model("DynamicAuditResult", **fields_for_model)

    prompt_messages = []
    
    # プロンプトの組み立て
    text_content = f"""
あなたは公認会計士です。以下の監査手続き指示書と関連資料に基づき、テストを実施してください。

# 監査手続き指示書
{procedure}

# 関連資料
"""
    prompt_messages.append({"type": "text", "text": text_content})
    
    for i, content in enumerate(sample_contents):
        if isinstance(content, dict) and content.get("type") == "image":
             prompt_messages.append({
                "type": "image_url",
                "image_url": {"url": f"data:{content['mime_type']};base64,{content['data']}"}
            })
        elif isinstance(content, str):
            # テキストコンテンツの場合
            text_part = f"--- 資料 {i+1} ---\n{content}\n---"
            prompt_messages.append({"type": "text", "text": text_part})


    final_instruction = """
# 指示
すべての資料を精査し、監査手続き指示書に従ってテストを実施してください。
結果は、提供されたフォーマットに従って、JSON形式で正確に報告してください。
"""
    prompt_messages.append({"type": "text", "text": final_instruction})

    try:
        structured_llm = model.with_structured_output(DynamicAuditResult)
        response = structured_llm.invoke([HumanMessage(content=prompt_messages)])
        # Pydanticモデルを辞書に変換して返す
        return response.dict()
    except Exception as e:
        logging.error(f"LLMによるテスト実行中にエラーが発生しました: {e}")
        import traceback
        logging.error(traceback.format_exc())
        # エラー発生時はキーに対応する空の値を返す
        return {key: "" for key in sample_format_keys.keys()}


def process_samples(procedure: str, batch_path: Path, sample_format_keys: Dict[str, Any], test_mode: bool = False) -> list[dict]:
    """バッチ内の各サンプルの処理を実行する"""
    results = []
    sample_dirs = sorted([p for p in batch_path.iterdir() if p.is_dir()])
    
    # テストモードの場合は最初の3サンプルのみ処理
    if test_mode:
        sample_dirs = sample_dirs[:3]
        st.info("テストモード: 最初の3サンプルのみ処理します。")

    progress_bar = st.progress(0)
    total_samples = len(sample_dirs)

    for i, sample_dir in enumerate(sample_dirs):
        st.info(f"処理中: {sample_dir.name} ({i+1}/{total_samples})")
        
        sample_contents = [
            read_sample(p) for p in sorted(sample_dir.glob("*"))
        ]

        # LLMでテストを実行
        result_json = execute_audit_procedure(st.session_state.llm, procedure, sample_contents, sample_format_keys)

        # 結果をリストに追加
        results.append({
            "sample": sample_dir.name,
            "result": result_json
        })
        
        progress_bar.progress((i + 1) / total_samples)

    st.success("すべてのサンプルの処理が完了しました。")
    return results


def main():
    """メイン関数"""
    st.set_page_config(page_title="サンプルテスト自動化ツール", layout="wide")
    st.title("サンプルテスト自動化ツール")

    # LLMの初期化
    if "llm" not in st.session_state:
        st.session_state.llm = create_chat_model()

    # --- サイドバー ---
    with st.sidebar:
        st.header("設定")
        procedure = st.text_area("監査手続き", height=150, value="請求書と出荷記録を照合し、金額、日付、品目が一致することを確認してください。")
        
        batches = list_batches()
        selected_batch = st.selectbox("処理するバッチ", batches)

        templates = list_templates()
        selected_template = st.selectbox("調書テンプレート", templates)
        
        update_workpaper = st.checkbox("調書を更新", value=True, help="チェックすると、レイアウト分析とExcelへの転記を行います。")
        test_mode = st.checkbox("テストモード", value=True, help="チェックを入れると、最初の3サンプルのみ処理します。")
        
        execute_button = st.button("実行", type="primary")

    # --- メイン画面 ---
    if execute_button:
        # 入力チェック
        if not procedure or not selected_batch:
            st.error("「監査手続き」と「処理するバッチ」を選択してください。")
            st.stop()
        if update_workpaper and not selected_template:
            st.error("「調書を更新」にチェックが入っている場合は、「調書テンプレート」も選択してください。")
            st.stop()

        batch_path = BATCH_DIR / selected_batch
        template_path = FORMAT_DIR / selected_template if selected_template else None

        analysis_result = None
        # 1. レイアウト分析 (チェックボックスがONの場合のみ)
        if update_workpaper:
            with st.spinner("調書レイアウトを分析中です..."):
                analyzer = ExcelFormatAnalyzer(
                    model_name=MODEL_NAME,
                    api_key=OPENAI_API_KEY,
                    azure_api_key=AZURE_OPENAI_API_KEY,
                    azure_endpoint=AZURE_OPENAI_ENDPOINT,
                    azure_deployment=AZURE_OPENAI_DEPLOYMENT,
                    azure_api_version=AZURE_OPENAI_API_VERSION,
                )
                analysis_result = analyzer.analyze_format(str(template_path), output_dir=str(TEMP_DIR))

                if analysis_result and analysis_result["status"] == "success":
                    st.success("レイアウト分析が完了しました。")
                    with st.expander("分析結果の詳細"):
                        st.json(analysis_result)
                else:
                    st.error(f"レイアウト分析に失敗しました: {analysis_result.get('error_message', '不明なエラー')}")
                    st.stop()

        # 2. 監査テスト実行
        with st.spinner("監査テストを実行中です..."):
            sample_format_keys = {}
            if analysis_result:
                sample_format_keys = analysis_result.get("sample_fields", {}).get("columns", {})
            elif not update_workpaper:
                # 調書更新しない場合でも、基本的な結果は取得する
                sample_format_keys = {
                    "result": {"description": "監査手続きの結果。'OK' または 'NG' のいずれか。"},
                    "reason": {"description": "結果の根拠となった理由を200字以内で記述してください。"},
                    "data_detail": {"description": "各サンプルデータの内容を要約してください。"}
                }
            
            test_results = process_samples(procedure, batch_path, sample_format_keys, test_mode)
            st.subheader("監査テスト結果")
            st.write(test_results)
            # test_results を dataframe に変換 ※result の中身は json なので、json を展開して dataframe に変換
            test_results_df = pd.DataFrame([{**r, **r["result"]} for r in test_results])
            st.dataframe(test_results_df)

        # 3. 調書への転記 (チェックボックスがONの場合のみ)
        if update_workpaper and analysis_result and test_results:
            with st.spinner("結果を調書に転記中です..."):
                output_filename = f"updated_{Path(selected_template).name}"
                output_path = TEMP_DIR / output_filename
                
                success = fill_excel_with_results(
                    template_path=str(template_path),
                    output_path=str(output_path),
                    analysis_result=analysis_result,
                    test_results=test_results,
                    procedure=procedure,
                    model=st.session_state.llm
                )

                if success:
                    st.success("すべての処理が完了しました！")
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="完成した調書をダウンロード",
                            data=f,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("調書への転記中にエラーが発生しました。")
    else:
        st.info("サイドバーで設定を行い、「実行」ボタンを押してください。")


if __name__ == "__main__":
    main() 