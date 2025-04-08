# async_processor.py - OneDrive検索機能を組み込んだ非同期処理（改善版）
import logging
import traceback
import subprocess
import re
import os
from ollama_client import generate_ollama_response

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # 詳細なログを有効化

def process_query_async(query_text, original_data, ollama_url, ollama_model, ollama_timeout, teams_webhook, onedrive_search=None):
    """
    クエリを非同期で処理し、結果をTeamsに通知する（OneDrive検索機能付き・改善版）

    Args:
        query_text: 処理するクエリテキスト
        original_data: 元のTeamsリクエストデータ
        ollama_url: OllamaのURL
        ollama_model: 使用するOllamaモデル
        ollama_timeout: リクエストのタイムアウト時間（秒）
        teams_webhook: Teams Webhookインスタンス
        onedrive_search: OneDriveSearch インスタンス（Noneの場合は検索しない）
    """
    try:
        # クリーンなクエリを抽出
        clean_query = query_text.replace('ollama質問', '').strip()

        # OneDrive検索を実行するかどうかを判断
        use_onedrive = onedrive_search is not None and 'onedrive' not in clean_query.lower()

        # 検索ディレクトリパスを取得（検索するかどうか関係なく）
        search_path = None
        if onedrive_search is not None:
            search_path = onedrive_search.base_directory
        
        # 日付を含むかどうかを確認（日報検索に重要）
        date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', clean_query)
        has_date = bool(date_match)
        
        # OneDrive検索のログ
        if use_onedrive:
            if has_date:
                # 日付情報を詳細に出力
                year = date_match.group(1)
                month = date_match.group(2).zfill(2)
                day = date_match.group(3).zfill(2)
                date_str = f"{year}年{month}月{day}日"
                logger.info(f"日付指定のある検索を実行します: '{date_str}'")
                
                # 検索前に年度フォルダの存在チェック
                if onedrive_search:
                    try:
                        fiscal_year = onedrive_search.get_fiscal_year_folder(year, month)
                        # PowerShellでディレクトリ存在チェック
                        ps_command = f"""
                        Get-ChildItem -Path "{search_path}" -Recurse -Directory | 
                        Where-Object {{ $_.Name -like "*{fiscal_year}*" -or $_.FullName -like "*{fiscal_year}*" }} | 
                        Select-Object -First 1 | ForEach-Object {{ $_.FullName }}
                        """
                        result = subprocess.check_output(["powershell", "-Command", ps_command], text=True).strip()
                        if result:
                            logger.info(f"年度フォルダが見つかりました: {result}")
                            # フォルダ内のファイル例を確認
                            ps_command2 = f"""
                            Get-ChildItem -Path "{result}" -File | 
                            Select-Object -First 3 | ForEach-Object {{ $_.Name }}
                            """
                            file_examples = subprocess.check_output(["powershell", "-Command", ps_command2], text=True).strip()
                            if file_examples:
                                logger.info(f"フォルダ内のファイル例: \n{file_examples}")
                        else:
                            logger.warning(f"年度フォルダ '{fiscal_year}' が見つかりませんでした")
                    except Exception as e:
                        logger.error(f"年度フォルダチェック中にエラー: {str(e)}")
            else:
                logger.info(f"通常のOneDrive検索を実行します: '{clean_query}'")
        else:
            if onedrive_search is None:
                logger.info("OneDrive検索が無効化されています")
            else:
                logger.info("OneDriveに関する質問のため、検索をスキップします")

        # Ollamaで回答を生成（OneDrive検索結果を含む）
        response = generate_ollama_response(
            query_text, 
            ollama_url, 
            ollama_model, 
            ollama_timeout,
            onedrive_search if use_onedrive else None
        )
        logger.info(f"非同期処理による応答生成完了: {response[:100]}...")

        if teams_webhook:
            # TEAMS_WORKFLOW_URLを使用して直接Teamsに送信
            logger.info("Teamsに直接応答を送信します")
            result = teams_webhook.send_ollama_response(clean_query, response, None, search_path)
            logger.info(f"Teams送信結果: {result}")

            # 送信に失敗した場合の処理
            if result.get("status") == "error":
                logger.error(f"Teams送信エラー: {result.get('message', 'unknown error')}")
                # エラー発生時は何もしない（すでにログに記録済み）

        else:
            logger.error("Teams Webhookが設定されていないため、通知できません")

    except Exception as e:
        logger.error(f"非同期処理中にエラーが発生しました: {str(e)}")
        logger.error(traceback.format_exc())
