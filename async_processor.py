# async_processor.py - OneDrive検索機能を組み込んだ非同期処理
import logging
import traceback
from ollama_client import generate_ollama_response

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # 詳細なログを有効化

def process_query_async(query_text, original_data, ollama_url, ollama_model, ollama_timeout, teams_webhook, onedrive_search=None):
    """
    クエリを非同期で処理し、結果をTeamsに通知する（OneDrive検索機能付き）

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

        # OneDrive検索のログ
        if use_onedrive:
            logger.info(f"OneDrive検索を実行します: '{clean_query}'")
            # 検索結果はOllama処理内で取得される
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
            result = teams_webhook.send_ollama_response(clean_query, response)
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
