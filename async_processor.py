# async_processor.py - OneDrive検索機能を組み込んだ非同期処理（pypdf対応版）
import logging
import traceback
import time
import re
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
        
        # ログにリクエスト情報を記録
        logger.info(f"非同期処理を開始: query='{clean_query}', model={ollama_model}")
        
        # 複数のフォーマットに対応する日付パターン認識
        date_info = extract_date_from_query(clean_query)
        
        if date_info:
            year, month, day, format_type = date_info
            logger.info(f"日付指定のある検索を実行します: '{year}年{month}月{day}日' (元の形式: {format_type})")
        
        # OneDrive検索を実行するかどうかを判断
        use_onedrive = onedrive_search is not None and 'onedrive' not in clean_query.lower()

        # 検索ディレクトリパスを取得（検索するかどうか関係なく）
        search_path = None
        if onedrive_search is not None:
            search_path = onedrive_search.base_directory
        
        # OneDrive検索のログ
        if use_onedrive:
            logger.info(f"OneDrive検索を実行します: '{clean_query}'")
            logger.info(f"検索ディレクトリ: {search_path}")
            # 検索結果はOllama処理内で取得される
        else:
            if onedrive_search is None:
                logger.info("OneDrive検索が無効化されています")
            else:
                logger.info("OneDriveに関する質問のため、検索をスキップします")

        # Ollamaで回答を生成（OneDrive検索結果を含む）
        logger.info(f"Ollama APIリクエスト開始: {ollama_url}")
        start_time = time.time()
        
        response = generate_ollama_response(
            query_text, 
            ollama_url, 
            ollama_model, 
            ollama_timeout,
            onedrive_search if use_onedrive else None
        )
        
        end_time = time.time()
        logger.info(f"非同期処理による応答生成完了: 処理時間={end_time - start_time:.2f}秒, 応答長={len(response)}文字")
        logger.info(f"応答内容: {response[:150]}...")  # 応答の先頭部分をログに記録

        # pypdfがなくてPDF抽出に失敗した場合は、インストールを試みる
        if "PDFファイルからテキスト抽出がサポートされていない" in response and onedrive_search:
            logger.info("PDF抽出が失敗したため、pypdfのインストールを試みます")
            try:
                import subprocess
                subprocess.call(["pip", "install", "pypdf"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                logger.info("pypdfのインストールを試みました。次回のリクエストでPDF抽出機能が使用できるかもしれません")
            except:
                logger.warning("pypdfのインストールに失敗しました")

        if teams_webhook:
            # Teams送信前にログを記録
            logger.info(f"Teamsメッセージ送信を開始します: webhook_url={teams_webhook.webhook_url[:30]}...")
            
            # TEAMS_WORKFLOW_URLを使用して直接Teamsに送信
            start_send_time = time.time()
            
            # 回答が空の場合のバックアップメッセージ
            if not response or len(response.strip()) < 10:
                response = "申し訳ありません。有効な回答を生成できませんでした。しばらく経ってから再度お試しください。"
                logger.warning(f"生成された回答が空または短すぎるため、デフォルトメッセージを使用します")
            
            result = teams_webhook.send_ollama_response(clean_query, response, None, search_path)
            
            end_send_time = time.time()
            logger.info(f"Teams送信結果: {result}, 送信時間={end_send_time - start_send_time:.2f}秒")

            # 送信が成功したかどうかを確認
            if result.get("status") == "success":
                logger.info(f"✅ Teams送信成功: 形式={result.get('format')}, コード={result.get('code')}")
            else:
                logger.error(f"❌ Teams送信エラー: {result.get('message', 'unknown error')}")
                # エラーの詳細を記録
                error_msg = f"Teams送信エラー詳細: {result}"
                logger.error(error_msg)
                
                # 再試行 - シンプルな形式でもう一度
                try:
                    logger.info("シンプルな形式で再試行します")
                    # Teams Webhookの直接HTTPリクエスト
                    import requests
                    simple_payload = {
                        "text": f"### Ollama回答 (再送)\n\n**質問**: {clean_query}\n\n{response}\n\n*回答生成時刻: {time.strftime('%Y年%m月%d日 %H:%M:%S')}*"
                    }
                    headers = {'Content-Type': 'application/json; charset=utf-8'}
                    r = requests.post(teams_webhook.webhook_url, json=simple_payload, headers=headers, timeout=30)
                    logger.info(f"再試行結果: ステータスコード={r.status_code}")
                except Exception as retry_err:
                    logger.error(f"再試行にも失敗しました: {str(retry_err)}")

        else:
            logger.error("Teams Webhookが設定されていないため、通知できません")

    except Exception as e:
        logger.error(f"非同期処理中にエラーが発生しました: {str(e)}")
        logger.error(traceback.format_exc())
        
        # エラー情報をTeamsに送信（可能であれば）
        if teams_webhook:
            try:
                error_message = f"エラーが発生しました: {str(e)}\n\n詳細はサーバーログを確認してください。"
                teams_webhook.send_ollama_response(clean_query, error_message, None, search_path)
            except:
                pass  # エラー通知に失敗した場合は、これ以上何もしない

def extract_date_from_query(query):
    """
    クエリから日付情報を抽出する（複数のフォーマットに対応）
    
    Args:
        query: 検索クエリ文字列
    
    Returns:
        tuple: (年, 月, 日, フォーマット) または None（日付が見つからない場合）
    """
    # 1. YYYY年MM月DD日 形式を確認
    japanese_date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', query)
    if japanese_date_match:
        return (
            japanese_date_match.group(1),
            japanese_date_match.group(2).zfill(2),
            japanese_date_match.group(3).zfill(2),
            "和暦形式"
        )
    
    # 2. YYYY/MM/DD または YYYY-MM-DD 形式を確認
    slash_date_match = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', query)
    if slash_date_match:
        return (
            slash_date_match.group(1),
            slash_date_match.group(2).zfill(2),
            slash_date_match.group(3).zfill(2),
            "スラッシュ区切り形式"
        )
    
    # 3. YYYYMMDD 形式（8桁の数字）を確認
    numeric_date_match = re.search(r'\b(\d{4})(\d{2})(\d{2})\b', query)
    if numeric_date_match:
        return (
            numeric_date_match.group(1),
            numeric_date_match.group(2),
            numeric_date_match.group(3),
            "数値形式"
        )
    
    # 4. 独立した8桁の数字がYYYYMMDDとして有効かどうか確認
    for word in query.split():
        if word.isdigit() and len(word) == 8:
            year = word[:4]
            month = word[4:6]
            day = word[6:8]
            
            # 日付として有効かどうか確認（簡易チェック）
            try:
                month_int = int(month)
                day_int = int(day)
                if 1 <= month_int <= 12 and 1 <= day_int <= 31:
                    return (year, month, day, "数値形式")
            except ValueError:
                continue
    
    # 日付が見つからない
    return None
