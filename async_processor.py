# async_processor.py - OneDrive検索結果の表示改善版
import logging
import traceback
from ollama_client import generate_ollama_response
import json

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # 詳細なログを有効化

def process_query_async(query_text, original_data, ollama_url, ollama_model, ollama_timeout, teams_webhook, onedrive_search=None):
    """
    クエリを非同期で処理し、結果をTeamsに通知する（検索結果表示改善版）

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
            
            # 改善: 検索結果を段落や見出しで適切に整形して送信
            formatted_response = format_teams_response(clean_query, response, search_path)
            result = teams_webhook.send_ollama_response(clean_query, formatted_response, None, search_path)
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

def format_teams_response(query, response, search_path=None):
    """
    Teams通知用に応答を整形する
    
    Args:
        query: ユーザーのクエリ
        response: 生成された応答
        search_path: 検索パス
        
    Returns:
        整形された応答テキスト
    """
    # 以下のセクションが含まれているかを確認
    has_search_results = "【検索結果】" in response or "件の関連ファイルが見つかりました" in response
    has_file_info = "【ファイル" in response or "ファイル情報" in response
    has_file_content = "【ファイル内容】" in response or "【Word文書内容】" in response or "【Excel内容】" in response
    
    # 検索結果がある場合は、見やすく整形
    if has_search_results or has_file_info:
        # セクションごとに分割して処理
        sections = []
        
        # 検索結果セクションを抽出
        if "【検索結果】" in response:
            search_results_idx = response.find("【検索結果】")
            next_section_idx = find_next_section(response, search_results_idx)
            if next_section_idx > 0:
                search_results = response[search_results_idx:next_section_idx]
            else:
                search_results = response[search_results_idx:]
            sections.append(search_results)
        
        # ファイル情報セクションを抽出
        file_sections = []
        current_pos = 0
        while True:
            file_idx = response.find("【ファイル ", current_pos)
            if file_idx < 0:
                break
                
            next_section_idx = find_next_section(response, file_idx)
            if next_section_idx > 0:
                file_section = response[file_idx:next_section_idx]
            else:
                file_section = response[file_idx:]
                
            file_sections.append(file_section)
            current_pos = next_section_idx if next_section_idx > 0 else len(response)
        
        # 他のファイル情報セクションを探す
        for section_start in ["【Word文書情報】", "【Excel情報】", "【PowerPoint情報】", "【PDF文書情報】", "【テキストファイル情報】"]:
            section_idx = response.find(section_start)
            if section_idx >= 0:
                next_section_idx = find_next_section(response, section_idx)
                if next_section_idx > 0:
                    section = response[section_idx:next_section_idx]
                else:
                    section = response[section_idx:]
                file_sections.append(section)
        
        # ファイルコンテンツセクションを抽出
        content_sections = []
        for content_start in ["【ファイル内容】", "【Word文書内容】", "【Excel内容】", "【PowerPoint内容】", "【PDF内容】"]:
            content_idx = response.find(content_start)
            if content_idx >= 0:
                next_section_idx = find_next_section(response, content_idx)
                if next_section_idx > 0:
                    content = response[content_idx:next_section_idx]
                else:
                    content = response[content_idx:]
                content_sections.append(content)
        
        # 最終的な応答を組み立て
        formatted_response = ""
        
        # まず検索結果セクションを追加
        for section in sections:
            formatted_response += section + "\n\n"
        
        # ファイル情報セクションを追加（最大2個まで）
        for i, section in enumerate(file_sections[:2]):
            formatted_response += section + "\n\n"
        
        if len(file_sections) > 2:
            formatted_response += f"... 他 {len(file_sections) - 2} 件のファイル情報は省略されました\n\n"
        
        # ファイルコンテンツセクションを追加（最大2個まで）
        for i, section in enumerate(content_sections[:2]):
            formatted_response += section + "\n\n"
        
        if len(content_sections) > 2:
            formatted_response += f"... 他 {len(content_sections) - 2} 件のファイル内容は省略されました\n\n"
        
        # 整形された応答を返す
        return formatted_response.strip()
    
    # 検索結果がなければ、元の応答をそのまま返す
    return response

def find_next_section(text, start_pos):
    """
    次のセクション開始位置を見つける
    
    Args:
        text: 検索するテキスト
        start_pos: 検索開始位置
        
    Returns:
        次のセクション開始位置、見つからなければ -1
    """
    section_markers = [
        "【検索結果】", "【ファイル ", "【Word文書情報】", "【Excel情報】", 
        "【PowerPoint情報】", "【PDF文書情報】", "【テキストファイル情報】",
        "【ファイル内容】", "【Word文書内容】", "【Excel内容】", "【PowerPoint内容】", "【PDF内容】"
    ]
    
    next_pos = -1
    for marker in section_markers:
        pos = text.find(marker, start_pos + 1)
        if pos >= 0 and (next_pos < 0 or pos < next_pos):
            next_pos = pos
            
    return next_pos
