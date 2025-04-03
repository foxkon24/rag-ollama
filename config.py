# config.py - 環境変数の読み込みと設定（OneDrive検索設定対応版）
import os
import sys
import logging
from dotenv import load_dotenv

logger = logging.getLogger(__name__)

def load_config(script_dir):
    """
    環境変数を読み込み、設定する

    Args:
        script_dir: スクリプトのディレクトリパス

    Returns:
        dict: 設定値を含む辞書
    """
    # 環境変数の読み込み - 明示的にパスを指定
    try:
        env_path = os.path.join(script_dir, '.env')

        if os.path.exists(env_path):
            logger.info(f".env ファイルを読み込みます: {env_path}")
            load_dotenv(env_path)
            logger.info(".env ファイルの読み込みが完了しました")
        else:
            logger.error(f".env ファイルが見つかりません: {env_path}")

    except ImportError:
        logger.error("python-dotenv がインストールされていません。pip install python-dotenv でインストールしてください。")
        sys.exit(1)

    # 環境変数の読み込み
    config = {
        "OLLAMA_URL": os.getenv("OLLAMA_URL"),
        "OLLAMA_MODEL": os.getenv("OLLAMA_MODEL"),
        "OLLAMA_TIMEOUT": int(os.getenv("OLLAMA_TIMEOUT", "60")),  # デフォルト60秒
        "TEAMS_OUTGOING_TOKEN": os.getenv("TEAMS_OUTGOING_TOKEN"),
        "TEAMS_WORKFLOW_URL": os.getenv("TEAMS_WORKFLOW_URL"),  # Logic Apps URLを使用
        "HOST": os.getenv("HOST", "0.0.0.0"),
        "PORT": int(os.getenv("PORT", 5010)),
        "DEBUG": os.getenv("FLASK_DEBUG", "0") == "1",
        # OneDrive検索設定
        "ONEDRIVE_SEARCH_ENABLED": os.getenv("ONEDRIVE_SEARCH_ENABLED", "0") == "1",
        "ONEDRIVE_SEARCH_DIR": os.getenv("ONEDRIVE_SEARCH_DIR", ""),
        "ONEDRIVE_MAX_FILES": int(os.getenv("ONEDRIVE_MAX_FILES", "5")),
        "ONEDRIVE_FILE_TYPES": parse_file_types(os.getenv("ONEDRIVE_FILE_TYPES", "")),
        "SKIP_VERIFICATION": os.getenv("SKIP_VERIFICATION", "0") == "1"
    }

    # 読み込まれた環境変数を確認（センシティブな情報は一部のみ表示）
    logger.info(f"OLLAMA_URL: {config['OLLAMA_URL']}")
    logger.info(f"OLLAMA_MODEL: {config['OLLAMA_MODEL']}")
    logger.info(f"OLLAMA_TIMEOUT: {config['OLLAMA_TIMEOUT']}秒")
    logger.info(f"TEAMS_OUTGOING_TOKEN: {'設定済み' if config['TEAMS_OUTGOING_TOKEN'] else 'なし'}")

    # Webhook URLの表示
    webhook_url = config['TEAMS_WORKFLOW_URL']
    if webhook_url:
        if "logic.azure.com" in webhook_url:
            logger.info(f"Teams Workflow URL: {webhook_url[:30]}... (Azure Logic Apps)")
        else:
            logger.info(f"Teams Workflow URL: {webhook_url[:30]}...")
    else:
        logger.error("Teams Workflow URLが設定されていません")

    # OneDrive検索設定の表示
    logger.info(f"OneDrive検索: {'有効' if config['ONEDRIVE_SEARCH_ENABLED'] else '無効'}")
    if config['ONEDRIVE_SEARCH_ENABLED']:
        logger.info(f"OneDrive検索ディレクトリ: {config['ONEDRIVE_SEARCH_DIR'] if config['ONEDRIVE_SEARCH_DIR'] else 'OneDriveルート'}")
        logger.info(f"OneDrive最大ファイル数: {config['ONEDRIVE_MAX_FILES']}")
        logger.info(f"OneDrive検索対象ファイル: {', '.join(config['ONEDRIVE_FILE_TYPES']) if config['ONEDRIVE_FILE_TYPES'] else '全ファイル'}")

    # 環境変数のバックアップ（.envが読み込めなかった場合）
    if not config['OLLAMA_URL']:
        config['OLLAMA_URL'] = "http://localhost:11434/api/generate"
        logger.warning(f"OLLAMA_URL が設定されていないため、デフォルト値を使用します: {config['OLLAMA_URL']}")

    if not config['OLLAMA_MODEL']:
        config['OLLAMA_MODEL'] = "llama3"
        logger.warning(f"OLLAMA_MODEL が設定されていないため、デフォルト値を使用します: {config['OLLAMA_MODEL']}")

    # 環境変数が読み込まれない場合は.envファイルの内容をハードコーディング
    if not config['TEAMS_WORKFLOW_URL']:
        logger.warning("TEAMS_WORKFLOW_URL が設定されていないため、.env ファイルから直接読み込みを試みます")

        try:
            with open(env_path, 'r') as f:
                for line in f:
                    if line.strip() and not line.startswith('#'):
                        key, value = line.strip().split('=', 1)

                        if key == 'TEAMS_WORKFLOW_URL':
                            config['TEAMS_WORKFLOW_URL'] = value
                            logger.info(f"TEAMS_WORKFLOW_URL を設定しました: {value[:30]}...")
                        elif key == 'TEAMS_OUTGOING_TOKEN' and not config['TEAMS_OUTGOING_TOKEN']:
                            config['TEAMS_OUTGOING_TOKEN'] = value
                            logger.info("TEAMS_OUTGOING_TOKEN を設定しました")

        except Exception as e:
            logger.error(f".env ファイルの直接読み込み中にエラーが発生しました: {str(e)}")

    # それでも設定されていない場合はハードコーディング（本番環境では使用しないでください）
    if not config['TEAMS_WORKFLOW_URL']:
        config['TEAMS_WORKFLOW_URL'] = "https://prod-30.japaneast.logic.azure.com:443/workflows/c5bf0a8922ca44afb03ea9bcc53f54b6/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Pdgejv1s91AIb2_XBetvzA8NcJ9qQSFOTeqUUKIvcAA"
        logger.warning("TEAMS_WORKFLOW_URL がハードコーディングされた値を使用しています")

    if not config['TEAMS_OUTGOING_TOKEN']:
        config['TEAMS_OUTGOING_TOKEN'] = "5yt2f1X18I//jX0BoVREgMqZl8QLl+lymis6gkvDObY="
        logger.warning("TEAMS_OUTGOING_TOKEN がハードコーディングされた値を使用しています")

    return config

def parse_file_types(file_types_str):
    """
    カンマ区切りのファイル拡張子文字列をリストに変換する

    Args:
        file_types_str: カンマ区切りのファイル拡張子文字列

    Returns:
        list: ファイル拡張子のリスト
    """
    if not file_types_str:
        return []

    # カンマ区切りで分割し、各項目の空白を除去
    types = [t.strip() for t in file_types_str.split(',')]

    # 空の要素を除去し、各拡張子の先頭にドットがない場合は追加
    valid_types = []
    for t in types:
        if t:
            if not t.startswith('.'):
                t = '.' + t
            valid_types.append(t)

    return valid_types
