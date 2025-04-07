# routes.py - Flaskルート定義（OneDrive検索機能追加版、署名検証改善版）
import re
import logging
import threading
import traceback
from datetime import datetime
from flask import request, jsonify, render_template
from teams_auth import verify_teams_token, debug_teams_signature
from async_processor import process_query_async
import requests

logger = logging.getLogger(__name__)

def register_routes(app, config, teams_webhook, onedrive_search=None):
    """
    Flaskアプリにルートを登録する

    Args:
        app: Flaskアプリインスタンス
        config: 設定辞書
        teams_webhook: Teams Webhookインスタンス
        onedrive_search: OneDriveSearch インスタンス（Noneの場合は検索しない）
    """

    @app.route('/webhook', methods=['POST'])
    def teams_webhook_handler():
        """
        Teamsからのwebhookリクエストを処理する（OneDrive検索機能付き）
        """
        try:
            logger.info("Webhookリクエストを受信しました")
            logger.info(f"リクエストヘッダー: {request.headers}")
            logger.debug(f"リクエストデータ: {request.get_data()}")  # デバッグレベルに変更

            # 署名検証をスキップするかどうか（.env設定より）
            skip_verification = config['SKIP_VERIFICATION']

            # Teamsからの署名を検証
            signature = request.headers.get('Authorization')
            logger.info(f"受信した認証ヘッダー: {signature}")
            
            if not skip_verification:
                # リクエストデータを取得
                request_data = request.get_data()
                logger.debug(f"検証するデータ: {request_data[:100]}...")
                
                # デバッグ情報を記録
                debug_info = debug_teams_signature(request_data, config['TEAMS_OUTGOING_TOKEN'])
                logger.debug(f"デバッグ署名情報: {debug_info}")
                
                # 署名を検証（改善された検証ロジック）
                if not verify_teams_token(request_data, signature, config['TEAMS_OUTGOING_TOKEN']):
                    # 本番環境では認証失敗時に403エラーを返す
                    if not config['DEBUG']:
                        logger.error("署名の検証に失敗しました。アクセスを拒否します。")
                        return jsonify({"error": "Unauthorized"}), 403
                    else:
                        # 開発環境では警告を出して続行
                        logger.warning("署名の検証に失敗しました。デバッグモードで続行します。")
                        logger.warning("本番環境では .env の SKIP_VERIFICATION=1 を設定してください。")

            data = request.json
            logger.info(f"処理するデータ: {data}")

            # Teamsからのメッセージを取得
            if 'text' in data:
                # HTMLタグを除去（Teams形式対応）
                query_text = re.sub(r'<.*?>|\r\n', ' ', data['text'])
                logger.info(f"整形後のクエリ: {query_text}")

                # OneDrive検索が有効かどうかを確認
                if onedrive_search:
                    logger.info("OneDrive検索機能が有効です")
                else:
                    logger.info("OneDrive検索機能は無効です")

                # 非同期で処理を行うためにスレッドを起動
                thread = threading.Thread(
                    target=process_query_async,
                    args=(query_text, data, config['OLLAMA_URL'], config['OLLAMA_MODEL'], 
                          config['OLLAMA_TIMEOUT'], teams_webhook, onedrive_search)
                )
                thread.daemon = True  # メインスレッド終了時に一緒に終了
                thread.start()

                # すぐに応答を返す (Teams Outgoing Webhookタイムアウト回避)
                return jsonify({
                    "type": "message",
                    "text": "リクエストを受け付けました。回答を生成中です..." + 
                           (" OneDriveからの関連情報も検索します。" if onedrive_search else "")
                })

            else:
                logger.error("テキストフィールドが見つかりません")
                return jsonify({"error": "テキストフィールドが見つかりません"}), 400

        except Exception as e:
            logger.error(f"Webhookの処理中にエラーが発生しました: {str(e)}")
            # スタックトレースをログに記録
            logger.error(traceback.format_exc())
            return jsonify({"error": str(e)}), 500

    @app.route('/health', methods=['GET'])
    def health_check():
        """
        システムの健全性を確認するエンドポイント
        """
        try:
            # Ollamaサーバーの状態を確認
            try:
                ollama_status = requests.get(config['OLLAMA_URL'].replace("/api/generate", "/api/version"), timeout=5).status_code == 200
            except:
                ollama_status = False

            # Teams Webhookの状態を確認
            teams_webhook_status = config['TEAMS_WORKFLOW_URL'] is not None and teams_webhook is not None
            
            # OneDrive検索の状態を確認
            onedrive_status = onedrive_search is not None

            return jsonify({
                "status": "ok" if ollama_status else "degraded",
                "timestamp": datetime.now().isoformat(),
                "system": "Ollama OneDrive Knowledge System",
                "ollama_status": "connected" if ollama_status else "disconnected",
                "teams_webhook_status": "configured" if teams_webhook_status else "not configured",
                "onedrive_status": "enabled" if onedrive_status else "disabled",
                "model": config['OLLAMA_MODEL']
            })

        except Exception as e:
            logger.error(f"ヘルスチェック中にエラーが発生しました: {str(e)}")
            return jsonify({
                "status": "error",
                "timestamp": datetime.now().isoformat(),
                "error": str(e)
            }), 500

    @app.route('/', methods=['GET'])
    def index():
        try:
            return render_template('index.html')
        except:
            is_onedrive = "有効" if onedrive_search else "無効"
            return f"Ollama Webhook System is running! OneDrive検索機能: {is_onedrive}"
