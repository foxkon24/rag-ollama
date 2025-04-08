# teams_webhook.py - Power Automateワークフロー対応版（修正版）
import requests
import logging
import json
from datetime import datetime
import traceback
import re
import os

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

class TeamsWebhook:
    def __init__(self, webhook_url):
        """
        Teams Webhookクラスの初期化

        Args:
            webhook_url: Teams Workflow URL (Logic Apps URL)
        """
        self.webhook_url = webhook_url
        logger.info(f"Teams Workflowを初期化: {webhook_url[:30]}...")

    def send_ollama_response(self, query, response, conversation_data=None, search_path=None):
        """
        Ollamaの応答をTeams Workflowに送信する（修正版）

        Args:
            query: ユーザーの質問
            response: Ollamaからの応答
            conversation_data: 会話データ（オプション）
            search_path: 検索に使用したディレクトリパス（オプション）

        Returns:
            dict: 送信結果
        """
        try:
            # 検索パスの短縮表示
            short_path = self._get_shortened_path(search_path) if search_path else "設定されたディレクトリ"
            
            # 現在の日時
            now = datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')
            
            # Power Automateワークフロー用のペイロード（主にこれを使用）
            # body.attachments 構造に修正
            power_automate_payload = {
                "body": {
                    "attachments": [
                        {
                            "contentType": "application/vnd.microsoft.card.adaptive",
                            "content": {
                                "type": "AdaptiveCard",
                                "body": [
                                    {
                                        "type": "TextBlock",
                                        "size": "Medium",
                                        "weight": "Bolder",
                                        "text": "Ollama回答",
                                        "wrap": True
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"質問: {query}",
                                        "wrap": True,
                                        "weight": "Bolder",
                                        "color": "Accent"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"検索対象: {short_path}",
                                        "wrap": True,
                                        "isSubtle": True,
                                        "size": "Small"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": response,
                                        "wrap": True,
                                        "spacing": "Medium"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"回答生成時刻: {now}",
                                        "wrap": True,
                                        "size": "Small",
                                        "isSubtle": True
                                    }
                                ],
                                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                                "version": "1.0"
                            }
                        }
                    ]
                }
            }

            # バックアップ用のシンプルなペイロード
            simple_payload = {
                "text": f"### Ollama回答\n\n**質問**: {query}\n\n**検索対象**: {short_path}\n\n{response}\n\n*回答生成時刻: {now}*"
            }

            # リクエストヘッダー
            headers = {
                'Content-Type': 'application/json'
            }

            # Power Automateワークフロー用のペイロードで試行
            logger.debug(f"Logic Apps送信ペイロード(Power Automate形式): {json.dumps(power_automate_payload)[:300]}...")

            try:
                r = requests.post(
                    self.webhook_url, 
                    json=power_automate_payload, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答: {r.status_code}, {r.text[:100] if r.text else '空のレスポンス'}")

                if r.status_code >= 200 and r.status_code < 300:
                    logger.info(f"Power Automate形式でのLogic Apps通知送信成功: {r.status_code}")
                    return {"status": "success", "code": r.status_code, "format": "Power Automate形式"}
                else:
                    logger.warning(f"Power Automate形式での送信失敗: {r.status_code}。シンプル形式で再試行します。")
                    # エラーレスポンスの詳細をログに記録
                    logger.warning(f"エラーレスポンス: {r.text}")

            except Exception as e:
                logger.warning(f"Power Automate形式送信エラー: {str(e)}。シンプル形式で再試行します。")

            # シンプル形式で試行（最後の手段）
            logger.debug(f"Logic Apps送信ペイロード(シンプル): {json.dumps(simple_payload)[:300]}...")

            try:
                r3 = requests.post(
                    self.webhook_url, 
                    json=simple_payload, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答(シンプル): {r3.status_code}, {r3.text[:100] if r3.text else '空のレスポンス'}")

                if r3.status_code >= 200 and r3.status_code < 300:
                    logger.info(f"シンプル形式でのLogic Apps通知送信成功: {r3.status_code}")
                    return {"status": "success", "code": r3.status_code, "format": "シンプル"}
                else:
                    logger.error(f"Logic Apps通知の送信に全て失敗しました: 最終ステータスコード={r3.status_code}")
                    return {"status": "error", "code": r3.status_code, "message": r3.text}

            except Exception as e3:
                logger.error(f"シンプル形式送信エラー: {str(e3)}")
                return {"status": "error", "message": str(e3)}

        except Exception as e:
            logger.error(f"Logic Apps通知の送信中にエラーが発生しました: {str(e)}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def _get_shortened_path(self, path):
        """
        長いパスを短縮して表示

        Args:
            path: 元のパス文字列

        Returns:
            短縮されたパス
        """
        if not path:
            return "デフォルト検索ディレクトリ"
            
        # ユーザー名を取得
        username = os.getenv("USERNAME", "owner")
        
        # パスが長い場合は短縮表示
        if len(path) > 50:
            # OneDriveパスの特定のパターンを検出
            if "OneDrive" in path:
                # 会社名を含むOneDriveパスのパターン
                company_match = re.search(r'OneDrive - ([^\\]+)', path)
                if company_match:
                    company = company_match.group(1)
                    # 短縮した会社名
                    short_company = company[:10] + "..." if len(company) > 10 else company
                    # パスの後半部分を取得
                    path_parts = path.split("\\")
                    if len(path_parts) > 3:
                        last_parts = path_parts[-3:]
                        return f"OneDrive - {short_company}\\...\\{last_parts[-3]}\\{last_parts[-2]}\\{last_parts[-1]}"
                    
            # 一般的な短縮
            path_parts = path.split("\\")
            if len(path_parts) > 3:
                first_part = path_parts[0]
                if ":" in first_part:  # ドライブレター
                    first_part = path_parts[0] + "\\" + path_parts[1]
                last_parts = path_parts[-2:]
                return f"{first_part}\\...\\{last_parts[0]}\\{last_parts[1]}"
        
        return path

    def send_direct_message(self, query, response, search_path=None):
        """
        send_ollama_responseのエイリアス（下位互換性のため）
        """
        return self.send_ollama_response(query, response, None, search_path)
