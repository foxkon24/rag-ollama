# teams_webhook.py - Logic Apps Workflowに対応した通知機能（再修正版）
import requests
import logging
import json
from datetime import datetime
import traceback

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # 詳細なログを有効化

class TeamsWebhook:
    def __init__(self, webhook_url):
        """
        Teams Webhookクラスの初期化

        Args:
            webhook_url: Teams Workflow URL (Logic Apps URL)
        """
        self.webhook_url = webhook_url
        logger.info(f"Teams Workflowを初期化: {webhook_url[:30]}...")

    def send_ollama_response(self, query, response, conversation_data=None):
        """
        Ollamaの応答をTeams Workflowに送信する

        Args:
            query: ユーザーの質問
            response: Ollamaからの応答
            conversation_data: 会話データ（オプション）

        Returns:
            dict: 送信結果
        """
        try:
            # エラーメッセージから判断した正しいペイロード形式
            # 'attachments'配列が直接ルートレベルにある形式
            root_attachments_payload = {
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
                                    "wrap": True
                                },
                                {
                                    "type": "TextBlock",
                                    "text": response,
                                    "wrap": True
                                }
                            ],
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "version": "1.0"
                        }
                    }
                ]
            }

            # バックアップ用のシンプルなペイロード
            simple_payload = {
                "text": f"### Ollama回答\n\n**質問**: {query}\n\n{response}"
            }

            # 旧形式のペイロード（既存形式）
            legacy_payload = {
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
                                        "wrap": True
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": response,
                                        "wrap": True
                                    }
                                ],
                                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                                "version": "1.0"
                            }
                        }
                    ]
                }
            }

            # リクエストヘッダー
            headers = {
                'Content-Type': 'application/json'
            }

            # 1. まず新しいルートレベルのattachments形式で試行
            logger.debug(f"Logic Apps送信ペイロード(ルートレベルattachments): {json.dumps(root_attachments_payload)[:300]}...")

            try:
                r = requests.post(
                    self.webhook_url, 
                    json=root_attachments_payload, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答(ルートレベルattachments): {r.status_code}, {r.text[:100] if r.text else '空のレスポンス'}")

                if r.status_code >= 200 and r.status_code < 300:
                    logger.info(f"ルートレベルattachments形式でのLogic Apps通知送信成功: {r.status_code}")
                    return {"status": "success", "code": r.status_code, "format": "ルートレベルattachments"}
                else:
                    logger.warning(f"ルートレベルattachments形式での送信失敗: {r.status_code}。従来形式で再試行します。")

            except Exception as e:
                logger.warning(f"ルートレベルattachments形式送信エラー: {str(e)}。従来形式で再試行します。")

            # 2. 従来形式で試行
            logger.debug(f"Logic Apps送信ペイロード(従来形式): {json.dumps(legacy_payload)[:300]}...")

            try:
                r2 = requests.post(
                    self.webhook_url, 
                    json=legacy_payload, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答(従来形式): {r2.status_code}, {r2.text[:100] if r2.text else '空のレスポンス'}")

                if r2.status_code >= 200 and r2.status_code < 300:
                    logger.info(f"従来形式でのLogic Apps通知送信成功: {r2.status_code}")
                    return {"status": "success", "code": r2.status_code, "format": "従来形式"}
                else:
                    logger.warning(f"従来形式での送信失敗: {r2.status_code}。シンプル形式で再試行します。")

            except Exception as e2:
                logger.warning(f"従来形式送信エラー: {str(e2)}。シンプル形式で再試行します。")

            # 3. シンプル形式で試行（最後の手段）
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

    def send_direct_message(self, query, response):
        """
        send_ollama_responseのエイリアス（下位互換性のため）
        """
        return self.send_ollama_response(query, response)
