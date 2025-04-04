# teams_webhook.py - 改良されたTeams通知機能
import requests
import logging
import json
from datetime import datetime
import traceback
import re
import os

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

    def send_ollama_response(self, query, response, search_results=None, search_path=None):
        """
        Ollamaの応答をTeams Workflowに送信する（改良版）

        Args:
            query: ユーザーの質問
            response: Ollamaからの応答
            search_results: 検索結果のリスト（オプション）
            search_path: 検索に使用したディレクトリパス（オプション）

        Returns:
            dict: 送信結果
        """
        try:
            # 検索パスの短縮表示
            short_path = self._get_shortened_path(search_path) if search_path else "設定されたディレクトリ"
            
            # 現在の日時
            now = datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')
            
            # 検索結果のサマリーを作成
            files_info = ""
            if search_results and len(search_results) > 0:
                files_info = "\n".join([f"- {r.get('name')} ({r.get('modified')})" for r in search_results[:3]])
                if len(search_results) > 3:
                    files_info += f"\n- 他 {len(search_results)-3} 件のファイル"
            else:
                files_info = "関連ファイルは見つかりませんでした"
            
            # 拡張されたTeamsカード形式
            enhanced_card = self.create_enhanced_card(query, response, search_results, search_path)
            
            # リクエストヘッダー
            headers = {
                'Content-Type': 'application/json'
            }

            # 1. 拡張カード形式で試行
            logger.debug(f"Logic Apps送信ペイロード(拡張カード): {json.dumps(enhanced_card)[:300]}...")

            try:
                r = requests.post(
                    self.webhook_url, 
                    json=enhanced_card, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答(拡張カード): {r.status_code}, {r.text[:100] if r.text else '空のレスポンス'}")

                if r.status_code >= 200 and r.status_code < 300:
                    logger.info(f"拡張カード形式でのLogic Apps通知送信成功: {r.status_code}")
                    return {"status": "success", "code": r.status_code, "format": "拡張カード"}
                else:
                    logger.warning(f"拡張カード形式での送信失敗: {r.status_code}。従来形式で再試行します。")

            except Exception as e:
                logger.warning(f"拡張カード形式送信エラー: {str(e)}。従来形式で再試行します。")
            
            # 従来形式のペイロード
            legacy_payload = {
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
            
            # バックアップ用のシンプルなペイロード
            simple_payload = {
                "text": f"### Ollama回答\n\n**質問**: {query}\n\n**検索対象**: {short_path}\n\n{response}\n\n*回答生成時刻: {now}*"
            }

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
    
    def create_enhanced_card(self, query, response, search_results, search_path):
        """
        改善されたTeams通知カードを作成する

        Args:
            query: ユーザーの質問
            response: Ollamaからの応答
            search_results: 検索結果リスト
            search_path: 検索パス
            
        Returns:
            dict: カードデータ
        """
        now = datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')
        short_path = self._get_shortened_path(search_path)
        
        # 検索結果サマリー
        found_files = ""
        if search_results and len(search_results) > 0:
            for i, result in enumerate(search_results[:3]):  # 上位3件表示
                file_name = result.get('name', '不明')
                modified = result.get('modified', '不明')
                found_files += f"- {file_name} (更新: {modified})\n"
            
            if len(search_results) > 3:
                found_files += f"- 他 {len(search_results)-3} 件のファイル\n"
        else:
            found_files = "関連ファイルは見つかりませんでした"
        
        return {
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
                                "type": "FactSet",
                                "facts": [
                                    {"title": "検索場所", "value": short_path},
                                    {"title": "検索結果", "value": f"{len(search_results) if search_results else 0}件のファイル"}
                                ]
                            },
                            {
                                "type": "TextBlock",
                                "text": "見つかったファイル:",
                                "wrap": True,
                                "size": "Small"
                            },
                            {
                                "type": "TextBlock",
                                "text": found_files,
                                "wrap": True,
                                "size": "Small",
                                "isSubtle": True
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

    def send_direct_message(self, query, response, search_results=None, search_path=None):
        """
        send_ollama_responseのエイリアス（下位互換性のため）
        """
        return self.send_ollama_response(query, response, search_results, search_path)
