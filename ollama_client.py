# ollama_client.py - Ollamaとの通信と応答生成（日本語クエリ改善版2.0）
import logging
import requests
import traceback
import re
import os

logger = logging.getLogger(__name__)

def generate_ollama_response(query, ollama_url, ollama_model, ollama_timeout, onedrive_search=None):
    """
    Ollamaを使用して回答を生成する（日本語クエリ改善版）

    Args:
        query: ユーザーからの質問
        ollama_url: OllamaのURL
        ollama_model: 使用するOllamaモデル
        ollama_timeout: リクエストのタイムアウト時間（秒）
        onedrive_search: OneDriveSearch インスタンス（Noneの場合は検索しない）

    Returns:
        生成された回答
    """
    try:
        logger.info(f"Ollama回答を生成します: {query}")

        # @ollama質問 タグを除去
        clean_query = query.replace('ollama質問', '').strip()
        logger.info(f"クリーニング後のクエリ: {clean_query}")

        # OneDrive検索が有効かつクエリがある場合は関連情報を検索
        onedrive_context = ""
        search_path = ""
        if onedrive_search and clean_query:
            # 検索ディレクトリのパスを取得（短縮表示用）
            search_path = onedrive_search.base_directory
            short_path = get_shortened_path(search_path)
            
            # 日付を含むかどうかを確認（日報検索に重要）
            date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', clean_query)
            has_date = bool(date_match)
            
            # 日付の詳細を抽出
            date_specific_keywords = []
            if has_date:
                year = date_match.group(1)
                month = date_match.group(2).zfill(2)
                day = date_match.group(3).zfill(2)
                
                # 様々な日付形式を追加
                date_pattern = f"{year}{month}{day}"  # YYYYMMDD
                file_date_pattern = f"_{year}{month}{day}"  # _YYYYMMDD（ファイル名パターン）
                date_str = f"{year}年{month}月{day}日"
                
                # 年度フォルダも考慮
                fiscal_year = onedrive_search.get_fiscal_year_folder(year, month)
                
                date_specific_keywords = [
                    date_pattern,
                    file_date_pattern,
                    date_str,
                    fiscal_year
                ]
                
                logger.info(f"日付特化検索キーワード: {date_specific_keywords}")
            
            # OneDriveから関連情報を取得
            logger.info(f"OneDriveから関連情報を検索: {clean_query} (日付指定: {has_date})")
            try:
                # 日付特化検索を実行
                if has_date and date_specific_keywords:
                    # デバッグ情報を収集
                    debug_info = onedrive_search.debug_search(clean_query)
                    logger.debug(f"検索デバッグ情報: {debug_info}")
                    
                    # 複数キーワードで検索
                    relevant_content = onedrive_search.get_relevant_content_with_multiple_keywords(date_specific_keywords)
                else:
                    # 通常検索
                    relevant_content = onedrive_search.get_relevant_content(clean_query)
                
                if relevant_content and "件の関連ファイルが見つかりました" in relevant_content:
                    onedrive_context = f"\n\n参考資料（OneDriveから取得 - {short_path}）:\n{relevant_content}"
                    logger.info(f"OneDriveから関連情報を取得: {len(onedrive_context)}文字")
                else:
                    # ファイルが見つからない場合は見つからないことを明示
                    if has_date:
                        # 日付から数字だけを抽出して表示
                        date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', clean_query)
                        if date_match:
                            year = date_match.group(1)
                            month = date_match.group(2)
                            day = date_match.group(3)
                            date_str = f"{year}年{month}月{day}日"
                            fiscal_year = onedrive_search.get_fiscal_year_folder(year, month)
                            onedrive_context = f"\n\n注意: {date_str}の日報は検索ディレクトリ（{short_path}）から見つかりませんでした。"
                            
                            # 検索結果がない場合、ディレクトリ内のファイル例を表示
                            try:
                                logger.info(f"年度フォルダを確認: {fiscal_year}")
                                ps_command = f"""
                                Get-ChildItem -Path "{search_path}" -Recurse -Filter "*{fiscal_year}*" -Directory | 
                                Select-Object -First 1 -ExpandProperty FullName | ForEach-Object {{
                                    Get-ChildItem -Path $_ -File | Select-Object -First 5 | ForEach-Object {{ $_.Name }}
                                }}
                                """
                                
                                sample_files = subprocess.check_output(["powershell", "-Command", ps_command], text=True).strip()
                                if sample_files:
                                    onedrive_context += f"\n\n{fiscal_year}フォルダ内のファイル例:\n{sample_files}"
                            except:
                                pass
                    else:
                        onedrive_context = f"\n\n注意: 関連する日報ファイルは検索ディレクトリ（{short_path}）から見つかりませんでした。"
                    
                    logger.info("関連情報は見つかりませんでした")
            except Exception as e:
                logger.error(f"OneDrive検索中にエラーが発生: {str(e)}")
                logger.error(traceback.format_exc())
                onedrive_context = f"\n\n注意: OneDriveでの検索中にエラーが発生しました。検索パス: {short_path}"

        # Ollamaとは何かを質問されているかを確認
        is_about_ollama = "ollama" in clean_query.lower() and ("とは" in clean_query or "什么" in clean_query or "what" in clean_query.lower())

        # プロンプトの構築（OneDriveコンテキストを含む）
        if is_about_ollama:
            # Ollamaに関する質問の場合、正確な情報を提供
            prompt = f"""以下の質問に正確に回答してください。Ollamaはビデオ共有プラットフォームではなく、
大規模言語モデル（LLM）をローカル環境で実行するためのオープンソースフレームワークです。

質問: {clean_query}

回答は以下のような正確な情報を含めてください:
- Ollamaは大規模言語モデルをローカルで実行するためのツール
- ローカルコンピュータでLlama、Mistral、Gemmaなどのモデルを実行できる
- プライバシーを保ちながらAI機能を利用できる
- APIを通じて他のアプリケーションから利用できる{onedrive_context}"""
        else:
            # 日報に関する質問の特別処理
            if "日報" in clean_query and onedrive_context and "件の関連ファイルが見つかりました" in onedrive_context:
                prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{onedrive_context}

上記の参考資料を基に具体的に回答してください。特に日付や内容を明確に述べてください。
参考資料に示された情報のみを使用し、ない情報は「資料には記載がありません」と正直に答えてください。"""
            elif "日報" in clean_query and not ("件の関連ファイルが見つかりました" in onedrive_context):
                # 日報が見つからない場合
                date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', clean_query)
                if date_match:
                    year = date_match.group(1)
                    month = date_match.group(2)
                    day = date_match.group(3)
                    short_path = get_shortened_path(search_path)
                    fiscal_year = onedrive_search.get_fiscal_year_folder(year, month) if onedrive_search else f"{int(year)-1 if int(month)<4 else year}年度"
                    
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{year}年{month}月{day}日の日報データは検索ディレクトリ（{short_path}）から見つかりませんでした。
可能性のある理由:
1. 指定された日付の日報が存在しない
2. ファイル名のフォーマットが「業務日報-近藤_{year}{month}{day}.pdf」のような形式でない
3. 日報が{fiscal_year}フォルダではなく別のフォルダに保存されている
4. アクセス権限の問題でファイルが見つけられない

日報が見つからなかったため、この日の日報内容についてはお答えできません。別の日付をお試しいただくか、ファイル名のパターンを変えて検索してみてください。"""
                else:
                    short_path = get_shortened_path(search_path)
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

ご質問の日報データは検索ディレクトリ（{short_path}）から見つかりませんでした。具体的な日付（例：2025年3月15日）を指定すると検索できる可能性があります。
日報検索には、年月日を含めた形で質問していただくとより正確に検索できます。
例: 「2025年3月15日の日報内容を教えてください」"""
            else:
                # OneDriveコンテキストを含むプロンプト
                if onedrive_context:
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{onedrive_context}

上記の参考資料を基に質問に回答してください。参考資料に関連情報がない場合は、あなたの知識を使って回答してください。"""
                else:
                    prompt = clean_query

        # Ollamaへのリクエストを構築（パラメータを調整）
        payload = {
            "model": ollama_model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "num_predict": 1024,      # 生成するトークン数を増加（長めの回答）
                "temperature": 0.7,       # バランスの取れた温度
                "top_k": 40,              # 効率化のため選択肢を制限
                "top_p": 0.9,             # 確率分布を制限して効率化
                "num_ctx": 4096,          # コンテキスト長を増加（OneDriveからの情報に対応）
                "seed": 42                # 一貫性のある回答のためのシード値
            }
        }

        logger.info(f"Ollamaにリクエストを送信: モデル={ollama_model}")

        # タイムアウト設定を追加
        try:
            # 環境変数で設定されたタイムアウトを使用
            response = requests.post(ollama_url, json=payload, timeout=ollama_timeout)
            logger.info(f"Ollama応答ステータスコード: {response.status_code}")

            if response.status_code == 200:
                result = response.json()
                logger.info(f"Ollama応答JSON: {result}")

                generated_text = result.get('response', '申し訳ありませんが、回答を生成できませんでした。')
                logger.info(f"生成されたテキスト: {generated_text[:100]}...")

                # レスポンスが空でないことを確認
                if not generated_text.strip():
                    generated_text = "申し訳ありませんが、有効な回答を生成できませんでした。"

                logger.info("回答が正常に生成されました")
                return generated_text

            else:
                # エラーが発生した場合のフォールバック応答
                return get_fallback_response(clean_query, is_about_ollama, search_path)

        except requests.exceptions.Timeout:
            logger.error("Ollamaリクエストがタイムアウトしました")
            # タイムアウト時のフォールバック応答
            return get_fallback_response(clean_query, is_about_ollama, search_path)

        except requests.exceptions.ConnectionError:
            logger.error("Ollamaサーバーに接続できませんでした")
            return get_fallback_response(clean_query, is_about_ollama, search_path)

    except Exception as e:
        logger.error(f"回答生成中にエラーが発生しました: {str(e)}")
        # スタックトレースをログに記録
        logger.error(traceback.format_exc())
        return f"エラーが発生しました: {str(e)}"

def get_shortened_path(path):
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