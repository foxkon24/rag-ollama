# ollama_client.py - Ollamaとの通信と応答生成（日本語クエリ改善版）
import logging
import requests
import traceback
import re
import os
import sys

logger = logging.getLogger(__name__)

def clean_response_text(text):
    """
    Ollamaからの応答テキストをクリーニングして文字化けを修正
    
    Args:
        text: 生のレスポンステキスト
        
    Returns:
        クリーニングされたテキスト
    """
    # 韓国語の「없습니다」を「ありません」に置換
    text = text.replace("없습니다", "ありません")
    
    # その他の非ASCII文字をチェック
    import re
    
    # 日本語、英数字、一般的な記号以外の文字を検出
    strange_chars = re.findall(r'[^\u0000-\u007F\u3000-\u30FF\u4E00-\u9FFF\u3040-\u309F\uFF00-\uFFEF\s.,!?;:()\[\]{}\'\"。、！？「」『』（）［］｛｝＜＞・…ー－]', text)
    
    if strange_chars:
        logger.warning(f"レスポンステキストに不明な文字が含まれています: {set(strange_chars)}")
        
        # 不明な文字を空白に置換
        for char in set(strange_chars):
            if char != " ":  # 空白は保持
                text = text.replace(char, " ")
    
    return text

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
            has_date = bool(re.search(r'\d{4}年\d{1,2}月\d{1,2}日', clean_query))
            
            # OneDriveから関連情報を取得
            logger.info(f"OneDriveから関連情報を検索: {clean_query} (日付指定: {has_date})")
            try:
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
                            onedrive_context = f"\n\n注意: {date_str}の日報は検索ディレクトリ（{short_path}）から見つかりませんでした。"
                    else:
                        onedrive_context = f"\n\n注意: 関連する日報ファイルは検索ディレクトリ（{short_path}）から見つかりませんでした。"
                    
                    logger.info("関連情報は見つかりませんでした")
            except Exception as e:
                logger.error(f"OneDrive検索中にエラーが発生: {str(e)}")
                onedrive_context = f"\n\n注意: OneDriveでの検索中にエラーが発生しました。検索パス: {short_path}"

        # Ollamaとは何かを質問されているかを確認
        is_about_ollama = "ollama" in clean_query.lower() and ("とは" in clean_query or "什么" in clean_query or "what" in clean_query.lower())

        # プロンプトの構築
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
            if "日報" in clean_query and onedrive_context:
                # ファイルが見つかったか確認
                files_found = "件の関連ファイルが見つかりました" in onedrive_context
                
                # 特定の日付の日報を探しているか確認
                date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', clean_query)
                
                if files_found:
                    # PDFテキスト抽出がサポートされていないケースを処理
                    if "このPDFファイルの内容を直接表示することはできませんが、ファイルは確かに存在しています" in onedrive_context:
                        # 日付情報を含むPDFファイルが存在することを明示
                        pdf_file_match = re.search(r'ファイル名: (業務日報-近藤_\d{8}\.pdf)', onedrive_context)
                        if pdf_file_match:
                            pdf_filename = pdf_file_match.group(1)
                            prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{onedrive_context}

上記の情報から、質問された日付の日報ファイルは存在していますが、PDFの内容を直接抽出することができません。
ファイルの存在は確認できていますので、内容を確認するには直接PDFファイルを開く必要があります。"""
                        else:
                            # 通常の処理
                            prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{onedrive_context}

上記の参考資料を基に具体的に回答してください。特に日付や内容を明確に述べてください。
参考資料に示された情報のみを使用し、ない情報は「資料には記載がありません」と正直に答えてください。"""
                    else:
                        # 通常の処理
                        prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{onedrive_context}

上記の参考資料を基に具体的に回答してください。特に日付や内容を明確に述べてください。
参考資料に示された情報のみを使用し、ない情報は「資料には記載がありません」と正直に答えてください。"""
                elif not ("件の関連ファイルが見つかりました" in onedrive_context):
                    # 日報が見つからない場合
                    if date_match:
                        year = date_match.group(1)
                        month = date_match.group(2)
                        day = date_match.group(3)
                        short_path = get_shortened_path(search_path)
                        prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{year}年{month}月{day}日の日報データは検索ディレクトリ（{short_path}）から見つかりませんでした。
以下のいずれかの理由が考えられます：
1. 指定された日付の日報が存在しない
2. 検索可能な場所に保存されていない
3. ファイル名が通常と異なる形式で保存されている
4. アクセス権限の問題でファイルが見つけられない

この日付の日報内容については情報がないため、お答えできません。別の日付をお試しいただくか、システム管理者にお問い合わせください。"""
                    else:
                        short_path = get_shortened_path(search_path)
                        prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

ご質問の日報データは検索ディレクトリ（{short_path}）から見つかりませんでした。具体的な日付（例：2024年10月26日）を指定すると検索できる可能性があります。
日報検索には、年月日を含めた形で質問していただくとより正確に検索できます。"""
                else:
                    # OneDriveコンテキストを含むプロンプト
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{onedrive_context}

上記の参考資料を基に質問に回答してください。参考資料に関連情報がない場合は、あなたの知識を使って回答してください。"""
            else:
                # 通常のプロンプト処理
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
                
                # テキストをクリーニング
                generated_text = clean_response_text(generated_text)
                
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
                
        # 一般的な短縮
        path_parts = path.split("\\")
        if len(path_parts) > 3:
            first_part = path_parts[0]
            if ":" in first_part:  # ドライブレター
                first_part = path_parts[0] + "\\" + path_parts[1]
            last_parts = path_parts[-2:]
            return f"{first_part}\\...\\{last_parts[0]}\\{last_parts[1]}"
    
    return path

def get_fallback_response(query, is_about_ollama=False, search_path=""):
    """
    タイムアウトやエラー時に使用するフォールバック応答を返す（文字化け対策版）
    """
    short_path = get_shortened_path(search_path)
    
    if is_about_ollama:
        return """Ollamaは、大規模言語モデル（LLM）をローカル環境で実行するためのオープンソースフレームワークです。

主な特徴:
1. ローカル実行: インターネット接続不要で自分のコンピュータ上でAIモデルを実行できます
2. 複数モデル対応: Llama2, Llama3, Mistral, Gemmaなど様々なモデルを利用できます
3. APIインターフェース: 他のアプリケーションから簡単に利用できるRESTful APIを提供します
4. 軽量設計: 一般的なハードウェアでも動作するよう最適化されています

Ollamaを使うと、プライバシーを保ちながら、AI機能を様々なソフトウェアに統合できます。
詳細は公式サイト: https://ollama.ai/ をご覧ください。"""
    
    # 日報に関する質問のフォールバック
    elif "日報" in query:
        date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', query)
        if date_match:
            year = date_match.group(1)
            month = date_match.group(2)
            day = date_match.group(3)
            return f"{year}年{month}月{day}日の日報データを取得できませんでした。サーバーの応答に問題があるか、該当する日報が検索ディレクトリ（{short_path}）に存在しない可能性があります。時間をおいて再度お試しいただくか、システム管理者にお問い合わせください。"
        else:
            return f"日報データを取得できませんでした。具体的な日付（例：2024年10月26日）を指定して再度お試しください。検索ディレクトリ: {short_path}"
    else:
        return f"「{query}」についてのご質問ありがとうございます。ただいまOllamaサーバーの処理に時間がかかっています。少し時間をおいてから再度お試しいただくか、より具体的な質問を入力してください。"
