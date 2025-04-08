# ollama_client.py - Ollamaとの通信と応答生成（検索結果表示改善版）
import logging
import requests
import traceback
import re
import os
import json
from datetime import datetime

logger = logging.getLogger(__name__)

def generate_ollama_response(query, ollama_url, ollama_model, ollama_timeout, onedrive_search=None):
    """
    Ollamaを使用して回答を生成する（検索結果表示改善版）

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
        search_results = []
        if onedrive_search and clean_query:
            # 検索ディレクトリのパスを取得（短縮表示用）
            search_path = onedrive_search.base_directory
            short_path = get_shortened_path(search_path)
            
            # 日付を含むかどうかを確認（日報検索に重要）
            has_date = bool(re.search(r'\d{4}年\d{1,2}月\d{1,2}日', clean_query))
            
            # OneDriveから関連情報を取得
            logger.info(f"OneDriveから関連情報を検索: {clean_query} (日付指定: {has_date})")
            try:
                # 検索キーワード抽出（日付やその他のキーワード）
                search_keywords = extract_search_keywords(clean_query)
                logger.info(f"抽出した検索キーワード: {search_keywords}")
                
                # 検索結果をJSON形式で取得 (表示用に整形する前の生の検索結果)
                if search_keywords:
                    search_results = onedrive_search.search_files(search_keywords, max_results=5)
                    logger.info(f"検索結果: {len(search_results)}件")
                
                # 関連コンテンツ取得（表示用に整形）
                relevant_content = onedrive_search.get_relevant_content(clean_query)
                
                if relevant_content and "件の関連ファイルが見つかりました" in relevant_content:
                    onedrive_context = relevant_content
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
                            onedrive_context = f"【検索結果】\n{date_str}の日報は検索ディレクトリ（{short_path}）から見つかりませんでした。"
                    else:
                        onedrive_context = f"【検索結果】\n関連する日報ファイルは検索ディレクトリ（{short_path}）から見つかりませんでした。"
                    
                    logger.info("関連情報は見つかりませんでした")
            except Exception as e:
                logger.error(f"OneDrive検索中にエラーが発生: {str(e)}")
                onedrive_context = f"【検索エラー】\nOneDriveでの検索中にエラーが発生しました。検索パス: {short_path}"

        # Ollamaとは何かを質問されているかを確認
        is_about_ollama = "ollama" in clean_query.lower() and ("とは" in clean_query or "什么" in clean_query or "what" in clean_query.lower())

        # プロンプトの構築
        if is_about_ollama:
            # Ollamaに関する質問の場合、正確な情報を提供
            if onedrive_context:
                prompt = f"""以下の質問に正確に回答してください。Ollamaはビデオ共有プラットフォームではなく、
大規模言語モデル（LLM）をローカル環境で実行するためのオープンソースフレームワークです。

質問: {clean_query}

回答は以下のような正確な情報を含めてください:
- Ollamaは大規模言語モデルをローカルで実行するためのツール
- ローカルコンピュータでLlama、Mistral、Gemmaなどのモデルを実行できる
- プライバシーを保ちながらAI機能を利用できる
- APIを通じて他のアプリケーションから利用できる

また、以下の検索結果も参考にしてください：

{onedrive_context}"""
            else:
                prompt = f"""以下の質問に正確に回答してください。Ollamaはビデオ共有プラットフォームではなく、
大規模言語モデル（LLM）をローカル環境で実行するためのオープンソースフレームワークです。

質問: {clean_query}

回答は以下のような正確な情報を含めてください:
- Ollamaは大規模言語モデルをローカルで実行するためのツール
- ローカルコンピュータでLlama、Mistral、Gemmaなどのモデルを実行できる
- プライバシーを保ちながらAI機能を利用できる
- APIを通じて他のアプリケーションから利用できる"""
        else:
            # 日報に関する質問の特別処理
            is_daily_report_query = "日報" in clean_query or "報告" in clean_query or "活動" in clean_query
            
            if is_daily_report_query and len(search_results) > 0:
                # 日報が見つかった場合の特別処理
                prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

以下の検索結果を参考にして回答してください：

{onedrive_context}

上記の検索結果から抽出された情報のみを使って回答してください。
ファイルの具体的な内容、日付、活動内容などを明確に述べてください。
検索結果に記載されていない情報については「検索結果には記載がありません」と回答してください。

回答の最初に、検索ヒットしたファイル一覧を簡潔に表示してください。
次に、質問に関連する主な情報を要約して回答してください。"""
            elif is_daily_report_query and len(search_results) == 0:
                # 日報が見つからない場合
                date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', clean_query)
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

この日付の日報内容については情報がないため、お答えできません。別の日付をお試しいただくか、システム管理者にお問い合わせください。

回答には、この内容を簡潔にまとめてください。"""
                else:
                    short_path = get_shortened_path(search_path)
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

ご質問の日報データは検索ディレクトリ（{short_path}）から見つかりませんでした。具体的な日付（例：2024年10月26日）を指定すると検索できる可能性があります。
日報検索には、年月日を含めた形で質問していただくとより正確に検索できます。

回答には、この点を丁寧に説明し、日付指定の重要性を伝えてください。"""
            else:
                # 一般的な質問の場合
                if onedrive_context and "件の関連ファイルが見つかりました" in onedrive_context:
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

以下の検索結果を参考にして回答してください：

{onedrive_context}

上記の検索結果と自身の知識を組み合わせて、質問に対して適切に回答してください。
検索結果から得られた情報は「検索結果によると...」などと明示してください。
検索結果に関連情報がない場合は、自身の知識に基づいて回答してください。

回答の最初に、ヒットしたファイル一覧を簡潔に表示してください（最大3件まで）。
検索結果にファイル内容が表示されている場合は、その情報を優先して回答に使用してください。"""
                elif onedrive_context:
                    # 検索はしたが結果が見つからなかった場合
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

検索結果：{onedrive_context}

上記を踏まえつつ、自身の知識に基づいて質問に回答してください。
検索結果が見つからなかった理由について触れた上で、質問に対する一般的な回答を提供してください。"""
                else:
                    # OneDrive検索を実行していない場合は通常の質問として処理
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

                # 生成された結果と検索結果を組み合わせて整形
                final_response = format_response_with_search_results(generated_text, clean_query, search_results, onedrive_context)
                
                logger.info("回答が正常に生成されました")
                return final_response

            else:
                # エラーが発生した場合のフォールバック応答
                return get_fallback_response(clean_query, is_about_ollama, search_path, search_results)

        except requests.exceptions.Timeout:
            logger.error("Ollamaリクエストがタイムアウトしました")
            # タイムアウト時のフォールバック応答
            return get_fallback_response(clean_query, is_about_ollama, search_path, search_results)

        except requests.exceptions.ConnectionError:
            logger.error("Ollamaサーバーに接続できませんでした")
            return get_fallback_response(clean_query, is_about_ollama, search_path, search_results)

    except Exception as e:
        logger.error(f"回答生成中にエラーが発生しました: {str(e)}")
        # スタックトレースをログに記録
        logger.error(traceback.format_exc())
        return f"エラーが発生しました: {str(e)}"

def format_response_with_search_results(generated_text, query, search_results, onedrive_context=""):
    """
    生成された応答と検索結果を組み合わせて整形する
    
    Args:
        generated_text: Ollamaから生成されたテキスト
        query: ユーザーの質問
        search_results: 検索結果リスト
        onedrive_context: OneDrive検索コンテキスト
        
    Returns:
        整形された最終応答
    """
    # 検索結果がない場合は生成されたテキストをそのまま返す
    if not search_results:
        return generated_text
    
    # 生成されたテキストが既に検索結果のサマリーを含んでいる場合はそのまま返す
    if "件の関連ファイルが見つかりました" in generated_text or "ヒットしたファイル" in generated_text:
        return generated_text
    
    # 検索結果のファイル情報リストを作成
    file_list = ""
    for i, result in enumerate(search_results[:3]):  # 最初の3件のみ表示
        file_name = result.get('name', '不明なファイル')
        modified = result.get('modified', '不明な日時')
        size = format_file_size(result.get('size', 0))
        file_list += f"{i+1}. {file_name} (更新日時: {modified}, サイズ: {size})\n"
    
    # 3件以上ある場合の表示
    if len(search_results) > 3:
        file_list += f"... 他 {len(search_results) - 3} 件のファイルが見つかりました\n"
    
    # 日報関連の質問かどうかを判定
    is_daily_report_query = "日報" in query or "報告" in query
    
    # 日報関連の質問の場合は、検索結果を目立つように表示
    if is_daily_report_query:
        final_response = f"""【検索結果】
{file_list}
-------------------
{generated_text}"""
    else:
        # 一般的な質問の場合は、生成されたテキストを優先し、最後に検索結果を添付
        final_response = f"""{generated_text}

-------------------
【参考検索結果】
{file_list}"""
    
    return final_response

def extract_search_keywords(query):
    """
    クエリから検索キーワードを抽出する
    
    Args:
        query: ユーザーの質問
        
    Returns:
        検索キーワードのリスト
    """
    keywords = []
    
    # 日付形式（YYYY年MM月DD日）を抽出
    date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', query)
    if date_match:
        year = date_match.group(1)
        month = date_match.group(2)
        day = date_match.group(3)
        keywords.append(f"{year}年{month}月{day}日")
    
    # 重要な名詞を抽出（簡易的なアプローチ）
    stop_words = ["について", "とは", "の", "を", "に", "は", "で", "が", "と", "から", "へ", "より", 
                 "内容", "知りたい", "あったのか", "何", "教えて", "どのような", "どんな", "ありました",
                 "the", "a", "an", "and", "or", "of", "to", "in", "for", "on", "by"]
    
    # クエリを単語に分割して重要そうな単語を抽出
    for word in query.split():
        # 単語をクリーニング
        clean_word = word.strip(',.;:!?()[]{}"\'')
        # 長さ2以上で、ストップワードに含まれない単語を抽出
        if clean_word and len(clean_word) > 2 and clean_word.lower() not in stop_words:
            # 既に日付として抽出されていなければ追加
            if date_match and clean_word not in date_match.group(0):
                keywords.append(clean_word)
            elif not date_match:
                keywords.append(clean_word)
    
    # 「日報」を含む質問で、キーワードに「日報」がない場合は追加
    if "日報" in query and not any(k for k in keywords if "日報" in k):
        keywords.append("日報")
    
    # キーワードがまだ少ない場合は、より短い単語も追加
    if len(keywords) < 2:
        for word in query.split():
            clean_word = word.strip(',.;:!?()[]{}"\'')
            if clean_word and len(clean_word) == 2 and clean_word.lower() not in stop_words:
                if clean_word not in keywords:
                    keywords.append(clean_word)
    
    return keywords

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

def format_file_size(size_in_bytes):
    """
    ファイルサイズを読みやすい形式に変換
    
    Args:
        size_in_bytes: バイト単位のサイズ
        
    Returns:
        フォーマットされたサイズ文字列
    """
    # バイト→KB→MB→GBの変換
    if size_in_bytes < 1024:
        return f"{size_in_bytes} バイト"
    elif size_in_bytes < 1024 * 1024:
        return f"{size_in_bytes / 1024:.1f} KB"
    elif size_in_bytes < 1024 * 1024 * 1024:
        return f"{size_in_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_in_bytes / (1024 * 1024 * 1024):.1f} GB"

def get_fallback_response(query, is_about_ollama=False, search_path="", search_results=None):
    """
    タイムアウトやエラー時に使用するフォールバック応答を返す

    Args:
        query: ユーザーの質問
        is_about_ollama: Ollamaに関する質問かどうか
        search_path: 検索したパス
        search_results: 検索結果リスト

    Returns:
        フォールバック応答
    """
    short_path = get_shortened_path(search_path)
    
    # 検索結果があれば、ファイル一覧を表示
    search_info = ""
    if search_results:
        search_info = "検索された関連ファイル:\n"
        for i, result in enumerate(search_results[:5]):
            file_name = result.get('name', '不明なファイル')
            modified = result.get('modified', '不明な日時')
            search_info += f"{i+1}. {file_name} (更新日時: {modified})\n"
        search_info += "\n"
    
    if is_about_ollama:
        return f"""{search_info}Ollamaは、大規模言語モデル（LLM）をローカル環境で実行するためのオープンソースフレームワークです。

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
            return f"""{search_info}{year}年{month}月{day}日の日報データを取得できませんでした。サーバーの応答に問題があるか、該当する日報が検索ディレクトリ（{short_path}）に存在しない可能性があります。時間をおいて再度お試しいただくか、システム管理者にお問い合わせください。"""
        else:
            return f"""{search_info}日報データを取得できませんでした。具体的な日付（例：2024年10月26日）を指定して再度お試しください。検索ディレクトリ: {short_path}"""
    else:
        return f"""{search_info}「{query}」についてのご質問ありがとうございます。ただいまOllamaサーバーの処理に時間がかかっています。少し時間をおいてから再度お試しいただくか、より具体的な質問を入力してください。"""
