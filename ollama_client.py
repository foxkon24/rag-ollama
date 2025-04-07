# ollama_client.py - Ollamaとの通信と応答生成（コンテンツベース回答改善版）
import logging
import requests
import traceback
import re
import os

logger = logging.getLogger(__name__)

def generate_ollama_response(query, ollama_url, ollama_model, ollama_timeout, onedrive_search=None):
    """
    Ollamaを使用して回答を生成する（コンテンツベース回答改善版）

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
        retrieved_files = []
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
                
                # 関連ファイル情報の抽出（メタデータ表示用）
                file_matches = re.findall(r'=== ファイル \d+: (.+?) ===\n更新日時: (.+?)\n', relevant_content)
                retrieved_files = [{"name": name, "date": date} for name, date in file_matches]
                
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
                            onedrive_context = f"注意: {date_str}の日報は検索ディレクトリ（{short_path}）から見つかりませんでした。"
                    else:
                        onedrive_context = f"注意: 関連する日報ファイルは検索ディレクトリ（{short_path}）から見つかりませんでした。"
                    
                    logger.info("関連情報は見つかりませんでした")
            except Exception as e:
                logger.error(f"OneDrive検索中にエラーが発生: {str(e)}")
                onedrive_context = f"注意: OneDriveでの検索中にエラーが発生しました。検索パス: {short_path}"

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
- APIを通じて他のアプリケーションから利用できる"""
        else:
            # 日報に関する質問の特別処理
            if "日報" in clean_query and onedrive_context and "件の関連ファイルが見つかりました" in onedrive_context:
                prompt = f"""以下の質問に、提供された資料の内容に基づいて具体的に回答してください。

質問: {clean_query}

### 参考資料:
{onedrive_context}

### 指示:
1. 上記の参考資料を精査し、質問に関連する具体的な情報を探してください。
2. 資料から得られる事実のみに基づいて回答を作成してください。
3. 日報の内容を詳細に分析し、質問に関連する重要な情報を抽出してください。
4. 以下の2部構成で回答を提供してください:
   【回答】セクション:
   - 質問に対する直接的な回答を示す
   - 参考資料から見つかった具体的な詳細情報を提示する
   - 情報源として参照したファイル名と日付を明記する
   
   【検索結果原文】セクション:
   - ここに元の検索結果をそのまま含める（このセクションは後で自動的に追加されるので、あなたは【回答】セクションのみ作成してください）
5. 資料に情報がない場合は、「資料には記載がありません」と明示してください。
6. 推測や一般的な知識による補完は行わず、資料に記載されている事実のみを使用してください。

あなたの回答は【回答】の見出しで始め、検索結果原文は自動的に追加されるので含めないでください。"""

            elif "日報" in clean_query and not ("件の関連ファイルが見つかりました" in onedrive_context):
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

この日付の日報内容については情報がないため、お答えできません。別の日付をお試しいただくか、システム管理者にお問い合わせください。"""
                else:
                    short_path = get_shortened_path(search_path)
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

ご質問の日報データは検索ディレクトリ（{short_path}）から見つかりませんでした。具体的な日付（例：2024年10月26日）を指定すると検索できる可能性があります。
日報検索には、年月日を含めた形で質問していただくとより正確に検索できます。"""
            else:
                # OneDriveコンテキストを含むプロンプト
                if onedrive_context and "件の関連ファイルが見つかりました" in onedrive_context:
                    # ファイルのメタデータ情報を追加
                    file_info = ""
                    if retrieved_files:
                        file_info = "検索結果ファイル一覧:\n"
                        for i, file in enumerate(retrieved_files):
                            file_info += f"{i+1}. {file['name']} (更新: {file['date']})\n"
                    
                    prompt = f"""以下の質問に対して、提供された資料の内容に基づいて詳細かつ具体的に回答してください。

質問: {clean_query}

{file_info}

### 参考資料:
{onedrive_context}

### 指示:
1. 上記の参考資料を詳細に分析し、質問に対する適切な情報を抽出してください。
2. 資料から得られる具体的な事実に基づいて回答を作成してください。
3. 重要な情報、具体的な数字、日付、名前などの詳細を正確に引用してください。
4. 以下の2部構成で回答を提供してください:
   【回答】セクション:
   - 質問への直接的な回答
   - 参考資料から抽出した関連する重要な詳細情報
   - 必要に応じて、情報の出典となったファイル名
   
   【検索結果原文】セクション:
   - ここには元の検索結果がそのまま含められます（このセクションは後で自動的に追加されるので、あなたは【回答】セクションのみ作成してください）
5. 資料に記載されていない情報については、「この点については資料に記載がありません」と明示してください。

あなたの回答は【回答】の見出しで始め、検索結果原文は自動的に追加されるので含めないでください。"""
                elif onedrive_context:
                    # 検索結果がない場合のプロンプト
                    prompt = f"""以下の質問に日本語で丁寧に回答してください。

質問: {clean_query}

{onedrive_context}

上記の注意事項を踏まえて、あなたの知識に基づいて質問に回答してください。"""
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
                
                # OneDriveから検索結果があった場合の処理
                if onedrive_context and "件の関連ファイルが見つかりました" in onedrive_context and not is_about_ollama:
                    # 【回答】セクションの確認
                    if "【回答】" not in generated_text:
                        generated_text = f"【回答】\n{generated_text}"
                    
                    # 検索結果原文を追加
                    search_result_raw = "\n\n【検索結果原文】\n" + onedrive_context
                    
                    # PDFファイルの情報がある場合、見やすく整形
                    if ".pdf" in search_result_raw.lower() and "pdf名:" in search_result_raw.lower():
                        # PDFファイル情報の整形（PDFの内容部分を見やすく）
                        pdf_sections = re.findall(r'(=== ファイル \d+:.+?PDF名:.+?-+\n)(.*?)(?===|\Z)', search_result_raw, re.DOTALL)
                        if pdf_sections:
                            formatted_search_result = ""
                            for header, content in pdf_sections:
                                formatted_search_result += header + "\n"
                                
                                # 抽出されたPDFテキストの整形
                                if "PDFページ数:" in content:
                                    # 新しいPDF抽出形式
                                    formatted_search_result += content
                                else:
                                    # 従来の形式またはPowerShellからの出力
                                    formatted_search_result += content.strip() + "\n\n"
                            
                            # 元の検索結果の最初の部分（件数情報など）を保持
                            first_file_pos = search_result_raw.find("=== ファイル 1:")
                            if first_file_pos > 0:
                                header_part = search_result_raw[:first_file_pos]
                                search_result_raw = header_part + formatted_search_result
                    
                    # 長すぎる場合は省略
                    if len(generated_text) + len(search_result_raw) > 8000:
                        # 長すぎる場合は省略するが、PDF情報部分は保持する
                        max_length = 4000
                        if ".pdf" in search_result_raw.lower() and "pdf名:" in search_result_raw.lower():
                            # PDFファイル情報の前後を保持
                            pdf_intro_pos = search_result_raw.lower().find("pdf名:")
                            if pdf_intro_pos > 0:
                                # PDF情報の前半部分
                                first_part = search_result_raw[:pdf_intro_pos+500]
                                # 残りの部分は省略
                                search_result_raw = first_part + "\n...(省略)...\n"
                        else:
                            # 単純に切り詰め
                            search_result_raw = search_result_raw[:max_length] + "\n...(省略)...\n"
                    
                    # 参照情報を回答の最後に追加
                    generated_text += search_result_raw
                
                # 検索したファイル情報を最後に追加（検索結果原文を追加していない場合のみ）
                elif retrieved_files and not is_about_ollama:
                    source_info = "\n\n[参照ファイル情報]\n"
                    for i, file in enumerate(retrieved_files[:5]):  # 最大5件まで表示
                        source_info += f"・{file['name']} (更新: {file['date']})\n"
                    
                    # 参照情報を回答の最後に追加
                    generated_text += source_info

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
    タイムアウトやエラー時に使用するフォールバック応答を返す

    Args:
        query: ユーザーの質問
        is_about_ollama: Ollamaに関する質問かどうか
        search_path: 検索したパス

    Returns:
        フォールバック応答
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
