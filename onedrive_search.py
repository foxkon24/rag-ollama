# onedrive_search.py - OneDriveファイル検索機能（内容表示改善版）
import os
import logging
import subprocess
import re
import time
import json
from datetime import datetime
import base64
import tempfile

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

class OneDriveSearch:
    def __init__(self, base_directory=None, file_types=None, max_results=10):
        """
        OneDrive検索クラスの初期化

        Args:
            base_directory: 検索の基準ディレクトリ（指定がない場合はOneDriveルート）
            file_types: 検索対象のファイル拡張子リスト
            max_results: デフォルトの最大検索結果数
        """
        # OneDriveのルートディレクトリを取得（環境に応じて調整が必要）
        self.onedrive_root = os.path.expanduser("~/OneDrive")

        # ユーザー名を取得
        self.username = os.getenv("USERNAME", "owner")

        if not os.path.exists(self.onedrive_root):
            # 標準的なOneDriveパスが見つからない場合は代替パスを試す
            alt_paths = [
                os.path.expanduser("~/OneDrive - Company"),  # 企業アカウント用
                os.path.expanduser("~/OneDrive - Personal"),  # 個人アカウント用
                f"C:\\Users\\{self.username}\\OneDrive",  # 絶対パス
                f"D:\\OneDrive",  # 別ドライブ
                # 共立電機製作所のパターンを追加
                f"C:\\Users\\{self.username}\\OneDrive - 株式会社　共立電機製作所"
            ]

            for path in alt_paths:
                if os.path.exists(path):
                    self.onedrive_root = path
                    break

        logger.info(f"OneDriveルートディレクトリ: {self.onedrive_root}")

        # 検索の基準ディレクトリを設定
        self.base_directory = base_directory if base_directory else self.onedrive_root
        logger.info(f"検索基準ディレクトリ: {self.base_directory}")

        # ファイルタイプの設定
        self.file_types = file_types if file_types else []
        logger.info(f"検索対象ファイルタイプ: {', '.join(self.file_types) if self.file_types else '全ファイル'}")

        # 最大結果数
        self.max_results = max_results
        logger.info(f"デフォルト最大検索結果数: {self.max_results}")

        # 検索結果キャッシュ（パフォーマンス向上のため）
        self.search_cache = {}
        self.cache_expiry = 300  # キャッシュの有効期限（秒）

    def search_files(self, keywords, file_types=None, max_results=None, use_cache=True):
        """
        OneDrive内のファイルをキーワードで検索

        Args:
            keywords: 検索キーワード（文字列またはリスト）
            file_types: 検索対象の拡張子リスト（例: ['.pdf', '.docx']）
            max_results: 最大結果数
            use_cache: キャッシュを使用するかどうか

        Returns:
            検索結果のリスト [{'path': ファイルパス, 'name': ファイル名, 'modified': 更新日時}]
        """
        # デフォルト値の設定
        if file_types is None:
            file_types = self.file_types

        if max_results is None:
            max_results = self.max_results

        # キャッシュキーの生成
        cache_key = f"{str(keywords)}_{str(file_types)}_{max_results}"

        # キャッシュチェック
        if use_cache and cache_key in self.search_cache:
            cache_entry = self.search_cache[cache_key]
            cache_time = cache_entry['timestamp']
            current_time = time.time()

            # キャッシュが有効期限内なら使用
            if current_time - cache_time < self.cache_expiry:
                logger.info(f"キャッシュから検索結果を返します: {len(cache_entry['results'])}件")
                return cache_entry['results']

        # キーワードを文字列から配列に変換
        if isinstance(keywords, str):
            keywords = keywords.split()

        # 日報関連の特別キーワードを抽出
        date_keywords = []
        search_terms = []

        for k in keywords:
            # 日付形式（YYYY年MM月DD日）を抽出
            if re.search(r'\d{4}年\d{1,2}月\d{1,2}日', k):
                date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', k)
                if date_match:
                    year = date_match.group(1)
                    month = date_match.group(2).zfill(2)  # 1桁の月を2桁に
                    day = date_match.group(3).zfill(2)    # 1桁の日を2桁に
                    date_pattern = f"{year}{month}{day}"
                    date_pattern2 = f"{year}-{month}-{day}"
                    date_pattern3 = f"{year}/{month}/{day}"
                    date_keywords.extend([date_pattern, date_pattern2, date_pattern3])
            else:
                # 日本語検索キーワードは短くして検索精度を上げる
                if len(k) > 2 and re.search(r'[ぁ-んァ-ン一-龥]', k):
                    search_terms.append(k)

        # 少なくとも日付キーワードは追加
        if date_keywords:
            search_terms.extend(date_keywords)

        # 検索キーワードがない場合、元のキーワードの先頭2つを使用
        if not search_terms and keywords:
            search_terms = keywords[:2]

        # 最低1つのキーワードを確保
        if not search_terms and isinstance(keywords, str) and keywords:
            search_terms = [keywords]

        logger.info(f"OneDrive検索を実行: キーワード={search_terms}, ファイルタイプ={file_types}")

        try:
            # 検索パスのエスケープ
            search_path = self.base_directory.replace('\\', '\\\\')

            # ファイルタイプのフィルタ文字列を構築
            extension_filter = ""
            if file_types:
                ext_conditions = []
                for ext in file_types:
                    clean_ext = ext.replace('.', '')
                    ext_conditions.append(f"$_.Extension -eq '.{clean_ext}'")

                if ext_conditions:
                    extension_filter = f"({' -or '.join(ext_conditions)})"

            # 日付キーワードと通常キーワードで異なる検索戦略を使用
            date_conditions = []
            keyword_conditions = []

            # 日付パターンの検索条件
            for date_key in date_keywords:
                # ファイル名に日付が含まれるかより柔軟に検索
                date_conditions.append(f"$_.Name -like '*{date_key}*' -or $_.Name -match '{date_key}'")
                
                # 日付フォルダ構造にも対応（例：2023/10/26 や 2023-10-26 のようなフォルダ）
                if len(date_key) == 8 and date_key.isdigit():  # YYYYMMDD形式
                    year = date_key[:4]
                    month = date_key[4:6]
                    day = date_key[6:8]
                    date_folder_pattern = f"*\\\\{year}*\\\\{month}*\\\\{day}*"
                    date_conditions.append(f"$_.FullName -like '{date_folder_pattern}'")

            # 通常キーワードの検索条件
            for term in search_terms:
                if term not in date_keywords:  # 重複を避ける
                    keyword_conditions.append(f"$_.Name -like '*{term}*'")

            # 検索条件の組み立て
            search_condition = ""
            if date_conditions:
                date_part = f"({' -or '.join(date_conditions)})"
                if keyword_conditions:
                    keyword_part = f"({' -or '.join(keyword_conditions)})"
                    search_condition = f"{date_part} -and {keyword_part}"
                else:
                    search_condition = date_part
            elif keyword_conditions:
                search_condition = f"({' -or '.join(keyword_conditions)})"
            else:
                # 検索条件がない場合は全ファイルを対象に
                search_condition = "$true"

            # 複合条件を構築
            if extension_filter and search_condition:
                full_condition = f"{search_condition} -and {extension_filter}"
            elif extension_filter:
                full_condition = extension_filter
            else:
                full_condition = search_condition

            # PowerShellコマンドの構築（改善版）
            ps_command = f"""
            $OutputEncoding = [System.Text.Encoding]::UTF8
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            $ErrorActionPreference = "Continue"
            $results = @()
            try {{
                $files = Get-ChildItem -Path "{search_path}" -Recurse -File -ErrorAction SilentlyContinue
                $filtered = $files | Where-Object {{ {full_condition} }} | Select-Object -First {max_results}
                foreach ($file in $filtered) {{
                    $results += @{{
                        path = $file.FullName
                        name = $file.Name
                        modified = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                        size = $file.Length
                        extension = $file.Extension
                        created = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                    }}
                }}
                ConvertTo-Json -InputObject $results -Depth 2
            }} catch {{
                Write-Error "エラーが発生しました: $_"
                ConvertTo-Json -InputObject @() -Depth 2
            }}
            """

            # コマンド実行
            logger.debug(f"実行するPowerShellコマンド: {ps_command}")

            # PowerShellを起動（エンコーディングを明示的に指定）
            process = subprocess.Popen(
                ["powershell", "-ExecutionPolicy", "Bypass", "-NoProfile", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=False  # バイナリモードに変更
            )

            stdout, stderr = process.communicate()

            # バイナリ出力の適切な処理
            try:
                # 複数のエンコーディングでの解読を試みる
                stdout_text = None
                for encoding in ['utf-8', 'shift-jis', 'cp932']:
                    try:
                        stdout_text = stdout.decode(encoding, errors='replace')
                        if stdout_text and '[' in stdout_text:
                            break
                    except:
                        continue
                
                if not stdout_text:
                    stdout_text = stdout.decode('utf-8', errors='replace')
                
                # ステラーエラー出力の処理
                if stderr:
                    try:
                        stderr_text = stderr.decode('utf-8', errors='replace')
                        logger.error(f"PowerShell検索エラー: {stderr_text}")
                    except:
                        logger.error("PowerShell検索エラー: デコードできないエラーメッセージ")
            except Exception as e:
                logger.error(f"PowerShell出力のデコード中にエラー: {str(e)}")
                stdout_text = ""

            # 結果の解析
            results = []
            if stdout_text:
                try:
                    json_start = stdout_text.find('[')
                    if json_start >= 0:
                        json_data = stdout_text[json_start:]
                        parsed_data = json.loads(json_data)

                        # 結果が1件の場合、配列に変換
                        if isinstance(parsed_data, dict):
                            results = [parsed_data]
                        else:
                            results = parsed_data
                    else:
                        logger.warning(f"JSON形式の検索結果が見つかりませんでした: {stdout_text[:100]}...")

                except json.JSONDecodeError as e:
                    logger.error(f"JSON解析エラー: {e}")
                    logger.debug(f"解析できないJSON: {stdout_text[:200]}...")

            # 結果のフォーマットと表示
            logger.info(f"検索結果: {len(results)}件")
            for i, result in enumerate(results[:3]):  # 最初の3件のみログ表示
                logger.info(f"結果{i+1}: {result.get('name')} - {result.get('path')}")

            # キャッシュに保存
            self.search_cache[cache_key] = {
                'results': results,
                'timestamp': time.time()
            }

            return results

        except Exception as e:
            logger.error(f"OneDrive検索中にエラーが発生しました: {str(e)}")
            logger.error(f"詳細: {str(e.__class__.__name__)}")
            return []

    def read_file_content(self, file_path, max_chars=10000):
        """
        ファイルの内容を読み込む（改善版）

        Args:
            file_path: 読み込むファイルパス
            max_chars: 最大文字数

        Returns:
            ファイルの内容（文字列）
        """
        try:
            # ファイルの拡張子を取得
            _, ext = os.path.splitext(file_path.lower())

            # テキストファイル形式かどうか判定
            text_extensions = ['.txt', '.md', '.csv', '.json', '.xml', '.html', '.htm', '.log', '.py', '.js', '.css']
            is_text = ext in text_extensions

            if is_text:
                # テキストファイルとして読み込み
                encodings = ['utf-8', 'shift-jis', 'cp932', 'euc-jp', 'iso-2022-jp']

                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                            content = f.read(max_chars)
                            logger.info(f"テキストファイルとして読み込みました ({encoding}): {file_path}")
                            return self._format_text_content(content, file_path)
                    except UnicodeDecodeError:
                        continue

                # 全てのエンコーディングで失敗した場合
                with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read(max_chars)
                    logger.warning(f"代替エンコーディングで読み込みました: {file_path}")
                    return self._format_text_content(content, file_path)
            else:
                # バイナリファイルの場合はPowerShellを使用して内容を抽出
                if ext in ['.pdf']:
                    return self._extract_pdf_content(file_path)
                elif ext in ['.docx', '.doc']:
                    return self._extract_word_content(file_path)
                elif ext in ['.xlsx', '.xls']:
                    return self._extract_excel_content(file_path)
                elif ext in ['.pptx', '.ppt']:
                    return self._extract_powerpoint_content(file_path)
                else:
                    # サポートされていないファイル形式
                    logger.warning(f"サポートされていないファイル形式です: {ext}")
                    file_info = self._get_file_info(file_path)
                    return f"""【ファイル情報】
ファイル名: {os.path.basename(file_path)}
ファイル形式: {ext}
サイズ: {self._format_file_size(os.path.getsize(file_path))}
最終更新: {datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')}

※ このファイル形式({ext})の内容は直接表示できません。"""

        except PermissionError:
            logger.error(f"ファイル '{file_path}' へのアクセス権限がありません")
            return f"""【アクセスエラー】
ファイル '{os.path.basename(file_path)}' へのアクセス権限がありません。
システム管理者に確認してください。"""
        except FileNotFoundError:
            logger.error(f"ファイル '{file_path}' が見つかりません")
            return f"""【ファイルエラー】
ファイル '{os.path.basename(file_path)}' が見つかりません。
削除または移動された可能性があります。"""
        except Exception as e:
            logger.error(f"ファイル読み込み中にエラーが発生しました: {str(e)}")
            return f"""【読み込みエラー】
ファイル '{os.path.basename(file_path)}' の読み込み中にエラーが発生しました。
エラー詳細: {str(e)}"""

    def _format_text_content(self, content, file_path):
        """
        テキストコンテンツを整形する
        
        Args:
            content: 元のコンテンツ
            file_path: ファイルパス
            
        Returns:
            整形されたコンテンツ
        """
        file_name = os.path.basename(file_path)
        file_size = self._format_file_size(os.path.getsize(file_path))
        modified_time = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
        
        # ファイルの内容が長すぎる場合はトリミング
        if len(content) > 8000:
            content = content[:8000] + "\n...(以下省略)..."
            
        # CSVっぽいファイルの場合は表形式にフォーマット
        if file_path.lower().endswith('.csv') or (content.count(',') > 5 and content.count('\n') > 3):
            try:
                formatted_content = self._format_csv_content(content)
                if formatted_content:
                    content = formatted_content
            except:
                pass
                
        return f"""【テキストファイル情報】
ファイル名: {file_name}
サイズ: {file_size}
最終更新: {modified_time}

【ファイル内容】
{content}
"""

    def _format_csv_content(self, content, max_rows=30, max_cols=10):
        """
        CSV内容を見やすく整形する
        
        Args:
            content: CSVデータ
            max_rows: 表示する最大行数
            max_cols: 表示する最大列数
            
        Returns:
            整形されたCSV内容
        """
        lines = content.strip().split('\n')
        if not lines:
            return None
            
        rows = []
        for i, line in enumerate(lines[:max_rows+1]):
            cells = line.split(',')
            # 列数が多すぎる場合は切り詰める
            if len(cells) > max_cols:
                cells = cells[:max_cols-1] + ['...']
            rows.append(cells)
            
        if not rows:
            return None
            
        # 各列の最大幅を計算
        col_widths = []
        for col_idx in range(len(rows[0])):
            col_width = 0
            for row in rows:
                if col_idx < len(row):
                    col_width = max(col_width, len(row[col_idx]))
            col_widths.append(min(col_width, 30))  # 30文字で切り詰め
            
        # 整形された表を生成
        result = []
        for i, row in enumerate(rows):
            formatted_row = "| " + " | ".join(cell.ljust(col_widths[j]) for j, cell in enumerate(row)) + " |"
            result.append(formatted_row)
            
            # ヘッダー行の後に区切り線を挿入
            if i == 0:
                separator = "|-" + "-|-".join("-" * width for width in col_widths) + "-|"
                result.append(separator)
                
        # 行数が多すぎる場合の切り捨て表示
        if len(lines) > max_rows:
            result.append(f"\n... 他 {len(lines) - max_rows} 行 (表示行数制限)")
            
        return "\n".join(result)

    def _extract_file_content_helper(self, file_path, ps_command):
        """
        PowerShellを使用してファイル内容を抽出する共通ヘルパー（改善版）

        Args:
            file_path: ファイルパス
            ps_command: 実行するPowerShellコマンド

        Returns:
            抽出したテキスト
        """
        try:
            # PowerShellにエンコーディング設定を追加
            ps_command = f"$OutputEncoding = [System.Text.Encoding]::UTF8; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; {ps_command}"
            
            process = subprocess.Popen(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=False  # バイナリモード
            )

            stdout, stderr = process.communicate()

            # エンコーディング処理
            try:
                for encoding in ['utf-8', 'shift-jis', 'cp932']:
                    try:
                        result = stdout.decode(encoding, errors='replace')
                        if result:
                            return result
                    except:
                        continue
                
                # デフォルトフォールバック
                return stdout.decode('utf-8', errors='replace')
            except:
                return "ファイル内容のデコードに失敗しました。"

        except Exception as e:
            logger.error(f"ファイル処理中にエラーが発生しました: {str(e)}")
            return f"ファイル処理エラー: {str(e)}"

    def _get_file_info(self, file_path):
        """
        ファイルの基本情報を取得
        
        Args:
            file_path: ファイルパス
            
        Returns:
            ファイル情報を含む辞書
        """
        try:
            stats = os.stat(file_path)
            file_name = os.path.basename(file_path)
            file_ext = os.path.splitext(file_name)[1].lower()
            
            return {
                'name': file_name,
                'extension': file_ext,
                'size': stats.st_size,
                'size_formatted': self._format_file_size(stats.st_size),
                'created': datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                'modified': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                'path': file_path
            }
        except Exception as e:
            logger.error(f"ファイル情報取得エラー: {str(e)}")
            return {
                'name': os.path.basename(file_path),
                'error': str(e)
            }

    def _format_file_size(self, size_in_bytes):
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

    def _extract_pdf_content(self, file_path):
        """
        PDFファイルからテキストを抽出する（改善版）
        """
        file_info = self._get_file_info(file_path)
        file_name = file_info.get('name', os.path.basename(file_path))
        
        # PDFテキスト抽出が可能なPowerShellコマンド
        cmd = f"""
        try {{
            # ファイルの基本情報を取得
            Write-Output "【PDF文書情報】"
            Write-Output "ファイル名: {file_name}"
            Write-Output "サイズ: {file_info.get('size_formatted', 'Unknown')}"
            Write-Output "最終更新: {file_info.get('modified', 'Unknown')}"
            Write-Output ""
            Write-Output "【PDF内容】"
            
            # COM Objectを使ってPDFからテキスト抽出を試みる
            $extractedText = ""
            
            # PDFから最初の2ページのテキストを抽出する（Acrobat Reader DLLが必要）
            if ([System.IO.File]::Exists("{file_path}")) {{
                try {{
                    # Acrobatを利用可能か確認
                    $acrobat = New-Object -ComObject AcroExch.PDDoc -ErrorAction SilentlyContinue
                    if ($acrobat -ne $null) {{
                        # PDFファイルを開く
                        if ($acrobat.Open("{file_path}")) {{
                            $pageCount = $acrobat.GetNumPages()
                            Write-Output "合計ページ数: $pageCount"
                            Write-Output ""
                            
                            # 最初の3ページまでを処理
                            $maxPages = [Math]::Min(3, $pageCount)
                            for ($i = 0; $i -lt $maxPages; $i++) {{
                                try {{
                                    $page = $acrobat.AcquirePage($i)
                                    $pageText = $page.GetText()
                                    Write-Output "--- ページ $($i+1) ---"
                                    Write-Output $pageText
                                    Write-Output ""
                                }} catch {{
                                    Write-Output "ページ $($i+1) のテキスト抽出でエラーが発生しました。"
                                }}
                            }}
                            
                            if ($pageCount > 3) {{
                                Write-Output "... (残り $($pageCount - 3) ページは省略) ..."
                            }}
                            
                            # PDFを閉じる
                            $acrobat.Close()
                        }} else {{
                            Write-Output "PDFファイルを開けませんでした。"
                        }}
                    }} else {{
                        Write-Output "Acrobat Readerが利用できないため、PDF内容を直接抽出できません。"
                        Write-Output "PDF内容を確認するにはファイルを直接開いてください。"
                    }}
                }} catch {{
                    Write-Output "PDFテキスト抽出中にエラーが発生しました: $_"
                    Write-Output "PDF内容を確認するにはファイルを直接開いてください。"
                }}
            }} else {{
                Write-Output "PDFファイルが見つかりません: {file_path}"
            }}
        }} catch {{
            Write-Output "エラーが発生しました: $_"
        }}
        """
        
        content = self._extract_file_content_helper(file_path, cmd)
        
        # COM Objectでの抽出に失敗した場合の代替メッセージ
        if "PDF内容を確認するにはファイルを直接開いてください" in content or "Acrobat Readerが利用できない" in content:
            pdf_base_info = f"""【PDF文書情報】
ファイル名: {file_name}
サイズ: {file_info.get('size_formatted', 'Unknown')}
最終更新: {file_info.get('modified', 'Unknown')}

【PDF内容】
PDF内容を抽出するには専用ソフトウェアが必要です。
このPDFの内容を確認するには、直接PDFビューアーで開いてください。
"""
            return pdf_base_info
        
        return content

    def _extract_word_content(self, file_path):
        """
        Wordファイル(docx)からテキストを抽出する（改善版）
        """
        file_info = self._get_file_info(file_path)
        file_name = file_info.get('name', os.path.basename(file_path))
        
        cmd = f"""
        try {{
            # ファイルの基本情報を取得
            Write-Output "【Word文書情報】"
            Write-Output "ファイル名: {file_name}"
            Write-Output "サイズ: {file_info.get('size_formatted', 'Unknown')}"
            Write-Output "最終更新: {file_info.get('modified', 'Unknown')}"
            Write-Output ""
            Write-Output "【Word文書内容】"
            
            # Wordアプリケーションを使用してテキスト抽出を試みる
            if ([System.IO.File]::Exists("{file_path}")) {{
                try {{
                    # Wordを利用可能か確認
                    $word = New-Object -ComObject Word.Application -ErrorAction SilentlyContinue
                    if ($word -ne $null) {{
                        $word.Visible = $false
                        try {{
                            $doc = $word.Documents.Open("{file_path}")
                            Start-Sleep -Seconds 1
                            $content = $doc.Content.Text
                            
                            # ページ数と文字数
                            $stats = "文字数: " + $doc.Characters.Count
                            if ($doc.BuiltInDocumentProperties.Item("Number of Pages") -ne $null) {{
                                $stats += ", ページ数: " + $doc.BuiltInDocumentProperties.Item("Number of Pages").Value
                            }}
                            Write-Output $stats
                            Write-Output ""
                            
                            # コンテンツを出力（8000文字まで）
                            if ($content.Length -gt 8000) {{
                                Write-Output $content.Substring(0, 8000)
                                Write-Output "...(以下省略)..."
                            }} else {{
                                Write-Output $content
                            }}
                            
                            # 文書を閉じる
                            $doc.Close()
                        }} catch {{
                            Write-Output "文書を開く際にエラーが発生しました: $_"
                        }} finally {{
                            # Wordを終了
                            $word.Quit()
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                        }}
                    }} else {{
                        Write-Output "Microsoft Wordが利用できないため、内容を直接抽出できません。"
                        Write-Output "文書内容を確認するにはファイルを直接開いてください。"
                    }}
                }} catch {{
                    Write-Output "Word文書の処理中にエラーが発生しました: $_"
                }}
            }} else {{
                Write-Output "Word文書が見つかりません: {file_path}"
            }}
        }} catch {{
            Write-Output "エラーが発生しました: $_"
        }}
        """
        
        content = self._extract_file_content_helper(file_path, cmd)
        
        # COM Objectでの抽出に失敗した場合の代替メッセージ
        if "Microsoft Wordが利用できない" in content:
            word_base_info = f"""【Word文書情報】
ファイル名: {file_name}
サイズ: {file_info.get('size_formatted', 'Unknown')}
最終更新: {file_info.get('modified', 'Unknown')}

【Word文書内容】
Word文書内容を抽出するにはMicrosoft Wordが必要です。
この文書の内容を確認するには、Microsoft Wordで開いてください。
"""
            return word_base_info
        
        return content

    def _extract_excel_content(self, file_path):
        """
        Excelファイル(xlsx)からデータを抽出する（改善版）
        """
        file_info = self._get_file_info(file_path)
        file_name = file_info.get('name', os.path.basename(file_path))
        
        cmd = f"""
        try {{
            # ファイルの基本情報を取得
            Write-Output "【Excel情報】"
            Write-Output "ファイル名: {file_name}"
            Write-Output "サイズ: {file_info.get('size_formatted', 'Unknown')}"
            Write-Output "最終更新: {file_info.get('modified', 'Unknown')}"
            Write-Output ""
            
            # Excelアプリケーションを使用してデータ抽出を試みる
            if ([System.IO.File]::Exists("{file_path}")) {{
                try {{
                    # Excelを利用可能か確認
                    $excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
                    if ($excel -ne $null) {{
                        $excel.Visible = $false
                        $excel.DisplayAlerts = $false
                        try {{
                            $workbook = $excel.Workbooks.Open("{file_path}")
                            
                            # ワークシート情報
                            $sheetCount = $workbook.Sheets.Count
                            Write-Output "【Excel内容】"
                            Write-Output "シート数: $sheetCount" 
                            
                            # 各シートの処理（最初の3シートまで）
                            $maxSheets = [Math]::Min(3, $sheetCount)
                            for ($s = 1; $s -le $maxSheets; $s++) {{
                                $sheet = $workbook.Sheets.Item($s)
                                Write-Output ""
                                Write-Output "--- シート $s: $($sheet.Name) ---"
                                
                                # シートの使用範囲を取得
                                $usedRange = $sheet.UsedRange
                                $rowCount = $usedRange.Rows.Count
                                $colCount = $usedRange.Columns.Count
                                
                                Write-Output "データ範囲: $rowCount 行 x $colCount 列"
                                
                                # 大きすぎるシートの場合は処理を制限
                                $maxRows = [Math]::Min(20, $rowCount)
                                $maxCols = [Math]::Min(10, $colCount)
                                
                                # データ出力（表形式）
                                Write-Output ""
                                
                                # ヘッダー行を作成
                                $headerRow = "| "
                                for ($c = 1; $c -le $maxCols; $c++) {{
                                    $cellValue = $usedRange.Cells.Item(1, $c).Text
                                    $headerRow += "$cellValue | "
                                }}
                                if ($colCount > $maxCols) {{
                                    $headerRow += "... |"
                                }}
                                Write-Output $headerRow
                                
                                # 区切り線
                                $separator = "|"
                                for ($c = 1; $c -le $maxCols; $c++) {{
                                    $separator += " ----- |"
                                }}
                                if ($colCount > $maxCols) {{
                                    $separator += " ... |"
                                }}
                                Write-Output $separator
                                
                                # データ行を出力
                                for ($r = 2; $r -le $maxRows; $r++) {{
                                    $dataRow = "| "
                                    for ($c = 1; $c -le $maxCols; $c++) {{
                                        $cellValue = $usedRange.Cells.Item($r, $c).Text
                                        $dataRow += "$cellValue | "
                                    }}
                                    if ($colCount > $maxCols) {{
                                        $dataRow += "... |"
                                    }}
                                    Write-Output $dataRow
                                }}
                                
                                # 行数が多い場合
                                if ($rowCount > $maxRows) {{
                                    Write-Output "... (残り $($rowCount - $maxRows) 行は省略) ..."
                                }}
                            }}
                            
                            # 残りのシート情報
                            if ($sheetCount > $maxSheets) {{
                                Write-Output ""
                                Write-Output "... (残り $($sheetCount - $maxSheets) シートは省略) ..."
                                $remainingSheets = "残りのシート: "
                                for ($s = $maxSheets + 1; $s -le $sheetCount; $s++) {{
                                    $remainingSheets += "$($workbook.Sheets.Item($s).Name), "
                                }}
                                Write-Output $remainingSheets.TrimEnd(', ')
                            }}
                            
                            # ワークブックを閉じる
                            $workbook.Close($false)
                        }} catch {{
                            Write-Output "Excelブックを開く際にエラーが発生しました: $_"
                        }} finally {{
                            # Excelを終了
                            $excel.Quit()
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                        }}
                    }} else {{
                        Write-Output "【Excel内容】"
                        Write-Output "Microsoft Excelが利用できないため、内容を直接抽出できません。"
                        Write-Output "データを確認するにはファイルを直接開いてください。"
                    }}
                }} catch {{
                    Write-Output "Excelファイルの処理中にエラーが発生しました: $_"
                }}
            }} else {{
                Write-Output "Excelファイルが見つかりません: {file_path}"
            }}
        }} catch {{
            Write-Output "エラーが発生しました: $_"
        }}
        """
        
        content = self._extract_file_content_helper(file_path, cmd)
        
        # COM Objectでの抽出に失敗した場合の代替メッセージ
        if "Microsoft Excelが利用できない" in content:
            excel_base_info = f"""【Excel情報】
ファイル名: {file_name}
サイズ: {file_info.get('size_formatted', 'Unknown')}
最終更新: {file_info.get('modified', 'Unknown')}

【Excel内容】
Excelデータを抽出するにはMicrosoft Excelが必要です。
このブックの内容を確認するには、Microsoft Excelで開いてください。
"""
            return excel_base_info
        
        return content

    def _extract_powerpoint_content(self, file_path):
        """
        PowerPointファイル(pptx)から内容を抽出する（改善版）
        """
        file_info = self._get_file_info(file_path)
        file_name = file_info.get('name', os.path.basename(file_path))
        
        cmd = f"""
        try {{
            # ファイルの基本情報を取得
            Write-Output "【PowerPoint情報】"
            Write-Output "ファイル名: {file_name}"
            Write-Output "サイズ: {file_info.get('size_formatted', 'Unknown')}"
            Write-Output "最終更新: {file_info.get('modified', 'Unknown')}"
            Write-Output ""
            
            # PowerPointアプリケーションを使用してデータ抽出を試みる
            if ([System.IO.File]::Exists("{file_path}")) {{
                try {{
                    # PowerPointを利用可能か確認
                    $ppt = New-Object -ComObject PowerPoint.Application -ErrorAction SilentlyContinue
                    if ($ppt -ne $null) {{
                        try {{
                            $presentation = $ppt.Presentations.Open("{file_path}")
                            
                            # スライド情報
                            $slideCount = $presentation.Slides.Count
                            Write-Output "【PowerPoint内容】"
                            Write-Output "スライド数: $slideCount"
                            Write-Output ""
                            
                            # 各スライドのテキスト（最初の5スライドまで）
                            $maxSlides = [Math]::Min(5, $slideCount)
                            for ($i = 1; $i -le $maxSlides; $i++) {{
                                $slide = $presentation.Slides.Item($i)
                                Write-Output "--- スライド $i ---"
                                
                                # スライドのタイトルを取得
                                $title = ""
                                foreach ($shape in $slide.Shapes) {{
                                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {{
                                        if ($shape.Type -eq 14 -or $shape.Name -like "*Title*") {{
                                            $title = $shape.TextFrame.TextRange.Text
                                            break
                                        }}
                                    }}
                                }}
                                
                                if ($title -ne "") {{
                                    Write-Output "タイトル: $title"
                                }}
                                
                                # すべてのテキストを抽出
                                Write-Output "テキスト内容:"
                                foreach ($shape in $slide.Shapes) {{
                                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {{
                                        $text = $shape.TextFrame.TextRange.Text
                                        if ($text -ne $title -and $text.Trim() -ne "") {{
                                            Write-Output "・$text"
                                        }}
                                    }}
                                }}
                                Write-Output ""
                            }}
                            
                            # 残りのスライド情報
                            if ($slideCount > $maxSlides) {{
                                Write-Output "... (残り $($slideCount - $maxSlides) スライドは省略) ..."
                            }}
                            
                            # プレゼンテーションを閉じる
                            $presentation.Close()
                        }} catch {{
                            Write-Output "PowerPointファイルを開く際にエラーが発生しました: $_"
                        }} finally {{
                            # PowerPointを終了
                            $ppt.Quit()
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
                        }}
                    }} else {{
                        Write-Output "【PowerPoint内容】"
                        Write-Output "Microsoft PowerPointが利用できないため、内容を直接抽出できません。"
                        Write-Output "内容を確認するにはファイルを直接開いてください。"
                    }}
                }} catch {{
                    Write-Output "PowerPointファイルの処理中にエラーが発生しました: $_"
                }}
            }} else {{
                Write-Output "PowerPointファイルが見つかりません: {file_path}"
            }}
        }} catch {{
            Write-Output "エラーが発生しました: $_"
        }}
        """
        
        content = self._extract_file_content_helper(file_path, cmd)
        
        # COM Objectでの抽出に失敗した場合の代替メッセージ
        if "Microsoft PowerPointが利用できない" in content:
            ppt_base_info = f"""【PowerPoint情報】
ファイル名: {file_name}
サイズ: {file_info.get('size_formatted', 'Unknown')}
最終更新: {file_info.get('modified', 'Unknown')}

【PowerPoint内容】
PowerPointの内容を抽出するにはMicrosoft PowerPointが必要です。
このプレゼンテーションの内容を確認するには、Microsoft PowerPointで開いてください。
"""
            return ppt_base_info
        
        return content

    def get_relevant_content(self, query, max_files=None, max_chars=8000):
        """
        クエリに関連する内容を取得（改善版）

        Args:
            query: 検索クエリ
            max_files: 取得する最大ファイル数
            max_chars: 取得する最大文字数

        Returns:
            関連コンテンツ（文字列）
        """
        # 最大ファイル数の設定
        if max_files is None:
            max_files = self.max_results

        # 日付を抽出（YYYY年MM月DD日）
        date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', query)
        date_str = None

        if date_match:
            year = date_match.group(1)
            month = date_match.group(2).zfill(2)
            day = date_match.group(3).zfill(2)
            date_str = f"{year}年{month}月{day}日"
            date_pattern = f"{year}{month}{day}"
            logger.info(f"日付指定を検出: {date_str} (パターン: {date_pattern})")

        # 検索クエリからストップワードを除去
        stop_words = ["について", "とは", "の", "を", "に", "は", "で", "が", "と", "から", "へ", "より", 
                     "内容", "知りたい", "あったのか", "何", "教えて", "どのような", "どんな", "ありました",
                     "the", "a", "an", "and", "or", "of", "to", "in", "for", "on", "by"]

        # クエリから重要な単語を抽出
        keywords = []

        # 先に日付を追加（もし存在すれば）
        if date_str:
            keywords.append(date_str)

        # その他のキーワードを追加
        for word in query.split():
            clean_word = word.strip(',.;:!?()[]{}"\'')
            if clean_word and len(clean_word) > 1 and clean_word.lower() not in stop_words:
                # 日付文字列の一部でなければ追加
                if date_str and date_str not in clean_word:
                    keywords.append(clean_word)
                elif not date_str:
                    keywords.append(clean_word)
                else:
                    pass

        # キーワードが少なすぎる場合のバックアップとして日報関連の単語を追加
        if len(keywords) < 2:
            if "日報" not in query.lower() and not any(k for k in keywords if "日報" in k):
                keywords.append("日報")

        if not keywords:
            return "検索キーワードが見つかりませんでした。具体的な日付や単語で検索してください。"

        logger.info(f"抽出されたキーワード: {keywords}")

        # ファイル検索
        search_results = self.search_files(keywords, max_results=max_files)

        if not search_results:
            keywords_str = ", ".join(keywords)
            # 日付指定がある場合は特別なメッセージ
            if date_str:
                return f"{date_str}の日報は見つかりませんでした。日付の表記が正しいか確認してください。"
            else:
                return f"キーワード '{keywords_str}' に関連するファイルは見つかりませんでした。"

        # 関連コンテンツの取得（改善版）
        relevant_content = f"【検索結果】{len(search_results)}件の関連ファイルが見つかりました\n\n"
        total_chars = len(relevant_content)

        for i, result in enumerate(search_results):
            file_path = result.get('path')
            file_name = result.get('name')
            modified = result.get('modified', '不明')
            file_ext = result.get('extension', os.path.splitext(file_name)[1].lower())
            file_size = self._format_file_size(result.get('size', 0))

            # ファイルごとのヘッダー
            file_header = f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            file_header += f"【ファイル {i+1}】{file_name}\n"
            file_header += f"更新日時: {modified}  サイズ: {file_size}\n"
            file_header += f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"

            # 残りの文字数を計算
            remaining_chars = max_chars - total_chars - len(file_header)
            
            # 十分な文字数がある場合のみファイルコンテンツを読み込む
            if remaining_chars > 200:  # 最低限の表示のために200文字確保
                # ファイルの内容を読み込み（文字数制限あり）
                content = self.read_file_content(file_path, max_chars=min(remaining_chars, 5000))
                
                # 全体コンテンツに追加
                file_content = file_header + content + "\n\n"
                
                # 文字数制限チェック
                if total_chars + len(file_content) > max_chars:
                    # 制限を超える場合は切り詰め
                    trim_length = max_chars - total_chars - len(file_header) - 50
                    if trim_length > 200:
                        trimmed_content = content[:trim_length] + "...(以下省略)..."
                        file_content = file_header + trimmed_content + "\n\n"
                    else:
                        # もう表示できない場合は次のファイルに進まずに終了
                        relevant_content += f"\n（残り{len(search_results) - i}件のファイルは文字数制限のため表示されません）"
                        break
            else:
                # 残り文字数が少ない場合はファイル名だけ表示
                relevant_content += f"\n（残り{len(search_results) - i}件のファイルは文字数制限のため表示されません）"
                break
                
            relevant_content += file_content
            total_chars += len(file_content)

        return relevant_content
