# onedrive_search.py - OneDriveファイル検索機能（内容抽出機能追加版）
import os
import logging
import subprocess
import re
import time
import json
from datetime import datetime
import tempfile
import shutil
import sys

# 必要に応じてPyPDF2、python-docx、openpyxlをインポート
# これらを適切にインストールしておく必要があります
try:
    import PyPDF2
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False
    
try:
    import docx
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False
    
try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

try:
    from pptx import Presentation
    PPTX_SUPPORT = True
except ImportError:
    PPTX_SUPPORT = False

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

        # ライブラリサポート状況をログに記録
        logger.info(f"PDFサポート: {'有効' if PDF_SUPPORT else '無効 - PyPDF2をインストールしてください'}")
        logger.info(f"DOCXサポート: {'有効' if DOCX_SUPPORT else '無効 - python-docxをインストールしてください'}")
        logger.info(f"Excelサポート: {'有効' if EXCEL_SUPPORT else '無効 - openpyxlをインストールしてください'}")
        logger.info(f"PowerPointサポート: {'有効' if PPTX_SUPPORT else '無効 - python-pptxをインストールしてください'}")

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

    def read_file_content(self, file_path):
        """
        ファイルの内容を読み込む（改善版）

        Args:
            file_path: 読み込むファイルパス

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
                return self._read_text_file(file_path)
            else:
                # バイナリファイルの処理
                if ext == '.pdf':
                    return self._extract_pdf_content(file_path)
                elif ext == '.docx':
                    return self._extract_word_content(file_path)
                elif ext == '.xlsx':
                    return self._extract_excel_content(file_path)
                elif ext == '.pptx':
                    return self._extract_powerpoint_content(file_path)
                else:
                    # サポートされていないファイル形式
                    logger.warning(f"サポートされていないファイル形式です: {ext}")
                    return f"このファイル形式({ext})は直接読み込めません。"

        except PermissionError:
            logger.error(f"ファイル '{file_path}' へのアクセス権限がありません")
            return f"ファイル '{os.path.basename(file_path)}' へのアクセス権限がありません。システム管理者に確認してください。"
        except FileNotFoundError:
            logger.error(f"ファイル '{file_path}' が見つかりません")
            return f"ファイル '{os.path.basename(file_path)}' が見つかりません。削除または移動された可能性があります。"
        except Exception as e:
            logger.error(f"ファイル読み込み中にエラーが発生しました: {str(e)}")
            return f"ファイル読み込みエラー: {str(e)}"

    def _read_text_file(self, file_path):
        """
        テキストファイルの内容を読み込む

        Args:
            file_path: ファイルパス

        Returns:
            テキスト内容
        """
        encodings = ['utf-8', 'shift-jis', 'cp932', 'euc-jp', 'iso-2022-jp']

        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                    content = f.read()
                    logger.info(f"テキストファイルとして読み込みました ({encoding}): {file_path}")
                    return content
            except UnicodeDecodeError:
                continue

        # 全てのエンコーディングで失敗した場合
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
            logger.warning(f"代替エンコーディングで読み込みました: {file_path}")
            return content

    def _extract_pdf_content(self, file_path):
        """
        PDFファイルからテキストを抽出する（改善版）
        
        Args:
            file_path: PDFファイルのパス
            
        Returns:
            抽出したテキスト
        """
        # ファイル情報の基本ヘッダー
        file_info = f"PDF名: {os.path.basename(file_path)}\n"
        try:
            file_stat = os.stat(file_path)
            file_info += f"ファイルサイズ: {file_stat.st_size} bytes\n"
            file_info += f"最終更新日時: {datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}\n"
            file_info += "----------------------------------------\n"
        except:
            file_info += "ファイル情報の取得に失敗しました\n"
            file_info += "----------------------------------------\n"
        
        # PyPDF2が利用可能な場合、Pythonで抽出
        if PDF_SUPPORT:
            try:
                text = ""
                with open(file_path, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    num_pages = len(pdf_reader.pages)
                    
                    for page_num in range(num_pages):
                        page = pdf_reader.pages[page_num]
                        page_text = page.extract_text()
                        if page_text:
                            text += f"--- ページ {page_num + 1} ---\n{page_text}\n\n"
                
                if text:
                    return file_info + text
                else:
                    logger.warning("PDFからテキストを抽出できませんでした。PowerShellによる抽出を試みます。")
            except Exception as e:
                logger.error(f"PyPDF2でのPDF抽出エラー: {str(e)}")
                # エラーが発生した場合、PowerShellによる抽出を試みる
        
        # PowerShellを使用した抽出を試みる
        try:
            # 一時ファイルを作成してPDF内容を抽出する際の安全対策
            temp_output = os.path.join(tempfile.gettempdir(), f"pdf_extract_{os.path.basename(file_path)}.txt")
            
            ps_command = f"""
            $OutputEncoding = [System.Text.Encoding]::UTF8
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            
            try {{
                # Word.Application経由でPDFを開いてテキストに変換する方法
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                
                try {{
                    $document = $word.Documents.Open("{file_path}")
                    $text = $document.Content.Text
                    $document.Close()
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
                    
                    # 抽出したテキストを返す
                    $text
                }} catch {{
                    # Word経由での抽出に失敗した場合、基本情報と共にエラーメッセージを返す
                    "PDFの内容抽出中にエラーが発生しました: $_`nこのPDFファイルはテキスト抽出に対応していないか、保護されている可能性があります。"
                }}
            }} catch {{
                "PDFテキスト抽出エラー: $_"
            }}
            """
            
            # PowerShellコマンドを実行
            result = self._extract_file_content_helper(file_path, ps_command)
            
            # 結果が空または短すぎる場合
            if not result or len(result) < 50:
                return file_info + "このPDFからテキストを抽出できませんでした。スキャンされたPDFやテキストレイヤーのないPDFの可能性があります。"
            
            return file_info + result
            
        except Exception as e:
            logger.error(f"PDF抽出中にエラーが発生しました: {str(e)}")
            return file_info + "PDFコンテンツの抽出中にエラーが発生しました。"

    def _extract_word_content(self, file_path):
        """
        Wordファイル(docx)からテキストを抽出する（改善版）
        
        Args:
            file_path: Wordファイルのパス
            
        Returns:
            抽出したテキスト
        """
        # ファイル情報の基本ヘッダー
        file_info = f"Word文書名: {os.path.basename(file_path)}\n"
        try:
            file_stat = os.stat(file_path)
            file_info += f"ファイルサイズ: {file_stat.st_size} bytes\n"
            file_info += f"最終更新日時: {datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}\n"
            file_info += "----------------------------------------\n"
        except:
            file_info += "ファイル情報の取得に失敗しました\n"
            file_info += "----------------------------------------\n"
        
        # python-docxが利用可能な場合、Pythonで抽出
        if DOCX_SUPPORT:
            try:
                text = ""
                doc = docx.Document(file_path)
                
                # 段落のテキストを抽出
                for para in doc.paragraphs:
                    if para.text.strip():
                        text += para.text + "\n"
                
                # 表のテキストを抽出
                for table in doc.tables:
                    for row in table.rows:
                        row_text = []
                        for cell in row.cells:
                            row_text.append(cell.text.strip())
                        text += " | ".join(row_text) + "\n"
                    text += "\n"
                
                if text:
                    return file_info + text
                else:
                    logger.warning("Word文書からテキストを抽出できませんでした。PowerShellによる抽出を試みます。")
            except Exception as e:
                logger.error(f"python-docxでのWord抽出エラー: {str(e)}")
                # エラーが発生した場合、PowerShellによる抽出を試みる
        
        # PowerShellを使用した抽出を試みる
        try:
            ps_command = f"""
            $OutputEncoding = [System.Text.Encoding]::UTF8
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            
            try {{
                # Word.Application経由でdocxを開いてテキストに変換
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                
                try {{
                    $document = $word.Documents.Open("{file_path}")
                    $text = $document.Content.Text
                    $document.Close()
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
                    
                    # 抽出したテキストを返す
                    $text
                }} catch {{
                    # Word経由での抽出に失敗した場合、基本情報と共にエラーメッセージを返す
                    "Word文書の内容抽出中にエラーが発生しました: $_`nこのWord文書はテキスト抽出に対応していないか、保護されている可能性があります。"
                }}
            }} catch {{
                "Wordテキスト抽出エラー: $_"
            }}
            """
            
            # PowerShellコマンドを実行
            result = self._extract_file_content_helper(file_path, ps_command)
            
            # 結果が空または短すぎる場合
            if not result or len(result) < 50:
                return file_info + "このWord文書からテキストを抽出できませんでした。"
            
            return file_info + result
            
        except Exception as e:
            logger.error(f"Word抽出中にエラーが発生しました: {str(e)}")
            return file_info + "Word文書のコンテンツの抽出中にエラーが発生しました。"

    def _extract_excel_content(self, file_path):
        """
        Excelファイル(xlsx)から内容を抽出する（改善版）
        
        Args:
            file_path: Excelファイルのパス
            
        Returns:
            抽出したテキスト
        """
        # ファイル情報の基本ヘッダー
        file_info = f"Excel名: {os.path.basename(file_path)}\n"
        try:
            file_stat = os.stat(file_path)
            file_info += f"ファイルサイズ: {file_stat.st_size} bytes\n"
            file_info += f"最終更新日時: {datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}\n"
            file_info += "----------------------------------------\n"
        except:
            file_info += "ファイル情報の取得に失敗しました\n"
            file_info += "----------------------------------------\n"
        
        # PowerShellを使用した抽出を試みる
        try:
            ps_command = f"""
            $OutputEncoding = [System.Text.Encoding]::UTF8
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            
            try {{
                # Excel.Application経由でxlsxを開いてテキストに変換
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                
                try {{
                    $workbook = $excel.Workbooks.Open("{file_path}")
                    $text = ""
                    
                    # 各シートのデータを抽出
                    foreach ($sheet in $workbook.Sheets) {{
                        $text += "`n### シート: $($sheet.Name) ###`n"
                        
                        # 使用範囲のデータを取得
                        $usedRange = $sheet.UsedRange
                        $rows = $usedRange.Rows.Count
                        $cols = $usedRange.Columns.Count
                        
                        if ($rows -gt 0 -and $cols -gt 0) {{
                            for ($r = 1; $r -le [Math]::Min($rows, 100); $r++) {{
                                $rowText = ""
                                for ($c = 1; $c -le [Math]::Min($cols, 20); $c++) {{
                                    $cell = $usedRange.Cells.Item($r, $c).Text
                                    $rowText += "$cell | "
                                }}
                                $text += "$rowText`n"
                            }}
                            
                            if ($rows -gt 100 -or $cols -gt 20) {{
                                $text += "（大きなシートのため、一部のデータのみ表示しています）`n"
                            }}
                        }} else {{
                            $text += "（このシートにはデータがありません）`n"
                        }}
                    }}
                    
                    $workbook.Close($false)
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
                    
                    # 抽出したテキストを返す
                    $text
                }} catch {{
                    # Excel経由での抽出に失敗した場合、基本情報と共にエラーメッセージを返す
                    "Excelファイルの内容抽出中にエラーが発生しました: $_`nこのExcelファイルはテキスト抽出に対応していないか、保護されている可能性があります。"
                }}
            }} catch {{
                "Excelテキスト抽出エラー: $_"
            }}
            """
            
            # PowerShellコマンドを実行
            result = self._extract_file_content_helper(file_path, ps_command)
            
            # 結果が空または短すぎる場合
            if not result or len(result) < 50:
                return file_info + "このExcelファイルからデータを抽出できませんでした。"
            
            return file_info + result
            
        except Exception as e:
            logger.error(f"Excel抽出中にエラーが発生しました: {str(e)}")
            return file_info + "Excelファイルのデータ抽出中にエラーが発生しました。"