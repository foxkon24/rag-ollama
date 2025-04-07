# onedrive_search.py - OneDriveファイル検索機能（改善版）
import os
import logging
import subprocess
import re
import time
import json
from datetime import datetime

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
                $filtered = $files | Where-Object {{ {full_condition} }} | Select-Object -First {max_results * 2}
                
                # ファイル名に直接キーワードが含まれるものを優先
                $nameMatches = $filtered | Where-Object {{ 
                    $file = $_
                    $containsAllKeywords = $true
                    foreach ($term in @({','.join([f"'{term}'" for term in search_terms]) if search_terms else "'dummy'"})) {{
                        if ($file.Name -notlike "*$term*") {{
                            $containsAllKeywords = $false
                            break
                        }}
                    }}
                    return $containsAllKeywords
                }} | Select-Object -First {max_results}
                
                # 残りの候補から追加
                $remainingSlots = {max_results} - $nameMatches.Count
                $otherMatches = @()
                if ($remainingSlots -gt 0) {{
                    $otherMatches = $filtered | Where-Object {{ 
                        $file = $_
                        $found = $false
                        foreach ($nm in $nameMatches) {{
                            if ($file.FullName -eq $nm.FullName) {{
                                $found = $true
                                break
                            }}
                        }}
                        return !$found
                    }} | Select-Object -First $remainingSlots
                }}
                
                # 結果を結合
                $finalResults = $nameMatches + $otherMatches
                
                foreach ($file in $finalResults) {{
                    $results += @{{
                        path = $file.FullName
                        name = $file.Name
                        modified = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                        size = $file.Length
                        created = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                        extension = $file.Extension
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
        ファイルの内容を読み込む

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
            else:
                # バイナリファイルの場合はPowerShellを使用して内容を抽出
                if ext in ['.pdf']:
                    return self._extract_pdf_content(file_path)
                elif ext in ['.docx']:
                    return self._extract_word_content(file_path)
                elif ext in ['.xlsx']:
                    return self._extract_excel_content(file_path)
                elif ext in ['.pptx']:
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

    def _extract_file_content_helper(self, file_path, ps_command):
        """
        PowerShellを使用してファイル内容を抽出する共通ヘルパー

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

    def _extract_pdf_content(self, file_path):
        """
        PDFファイルからテキストを抽出する（改善版）
        """
        cmd = f"""
        try {{
            # ファイルの存在確認
            if (Test-Path -Path "{file_path}" -PathType Leaf) {{
                $pdfInfo = "PDF名: {os.path.basename(file_path)}"
                $pdfInfo += "`nファイルサイズ: " + (Get-Item "{file_path}").Length + " bytes"
                $pdfInfo += "`n最終更新日時: " + (Get-Item "{file_path}").LastWriteTime
                $pdfInfo += "`n----------------------------------------"
                
                # PDF内容の抽出を試みる（サードパーティ製ツールがなくてもメタデータは取得）
                try {{
                    # PDFメタデータの取得を試みる
                    Add-Type -AssemblyName System.IO.Compression.FileSystem
                    $pdfInfo += "`n[PDF内容要約]`n"
                    
                    # バイナリデータから簡易メタデータ抽出（テキスト部分のみ）
                    $pdfBytes = [System.IO.File]::ReadAllBytes("{file_path}")
                    $pdfText = [System.Text.Encoding]::ASCII.GetString($pdfBytes)
                    
                    # タイトルの抽出試行
                    $titleMatch = [regex]::Match($pdfText, "/Title\\s*\\(([^)]+)\\)")
                    if ($titleMatch.Success) {{
                        $pdfInfo += "`nタイトル: " + $titleMatch.Groups[1].Value
                    }}
                    
                    # 作成者の抽出試行
                    $authorMatch = [regex]::Match($pdfText, "/Author\\s*\\(([^)]+)\\)")
                    if ($authorMatch.Success) {{
                        $pdfInfo += "`n作成者: " + $authorMatch.Groups[1].Value
                    }}
                    
                    # キーワードの抽出試行
                    $keywordsMatch = [regex]::Match($pdfText, "/Keywords\\s*\\(([^)]+)\\)")
                    if ($keywordsMatch.Success) {{
                        $pdfInfo += "`nキーワード: " + $keywordsMatch.Groups[1].Value
                    }}
                    
                    # 簡易テキスト抽出（完全ではない）
                    $textMatches = [regex]::Matches($pdfText, "\\(([\\w\\s,.;:!?-]{{10,}})\\)")
                    if ($textMatches.Count -gt 0) {{
                        $pdfInfo += "`n`n[PDFから抽出されたテキスト（部分的）]`n"
                        $extractedText = @()
                        foreach ($match in $textMatches) {{
                            $text = $match.Groups[1].Value
                            # 意味のありそうなテキストのみを抽出（短すぎるものや記号のみは除外）
                            if ($text.Length -gt 20 -and $text -match "[a-zA-Z0-9ぁ-んァ-ン一-龥]") {{
                                $extractedText += $text
                            }}
                        }}
                        
                        # 重複を避け、最も長いテキスト部分を最大5つまで表示
                        $uniqueTexts = $extractedText | Sort-Object -Property Length -Descending | Select-Object -First 5
                        $pdfInfo += $uniqueTexts -join "`n`n"
                    }} else {{
                        $pdfInfo += "`n`nPDFからテキストを抽出できませんでした。"
                    }}
                    
                }} catch {{
                    $pdfInfo += "`nPDFの内容抽出に失敗しました: $_"
                }}
                
                $pdfInfo
            }} else {{
                "ファイルが見つかりません: {file_path}"
            }}
        }} catch {{
            "エラーが発生しました: $_"
        }}
        """
        return self._extract_file_content_helper(file_path, cmd)

    def _extract_word_content(self, file_path):
        """
        Wordファイル(docx)からテキストを抽出する（改善版）
        """
        cmd = f"""
        try {{
            # ファイルの存在確認
            if (Test-Path -Path "{file_path}" -PathType Leaf) {{
                $wordInfo = "Word文書名: {os.path.basename(file_path)}`n"
                $wordInfo += "ファイルサイズ: " + (Get-Item "{file_path}").Length + " bytes`n"
                $wordInfo += "最終更新日時: " + (Get-Item "{file_path}").LastWriteTime + "`n"
                $wordInfo += "----------------------------------------`n"
                
                # DOCXファイルはZIPファイルなので展開して内容を抽出
                try {{
                    Add-Type -Assembly System.IO.Compression.FileSystem
                    $tempFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString())
                    [System.IO.Directory]::CreateDirectory($tempFolder)
                    
                    # ZIPとして展開
                    [System.IO.Compression.ZipFile]::ExtractToDirectory("{file_path}", $tempFolder)
                    
                    # document.xmlからテキスト抽出
                    $docXmlPath = [System.IO.Path]::Combine($tempFolder, "word", "document.xml")
                    if (Test-Path $docXmlPath) {{
                        $docXml = [System.Xml.XmlDocument]::new()
                        $docXml.Load($docXmlPath)
                        $nsManager = [System.Xml.XmlNamespaceManager]::new($docXml.NameTable)
                        $nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                        
                        # テキスト要素を抽出
                        $textNodes = $docXml.SelectNodes("//w:t", $nsManager)
                        $text = ""
                        foreach ($node in $textNodes) {{
                            $text += $node.InnerText + " "
                        }}
                        
                        $wordInfo += "`n[文書内容]`n" + $text
                    }} else {{
                        $wordInfo += "`n文書の内容を抽出できませんでした。"
                    }}
                    
                    # プロパティの抽出
                    $corePropsPath = [System.IO.Path]::Combine($tempFolder, "docProps", "core.xml")
                    if (Test-Path $corePropsPath) {{
                        $coreXml = [System.Xml.XmlDocument]::new()
                        $coreXml.Load($corePropsPath)
                        $nsManager = [System.Xml.XmlNamespaceManager]::new($coreXml.NameTable)
                        $nsManager.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")
                        $nsManager.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
                        
                        $title = $coreXml.SelectSingleNode("//dc:title", $nsManager)
                        $creator = $coreXml.SelectSingleNode("//dc:creator", $nsManager)
                        
                        $wordInfo += "`n`n[文書プロパティ]`n"
                        if ($title -and $title.InnerText) {{
                            $wordInfo += "タイトル: " + $title.InnerText + "`n"
                        }}
                        if ($creator -and $creator.InnerText) {{
                            $wordInfo += "作成者: " + $creator.InnerText + "`n"
                        }}
                    }}
                    
                    # 一時フォルダの削除
                    Remove-Item -Path $tempFolder -Recurse -Force
                }} catch {{
                    $wordInfo += "Word文書の内容抽出中にエラーが発生しました: $_"
                }}
                
                $wordInfo
            }} else {{
                "ファイルが見つかりません: {file_path}"
            }}
        }} catch {{
            "エラーが発生しました: $_"
        }}
        """
        return self._extract_file_content_helper(file_path, cmd)

    def _extract_excel_content(self, file_path):
        """
        Excelファイル(xlsx)から基本情報を抽出する（改善版）
        """
        cmd = f"""
        try {{
            # ファイルの存在確認
            if (Test-Path -Path "{file_path}" -PathType Leaf) {{
                $excelInfo = "Excel名: {os.path.basename(file_path)}`n"
                $excelInfo += "ファイルサイズ: " + (Get-Item "{file_path}").Length + " bytes`n"
                $excelInfo += "最終更新日時: " + (Get-Item "{file_path}").LastWriteTime + "`n"
                $excelInfo += "----------------------------------------`n"
                
                # XLSXファイルはZIPファイルなので展開して内容を抽出
                try {{
                    Add-Type -Assembly System.IO.Compression.FileSystem
                    $tempFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString())
                    [System.IO.Directory]::CreateDirectory($tempFolder)
                    
                    # ZIPとして展開
                    [System.IO.Compression.ZipFile]::ExtractToDirectory("{file_path}", $tempFolder)
                    
                    # シート名の抽出
                    $workbookPath = [System.IO.Path]::Combine($tempFolder, "xl", "workbook.xml")
                    if (Test-Path $workbookPath) {{
                        $workbookXml = [System.Xml.XmlDocument]::new()
                        $workbookXml.Load($workbookPath)
                        
                        $nsManager = [System.Xml.XmlNamespaceManager]::new($workbookXml.NameTable)
                        $nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                        
                        $sheetNodes = $workbookXml.SelectNodes("//x:sheet", $nsManager)
                        
                        $excelInfo += "`n[ブック構成]`n"
                        if ($sheetNodes.Count -gt 0) {{
                            $excelInfo += "シート数: " + $sheetNodes.Count + "`n"
                            $excelInfo += "シート名: "
                            $sheetNames = @()
                            foreach ($sheet in $sheetNodes) {{
                                $sheetNames += $sheet.GetAttribute("name")
                            }}
                            $excelInfo += $sheetNames -join ", "
                            $excelInfo += "`n"
                            
                            # 先頭シートの内容を抽出（簡易）
                            $firstSheetId = $sheetNodes[0].GetAttribute("sheetId")
                            $sheetPath = [System.IO.Path]::Combine($tempFolder, "xl", "worksheets", "sheet1.xml")
                            
                            if (Test-Path $sheetPath) {{
                                $sheetXml = [System.Xml.XmlDocument]::new()
                                $sheetXml.Load($sheetPath)
                                
                                $nsManager = [System.Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
                                $nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                                
                                # セル数をカウント
                                $cells = $sheetXml.SelectNodes("//x:c", $nsManager)
                                
                                $excelInfo += "`n[最初のシート情報]`n"
                                $excelInfo += "シート名: " + $sheetNames[0] + "`n"
                                $excelInfo += "セル数: " + $cells.Count + "`n"
                                
                                # データ内の文字列を簡易的に取得（最大10要素）
                                $stringTablePath = [System.IO.Path]::Combine($tempFolder, "xl", "sharedStrings.xml")
                                if (Test-Path $stringTablePath) {{
                                    $stringXml = [System.Xml.XmlDocument]::new()
                                    $stringXml.Load($stringTablePath)
                                    
                                    $nsManager = [System.Xml.XmlNamespaceManager]::new($stringXml.NameTable)
                                    $nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                                    
                                    $strings = $stringXml.SelectNodes("//x:t", $nsManager) | Select-Object -First 20
                                    
                                    if ($strings.Count -gt 0) {{
                                        $excelInfo += "`n[データサンプル]`n"
                                        $dataItems = @()
                                        foreach ($string in $strings) {{
                                            if ($string.InnerText.Trim() -ne "") {{
                                                $dataItems += $string.InnerText
                                            }}
                                        }}
                                        $excelInfo += $dataItems -join ", "
                                    }}
                                }}
                            }}
                        }} else {{
                            $excelInfo += "シート情報を取得できませんでした。"
                        }}
                    }} else {{
                        $excelInfo += "`nブック構成を取得できませんでした。"
                    }}
                    
                    # プロパティの抽出
                    $corePropsPath = [System.IO.Path]::Combine($tempFolder, "docProps", "core.xml")
                    if (Test-Path $corePropsPath) {{
                        $coreXml = [System.Xml.XmlDocument]::new()
                        $coreXml.Load($corePropsPath)
                        
                        $nsManager = [System.Xml.XmlNamespaceManager]::new($coreXml.NameTable)
                        $nsManager.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")
                        $nsManager.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
                        
                        $title = $coreXml.SelectSingleNode("//dc:title", $nsManager)
                        $creator = $coreXml.SelectSingleNode("//dc:creator", $nsManager)
                        
                        if (($title -and $title.InnerText) -or ($creator -and $creator.InnerText)) {{
                            $excelInfo += "`n`n[ブックプロパティ]`n"
                            if ($title -and $title.InnerText) {{
                                $excelInfo += "タイトル: " + $title.InnerText + "`n"
                            }}
                            if ($creator -and $creator.InnerText) {{
                                $excelInfo += "作成者: " + $creator.InnerText + "`n"
                            }}
                        }}
                    }}
                    
                    # 一時フォルダの削除
                    Remove-Item -Path $tempFolder -Recurse -Force
                }} catch {{
                    $excelInfo += "`nExcel内容の抽出中にエラー: $_"
                }}
                
                $excelInfo
            }} else {{
                "ファイルが見つかりません: {file_path}"
            }}
        }} catch {{
            "エラーが発生しました: $_"
        }}
        """
        return self._extract_file_content_helper(file_path, cmd)

    def _extract_powerpoint_content(self, file_path):
        """
        PowerPointファイル(pptx)から基本情報を抽出する（改善版）
        """
        cmd = f"""
        try {{
            # ファイルの存在確認
            if (Test-Path -Path "{file_path}" -PathType Leaf) {{
                $pptInfo = "PowerPoint名: {os.path.basename(file_path)}`n"
                $pptInfo += "ファイルサイズ: " + (Get-Item "{file_path}").Length + " bytes`n"
                $pptInfo += "最終更新日時: " + (Get-Item "{file_path}").LastWriteTime + "`n"
                $pptInfo += "----------------------------------------`n"
                
                # PPTXファイルはZIPファイルなので展開して内容を抽出
                try {{
                    Add-Type -Assembly System.IO.Compression.FileSystem
                    $tempFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString())
                    [System.IO.Directory]::CreateDirectory($tempFolder)
                    
                    # ZIPとして展開
                    [System.IO.Compression.ZipFile]::ExtractToDirectory("{file_path}", $tempFolder)
                    
                    # スライド数をカウント
                    $slidesFolder = [System.IO.Path]::Combine($tempFolder, "ppt", "slides")
                    if (Test-Path $slidesFolder) {{
                        $slides = Get-ChildItem -Path $slidesFolder -Filter "slide*.xml"
                        $slideCount = $slides.Count
                        
                        $pptInfo += "`n[プレゼンテーション構成]`n"
                        $pptInfo += "スライド数: $slideCount`n"
                        
                        # スライドのテキスト抽出（最初の数枚）
                        $pptInfo += "`n[スライド内容サンプル]`n"
                        
                        $slideTexts = @()
                        foreach ($slide in $slides | Select-Object -First 5) {{
                            $slideXml = [System.Xml.XmlDocument]::new()
                            $slideXml.Load($slide.FullName)
                            
                            $nsManager = [System.Xml.XmlNamespaceManager]::new($slideXml.NameTable)
                            $nsManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                            
                            $textNodes = $slideXml.SelectNodes("//a:t", $nsManager)
                            $slideText = "スライド " + $slide.Name.Replace("slide", "").Replace(".xml", "") + ": "
                            
                            $textItems = @()
                            foreach ($node in $textNodes) {{
                                if ($node.InnerText.Trim() -ne "") {{
                                    $textItems += $node.InnerText
                                }}
                            }}
                            
                            if ($textItems.Count -gt 0) {{
                                $slideText += $textItems -join " "
                                $slideTexts += $slideText
                            }}
                        }}
                        
                        $pptInfo += $slideTexts -join "`n"
                    }} else {{
                        $pptInfo += "`nスライド情報を取得できませんでした。"
                    }}
                    
                    # プロパティの抽出
                    $corePropsPath = [System.IO.Path]::Combine($tempFolder, "docProps", "core.xml")
                    if (Test-Path $corePropsPath) {{
                        $coreXml = [System.Xml.XmlDocument]::new()
                        $coreXml.Load($corePropsPath)
                        
                        $nsManager = [System.Xml.XmlNamespaceManager]::new($coreXml.NameTable)
                        $nsManager.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")
                        $nsManager.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
                        
                        $title = $coreXml.SelectSingleNode("//dc:title", $nsManager)
                        $creator = $coreXml.SelectSingleNode("//dc:creator", $nsManager)
                        
                        if (($title -and $title.InnerText) -or ($creator -and $creator.InnerText)) {{
                            $pptInfo += "`n`n[プレゼンテーションプロパティ]`n"
                            if ($title -and $title.InnerText) {{
                                $pptInfo += "タイトル: " + $title.InnerText + "`n"
                            }}
                            if ($creator -and $creator.InnerText) {{
                                $pptInfo += "作成者: " + $creator.InnerText + "`n"
                            }}
                        }}
                    }}
                    
                    # 一時フォルダの削除
                    Remove-Item -Path $tempFolder -Recurse -Force
                }} catch {{
                    $pptInfo += "`nPowerPoint内容の抽出中にエラー: $_"
                }}
                
                $pptInfo
            }} else {{
                "ファイルが見つかりません: {file_path}"
            }}
        }} catch {{
            "エラーが発生しました: $_"
        }}
        """
        return self._extract_file_content_helper(file_path, cmd)

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

        # 関連コンテンツの取得
        relevant_content = f"--- {len(search_results)}件の関連ファイルが見つかりました ---\n\n"
        total_chars = len(relevant_content)

        for i, result in enumerate(search_results):
            file_path = result.get('path')
            file_name = result.get('name')
            modified = result.get('modified', '不明')
            file_size = result.get('size', 0)
            
            # ファイルサイズを見やすい形式に変換
            if file_size:
                if file_size < 1024:
                    size_str = f"{file_size} バイト"
                elif file_size < 1024 * 1024:
                    size_str = f"{file_size / 1024:.1f} KB"
                else:
                    size_str = f"{file_size / (1024 * 1024):.1f} MB"
            else:
                size_str = "不明"

            # ファイルの内容を読み込み
            content = self.read_file_content(file_path)

            # コンテンツのプレビューを追加（文字数制限あり）
            preview_length = min(2000, len(content))  # 1ファイルあたり最大2000文字
            preview = content[:preview_length]
            if len(content) > preview_length:
                preview += "..."

            file_content = f"=== ファイル {i+1}: {file_name} ===\n"
            file_content += f"更新日時: {modified}\n"
            file_content += f"サイズ: {size_str}\n"
            file_content += f"パス: {file_path}\n\n"
            file_content += f"{preview}\n\n"

            # 最大文字数をチェック
            if total_chars + len(file_content) > max_chars:
                # 制限に達した場合は切り詰め
                remaining = max_chars - total_chars - 100  # 終了メッセージ用に余裕を持たせる
                if remaining > 0:
                    file_content = file_content[:remaining] + "...\n"
                else:
                    # もう追加できない場合
                    relevant_content += f"\n（残り{len(search_results) - i}件のファイルは文字数制限のため表示されません）"
                    break

            relevant_content += file_content
            total_chars += len(file_content)

        return relevant_content

# 使用例
if __name__ == "__main__":
    # ロギングの設定
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # インスタンス作成
    onedrive_search = OneDriveSearch()

    # 検索クエリ
    test_query = "2024年10月26日の日報内容"

    # 関連コンテンツを取得
    content = onedrive_search.get_relevant_content(test_query)

    print(f"検索結果: {content[:500]}...")  # 最初の500文字のみ表示
