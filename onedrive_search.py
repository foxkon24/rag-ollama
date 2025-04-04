logger.warning("Excelファイルからデータを抽出できませんでした。PowerShellによる抽出を試みます。")
            except Exception as e:
                logger.error(f"openpyxlでのExcel抽出エラー: {str(e)}")
                # エラーが発生した場合、PowerShellによる抽出を試みる
        
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

    def _extract_powerpoint_content(self, file_path):
        """
        PowerPointファイル(pptx)からテキストを抽出する（改善版）
        
        Args:
            file_path: PowerPointファイルのパス
            
        Returns:
            抽出したテキスト
        """
        # ファイル情報の基本ヘッダー
        file_info = f"PowerPoint名: {os.path.basename(file_path)}\n"
        try:
            file_stat = os.stat(file_path)
            file_info += f"ファイルサイズ: {file_stat.st_size} bytes\n"
            file_info += f"最終更新日時: {datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}\n"
            file_info += "----------------------------------------\n"
        except:
            file_info += "ファイル情報の取得に失敗しました\n"
            file_info += "----------------------------------------\n"
        
        # python-pptxが利用可能な場合、Pythonで抽出
        if PPTX_SUPPORT:
            try:
                text = ""
                prs = Presentation(file_path)
                
                # 各スライドのテキストを抽出
                for i, slide in enumerate(prs.slides):
                    text += f"\n### スライド {i+1} ###\n"
                    
                    # スライド内のテキスト形状を抽出
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text:
                            text += f"{shape.text.strip()}\n"
                
                if text:
                    return file_info + text
                else:
                    logger.warning("PowerPointからテキストを抽出できませんでした。PowerShellによる抽出を試みます。")
            except Exception as e:
                logger.error(f"python-pptxでのPowerPoint抽出エラー: {str(e)}")
                # エラーが発生した場合、PowerShellによる抽出を試みる
        
        # PowerShellを使用した抽出を試みる
        try:
            ps_command = f"""
            $OutputEncoding = [System.Text.Encoding]::UTF8
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            
            try {{
                # PowerPoint.Application経由でpptxを開いてテキストに変換
                $ppt = New-Object -ComObject PowerPoint.Application
                $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
                
                try {{
                    $presentation = $ppt.Presentations.Open("{file_path}", $false, $false, $false)
                    $text = ""
                    
                    # 各スライドのテキストを抽出
                    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {{
                        $slide = $presentation.Slides.Item($i)
                        $text += "`n### スライド $i ###`n"
                        
                        # 各シェイプからテキストを抽出
                        foreach ($shape in $slide.Shapes) {{
                            if ($shape.HasTextFrame) {{
                                $textFrame = $shape.TextFrame
                                if ($textFrame.HasText) {{
                                    $text += "$($textFrame.TextRange.Text)`n"
                                }}
                            }}
                        }}
                    }}
                    
                    $presentation.Close()
                    $ppt.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)
                    
                    # 抽出したテキストを返す
                    $text
                }} catch {{
                    # PowerPoint経由での抽出に失敗した場合、基本情報と共にエラーメッセージを返す
                    "PowerPointの内容抽出中にエラーが発生しました: $_`nこのPowerPointはテキスト抽出に対応していないか、保護されている可能性があります。"
                }}
            }} catch {{
                "PowerPointテキスト抽出エラー: $_"
            }}
            """
            
            # PowerShellコマンドを実行
            result = self._extract_file_content_helper(file_path, ps_command)
            
            # 結果が空または短すぎる場合
            if not result or len(result) < 50:
                return file_info + "このPowerPointからテキストを抽出できませんでした。"
            
            return file_info + result
            
        except Exception as e:
            logger.error(f"PowerPoint抽出中にエラーが発生しました: {str(e)}")
            return file_info + "PowerPointのテキスト抽出中にエラーが発生しました。"

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
            # PowerShellにエンコーディング設定を追加（コマンドの先頭に既に設定があるためスキップ）
            
            # プロセスのタイムアウト設定
            timeout = 60  # 60秒のタイムアウト
            
            # PowerShellプロセスを起動
            process = subprocess.Popen(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=False  # バイナリモード
            )

            try:
                # タイムアウトを設定して実行
                stdout, stderr = process.communicate(timeout=timeout)
                
                # エンコーディング処理
                for encoding in ['utf-8', 'shift-jis', 'cp932']:
                    try:
                        result = stdout.decode(encoding, errors='replace')
                        if result:
                            return result
                    except:
                        continue
                
                # デフォルトフォールバック
                return stdout.decode('utf-8', errors='replace')
                
            except subprocess.TimeoutExpired:
                # タイムアウトした場合はプロセスを強制終了
                process.kill()
                logger.error(f"PowerShellコマンドの実行がタイムアウトしました: {file_path}")
                return "ファイル処理がタイムアウトしました。ファイルが大きすぎるか、処理が複雑すぎる可能性があります。"
                
        except Exception as e:
            logger.error(f"ファイル処理中にエラーが発生しました: {str(e)}")
            return f"ファイル処理エラー: {str(e)}"
            
    def extract_text_with_python_tools(self, file_path):
        """
        適切なPythonライブラリを使用してファイルからテキストを抽出する
        
        Args:
            file_path: ファイルパス
            
        Returns:
            抽出したテキスト、または空文字列（抽出できない場合）
        """
        _, ext = os.path.splitext(file_path.lower())
        
        try:
            # ファイルタイプに基づいて適切な抽出メソッドを選択
            if ext == '.pdf' and PDF_SUPPORT:
                return self._extract_pdf_with_pypdf(file_path)
            elif ext == '.docx' and DOCX_SUPPORT:
                return self._extract_docx_with_python_docx(file_path)
            elif ext == '.xlsx' and EXCEL_SUPPORT:
                return self._extract_xlsx_with_openpyxl(file_path)
            elif ext == '.pptx' and PPTX_SUPPORT:
                return self._extract_pptx_with_python_pptx(file_path)
            else:
                return ""
        except Exception as e:
            logger.error(f"Pythonツールでのテキスト抽出エラー: {str(e)}")
            return ""
            
    def _extract_pdf_with_pypdf(self, file_path):
        """PyPDF2を使用してPDFからテキストを抽出"""
        text = ""
        with open(file_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                if page_text:
                    text += f"--- ページ {page_num + 1} ---\n{page_text}\n\n"
        return text
    
    def _extract_docx_with_python_docx(self, file_path):
        """python-docxを使用してWordドキュメントからテキストを抽出"""
        text = ""
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            if para.text.strip():
                text += para.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    row_text.append(cell.text.strip())
                text += " | ".join(row_text) + "\n"
            text += "\n"
        return text
    
    def _extract_xlsx_with_openpyxl(self, file_path):
        """openpyxlを使用してExcelファイルからデータを抽出"""
        text = ""
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            text += f"\n### シート: {sheet_name} ###\n"
            data_rows = []
            for row in sheet.iter_rows():
                row_data = []
                for cell in row:
                    if cell.value is not None:
                        row_data.append(str(cell.value))
                    else:
                        row_data.append("")
                if any(row_data):
                    data_rows.append(row_data)
            if data_rows:
                for row_data in data_rows:
                    text += " | ".join(row_data) + "\n"
            else:
                text += "（このシートにはデータがありません）\n"
        return text
    
    def _extract_pptx_with_python_pptx(self, file_path):
        """python-pptxを使用してPowerPointファイルからテキストを抽出"""
        text = ""
        prs = Presentation(file_path)
        for i, slide in enumerate(prs.slides):
            text += f"\n### スライド {i+1} ###\n"
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    text += f"{shape.text.strip()}\n"
        return text

    def get_relevant_content(self, query, max_files=None, max_chars=12000):
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

        # 日付が指定されている場合は、その日付を含むファイルを優先
        if date_str:
            search_results.sort(key=lambda x: date_str in x.get('name', ''), reverse=True)
        
        # 処理済みファイルカウンタ
        processed_files = 0
        extracted_content = False  # 本文抽出成功したかどうかのフラグ

        for i, result in enumerate(search_results):
            file_path = result.get('path')
            file_name = result.get('name')
            modified = result.get('modified', '不明')
            
            # 拡張子を取得
            _, ext = os.path.splitext(file_name.lower())

            # ファイルの内容を読み込み
            logger.info(f"ファイル内容を抽出中: {file_name}")
            content = self.read_file_content(file_path)
            
            # 「サポートされていません」というメッセージが含まれているかチェック
            unsupported = ("サポートされていません" in content) or ("抽出できませんでした" in content)
            
            # 本文抽出成功したかどうかを判定
            content_success = len(content) > 200 and not unsupported
            if content_success:
                extracted_content = True

            # コンテンツのプレビューを追加（文字数制限あり）
            preview_length = min(3000, len(content))  # 1ファイルあたり最大3000文字
            preview = content[:preview_length]
            if len(content) > preview_length:
                preview += "...\n(一部抜粋)"

            file_content = f"=== ファイル {i+1}: {file_name} ===\n"
            file_content += f"更新日時: {modified}\n"
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
            processed_files += 1
            
            # 十分な情報が得られたら終了（最大でも3ファイルまで）
            if processed_files >= 3:
                if i + 1 < len(search_results):
                    relevant_content += f"\n（他に{len(search_results) - processed_files}件のファイルがありますが、表示を省略しています）"
                break

        # 本文抽出の結果をログに記録
        if extracted_content:
            logger.info("少なくとも1つのファイルから本文を正常に抽出できました")
        else:
            logger.warning("すべてのファイルで本文抽出に問題がありました")
            if processed_files > 0:
                relevant_content += "\n注意: ファイルは見つかりましたが、内容の抽出に制限があります。Office系ファイルからのテキスト抽出には制限があります。\n"

        return relevant_content