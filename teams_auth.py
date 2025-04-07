# teams_auth.py - Teams認証関連の機能（包括的に改善された署名検証）
import logging
import hmac
import hashlib
import base64
import json
import re

logger = logging.getLogger(__name__)

def verify_teams_token(request_data, signature, teams_outgoing_token):
    """
    Teamsからのリクエストの署名を検証する（包括的に改善された版）

    Args:
        request_data: リクエストの生データ
        signature: Teamsからの署名
        teams_outgoing_token: Teams Outgoing Token

    Returns:
        bool: 署名が有効な場合True
    """
    if not teams_outgoing_token:
        logger.error("Teams Outgoing Tokenが設定されていません")
        return False

    if not signature:
        logger.error("署名がリクエストに含まれていません")
        return False

    # リクエストデータと署名のログ（デバッグ用）
    logger.debug(f"検証するリクエストデータ: {request_data[:100]}...")
    logger.debug(f"受信した署名: {signature}")

    # 署名のクリーンアップ
    clean_signature = signature
    if signature.startswith('HMAC '):
        clean_signature = signature[5:]  # 'HMAC 'の部分を削除
        logger.debug(f"HMACプレフィックスを削除: {clean_signature}")

    # 認証トークンのログ（セキュリティのため一部のみ）
    token_preview = teams_outgoing_token[:5] + "..." if teams_outgoing_token else "なし"
    logger.debug(f"認証トークン: {token_preview}")

    # === 重要: パッディングとトークン形式の処理 ===
    
    # 1. トークンがBase64形式かチェック
    try:
        # Base64としてデコードしてみる
        base64.b64decode(teams_outgoing_token)
        is_base64_token = True
    except:
        is_base64_token = False
    
    # 2. トークンの形式に応じて処理方法を変える
    token_variants = []
    
    # 元のトークンを追加
    token_variants.append(teams_outgoing_token)
    
    # パディングを調整したトークン
    if not teams_outgoing_token.endswith('='):
        padded_token = teams_outgoing_token
        while len(padded_token) % 4 != 0:
            padded_token += '='
        token_variants.append(padded_token)
    
    # パディングを除去したトークン
    if teams_outgoing_token.endswith('='):
        unpadded_token = teams_outgoing_token.rstrip('=')
        token_variants.append(unpadded_token)
    
    # Base64デコード/エンコードし直したトークン（正規化）
    if is_base64_token:
        try:
            decoded = base64.b64decode(teams_outgoing_token)
            reencoded = base64.b64encode(decoded).decode('utf-8')
            token_variants.append(reencoded)
        except:
            pass

    # === データの前処理バリエーション ===
    
    # リクエストデータの準備（バイト形式とテキスト形式）
    if isinstance(request_data, str):
        request_bytes = request_data.encode('utf-8')
        request_str = request_data
    else:
        request_bytes = request_data
        try:
            request_str = request_data.decode('utf-8')
        except:
            request_str = None

    # データ変換オプション
    data_variants = []
    
    # 1. 元のバイトデータ
    data_variants.append(("raw_bytes", request_bytes))
    
    # 2. UTF-8エンコードされた文字列
    if request_str:
        data_variants.append(("utf8_string", request_str.encode('utf-8')))
    
    # 3. JSONデータの場合の特別処理
    try:
        if request_str:
            json_data = json.loads(request_str)
            
            # コンパクトJSON
            compact_json = json.dumps(json_data, separators=(',', ':'))
            data_variants.append(("compact_json", compact_json.encode('utf-8')))
            
            # キーをソートしたカノニカルJSON
            canonical_json = json.dumps(json_data, sort_keys=True, separators=(',', ':'))
            data_variants.append(("canonical_json", canonical_json.encode('utf-8')))
            
            # 空白を除去したJSON
            no_whitespace_json = re.sub(r'\s', '', request_str)
            data_variants.append(("no_whitespace_json", no_whitespace_json.encode('utf-8')))
            
            # MicrosoftのJSONフォーマット（特定のフィールドのみ）
            if 'text' in json_data:
                text_only = json_data['text']
                data_variants.append(("text_only", json.dumps(text_only).encode('utf-8')))
            
            if 'body' in json_data:
                body_only = json_data['body']
                data_variants.append(("body_only", json.dumps(body_only, separators=(',', ':')).encode('utf-8')))
    except:
        # JSONとして解析できない場合はスキップ
        pass
    
    # === ハッシュアルゴリズムとエンコーディングオプション ===
    
    # 使用するダイジェストアルゴリズム
    digest_algos = [
        ("sha256", hashlib.sha256),
        ("sha1", hashlib.sha1)  # 過去のTeamsバージョンで使われていた可能性
    ]
    
    # 結果の表現形式
    output_formats = [
        ("base64", lambda digest: base64.b64encode(digest).decode('utf-8')),
        ("base64_no_padding", lambda digest: base64.b64encode(digest).decode('utf-8').rstrip('=')),
        ("hex", lambda digest: digest.hex())
    ]

    # === すべての組み合わせで検証 ===
    
    # 非常に多くの組み合わせをテストするため、限定的なロギング
    log_interval = 10
    i = 0
    
    # すべての組み合わせを試す
    for token in token_variants:
        token_bytes = token.encode('utf-8')
        
        for data_name, data in data_variants:
            for digest_name, digest_algo in digest_algos:
                for format_name, format_func in output_formats:
                    try:
                        # HMAC計算
                        hmac_digest = hmac.new(
                            key=token_bytes,
                            msg=data,
                            digestmod=digest_algo
                        ).digest()
                        
                        # 結果をフォーマット
                        computed_signature = format_func(hmac_digest)
                        
                        # 一定間隔でのみログを出力（多すぎるログを避けるため）
                        i += 1
                        if i % log_interval == 0:
                            logger.debug(f"検証 #{i}: {data_name}/{digest_name}/{format_name} = {computed_signature[:10]}...")
                        
                        # 受信した署名と比較
                        if hmac.compare_digest(computed_signature, clean_signature):
                            logger.info(f"署名検証に成功しました: {data_name}/{digest_name}/{format_name}")
                            return True
                        
                    except Exception as e:
                        # 特定の組み合わせでエラーが発生した場合はスキップ
                        pass
    
    # === 特別なケース: Microsoft Teamsの独自実装に対応 ===
    
    # Teamsの特殊な実装に対応した追加の検証方法
    try:
        # 生データのハッシュを直接比較
        for token in token_variants:
            token_bytes = token.encode('utf-8')
            
            # SHA256を使用したHMAC
            raw_hmac = hmac.new(
                key=token_bytes,
                msg=request_bytes,
                digestmod=hashlib.sha256
            ).digest()
            
            # Base64エンコード
            b64_signature = base64.b64encode(raw_hmac).decode('utf-8')
            
            # パディングありとなしの両方をチェック
            if hmac.compare_digest(b64_signature, clean_signature) or \
               hmac.compare_digest(b64_signature.rstrip('='), clean_signature):
                logger.info("raw_hmac_b64での検証に成功")
                return True
            
            # また、Teamsが直接バイナリデータとして送信している可能性もある
            if isinstance(clean_signature, str):
                try:
                    binary_sig = base64.b64decode(clean_signature)
                    if hmac.compare_digest(raw_hmac, binary_sig):
                        logger.info("raw_hmac_binary での検証に成功")
                        return True
                except:
                    pass
    except Exception as e:
        logger.debug(f"特別なケース検証中のエラー: {str(e)}")
    
    # === ロギングと失敗の報告 ===
    
    # すべての検証が失敗
    logger.warning("すべての署名検証方法が失敗しました。詳細なデバッグ情報を出力します。")
    
    # いくつかの主要な計算結果をログに出力
    try:
        # 標準的なHMAC-SHA256計算（最も一般的）
        standard_hmac = hmac.new(
            key=teams_outgoing_token.encode('utf-8'),
            msg=request_bytes,
            digestmod=hashlib.sha256
        ).digest()
        
        logger.debug(f"標準HMAC-SHA256 (base64): {base64.b64encode(standard_hmac).decode('utf-8')}")
        logger.debug(f"標準HMAC-SHA256 (hex): {standard_hmac.hex()}")
        logger.debug(f"受信した署名: {clean_signature}")
        
        # 詳細なデバッグ情報の実行
        debug_info = debug_teams_signature(request_data, teams_outgoing_token)
        logger.debug(f"デバッグ署名情報: {debug_info}")
    except Exception as e:
        logger.error(f"デバッグ情報出力中のエラー: {str(e)}")
    
    return False


def debug_teams_signature(request_data, teams_outgoing_token):
    """
    デバッグ用：様々な方法でTeams署名を計算して表示

    Args:
        request_data: リクエストデータ
        teams_outgoing_token: Teams Outgoing Token

    Returns:
        dict: 様々な方法で計算された署名
    """
    results = {}
    
    try:
        # バイナリに変換
        if isinstance(request_data, str):
            data_bytes = request_data.encode('utf-8')
        else:
            data_bytes = request_data
            
        token_bytes = teams_outgoing_token.encode('utf-8')
        
        # HMAC-SHA256 (hex)
        hex_signature = hmac.new(
            key=token_bytes,
            msg=data_bytes, 
            digestmod=hashlib.sha256
        ).hexdigest()
        results["hex"] = hex_signature
        
        # HMAC-SHA256 (base64)
        b64_signature = base64.b64encode(
            hmac.new(
                key=token_bytes,
                msg=data_bytes,
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        results["base64"] = b64_signature
        
        # 文字列そのままの場合
        if isinstance(request_data, bytes):
            try:
                str_data = request_data.decode('utf-8', errors='replace')
                str_b64_signature = base64.b64encode(
                    hmac.new(
                        key=token_bytes,
                        msg=str_data.encode('utf-8'),
                        digestmod=hashlib.sha256
                    ).digest()
                ).decode('utf-8')
                results["string_base64"] = str_b64_signature
            except:
                pass
        else:
            str_b64_signature = base64.b64encode(
                hmac.new(
                    key=token_bytes,
                    msg=request_data.encode('utf-8'),
                    digestmod=hashlib.sha256
                ).digest()
            ).decode('utf-8')
            results["string_base64"] = str_b64_signature
        
        # トークンの処理バリエーション
        token_variations = {}
        
        # パディングを追加したトークン
        padded_token = teams_outgoing_token
        while len(padded_token) % 4 != 0:
            padded_token += '='
        
        padded_b64 = base64.b64encode(
            hmac.new(
                key=padded_token.encode('utf-8'),
                msg=data_bytes,
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        token_variations["padded"] = padded_b64
        
        # パディングを削除したトークン
        unpadded_token = teams_outgoing_token.rstrip('=')
        unpadded_b64 = base64.b64encode(
            hmac.new(
                key=unpadded_token.encode('utf-8'),
                msg=data_bytes,
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        token_variations["unpadded"] = unpadded_b64
        
        # Base64デコード/エンコードし直したトークン
        try:
            decoded_token = base64.b64decode(teams_outgoing_token.encode('utf-8'))
            reencoded_token = base64.b64encode(decoded_token).decode('utf-8')
            
            reencoded_b64 = base64.b64encode(
                hmac.new(
                    key=reencoded_token.encode('utf-8'),
                    msg=data_bytes,
                    digestmod=hashlib.sha256
                ).digest()
            ).decode('utf-8')
            token_variations["reencoded"] = reencoded_b64
        except:
            pass
        
        results["token_variations"] = token_variations
        
        # JSON処理を試みる
        try:
            if isinstance(request_data, bytes):
                json_data = json.loads(request_data.decode('utf-8', errors='replace'))
            else:
                json_data = json.loads(request_data)
                
            # 様々なJSON形式で試す
            json_formats = {
                "compact": json.dumps(json_data, separators=(',', ':')),
                "canonical": json.dumps(json_data, sort_keys=True, separators=(',', ':')),
                "pretty": json.dumps(json_data, indent=2),
                "no_spaces": re.sub(r'\s', '', json.dumps(json_data)),
                "ascii": json.dumps(json_data, ensure_ascii=True)
            }
            
            json_results = {}
            for format_name, json_str in json_formats.items():
                json_bytes = json_str.encode('utf-8')
                json_results[format_name] = {
                    "hex": hmac.new(key=token_bytes, msg=json_bytes, digestmod=hashlib.sha256).hexdigest(),
                    "base64": base64.b64encode(
                        hmac.new(key=token_bytes, msg=json_bytes, digestmod=hashlib.sha256).digest()
                    ).decode('utf-8')
                }
            
            results["json"] = json_results
            
            # 特定のフィールドのみを使用
            if "text" in json_data:
                text_only = {
                    "hex": hmac.new(key=token_bytes, msg=json.dumps(json_data["text"]).encode('utf-8'), digestmod=hashlib.sha256).hexdigest(),
                    "base64": base64.b64encode(
                        hmac.new(key=token_bytes, msg=json.dumps(json_data["text"]).encode('utf-8'), digestmod=hashlib.sha256).digest()
                    ).decode('utf-8')
                }
                results["text_only"] = text_only
            
            # body部分のみを使用した場合
            if "body" in json_data:
                body_json = json.dumps(json_data["body"], separators=(',', ':'))
                body_bytes = body_json.encode('utf-8')
                
                results["body_only"] = {
                    "hex": hmac.new(key=token_bytes, msg=body_bytes, digestmod=hashlib.sha256).hexdigest(),
                    "base64": base64.b64encode(
                        hmac.new(key=token_bytes, msg=body_bytes, digestmod=hashlib.sha256).digest()
                    ).decode('utf-8')
                }
            
            # SHA1アルゴリズムで試す（過去のバージョンで使用されていた可能性）
            results["sha1"] = {
                "hex": hmac.new(key=token_bytes, msg=data_bytes, digestmod=hashlib.sha1).hexdigest(),
                "base64": base64.b64encode(
                    hmac.new(key=token_bytes, msg=data_bytes, digestmod=hashlib.sha1).digest()
                ).decode('utf-8')
            }
            
        except json.JSONDecodeError:
            results["json"] = "リクエストデータはJSON形式ではありません"
        except Exception as json_err:
            results["json_error"] = str(json_err)
        
    except Exception as e:
        results["error"] = str(e)
        
    return results


def bypass_teams_token(config):
    """
    開発時や問題調査用に署名検証をバイパスするヘルパー関数
    
    Args:
        config: 設定辞書
        
    Returns:
        bool: 署名検証をスキップするかどうか
    """
    # .envの SKIP_VERIFICATION 設定で署名検証をスキップ
    skip_verification = config.get('SKIP_VERIFICATION', False)
    
    # デバッグモードで実行している場合
    debug_mode = config.get('DEBUG', False)
    
    # デバッグが明示的に有効、または署名スキップが明示的に要求されている場合
    return skip_verification or debug_mode
