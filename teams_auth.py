# teams_auth.py - Teams認証関連の機能（署名検証改善版）
import logging
import hmac
import hashlib
import base64
import json
import re

logger = logging.getLogger(__name__)

def verify_teams_token(request_data, signature, teams_outgoing_token):
    """
    Teamsからのリクエストの署名を検証する（改善版）

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

    # トークンの前処理（末尾の=が欠けている場合の対応）
    if not teams_outgoing_token.endswith('='):
        padded_token = teams_outgoing_token
        while len(padded_token) % 4 != 0:
            padded_token += '='
        teams_outgoing_token = padded_token
        logger.debug(f"認証トークンにパディングを追加: {teams_outgoing_token[:5]}...")

    # Teamsで使用される署名検証方法の網羅的テスト
    verification_methods = [
        verify_raw_data,
        verify_json_compact,
        verify_json_normalized,
        verify_canonical_json,
        verify_utf8_encoded_json,
        verify_json_no_whitespace,
        verify_body_only,
        verify_payload_string
    ]

    for method in verification_methods:
        if method(request_data, clean_signature, teams_outgoing_token):
            logger.info(f"署名検証に成功しました: {method.__name__}")
            return True

    # すべての検証方法が失敗した場合
    logger.warning(f"すべての署名検証方法が失敗しました。詳細なデバッグ情報を出力します。")
    
    # 詳細なデバッグ情報
    debug_info = debug_teams_signature(request_data, teams_outgoing_token)
    logger.debug(f"デバッグ署名情報: {debug_info}")
    
    return False

def verify_raw_data(request_data, signature, token):
    """生のリクエストデータをそのまま使用して検証"""
    try:
        if isinstance(request_data, str):
            request_data = request_data.encode('utf-8')
        
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=token.encode('utf-8'),
                msg=request_data,
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        is_valid = hmac.compare_digest(computed_hmac_b64, signature)
        if is_valid:
            logger.debug(f"生データ(Base64)での検証に成功しました")
        return is_valid
    except Exception as e:
        logger.debug(f"生データでの検証中にエラー: {str(e)}")
        return False

def verify_json_compact(request_data, signature, token):
    """JSONをコンパクト形式に変換して検証"""
    try:
        # JSON文字列としてパース
        if isinstance(request_data, bytes):
            data = json.loads(request_data.decode('utf-8'))
        else:
            data = json.loads(request_data)
        
        # コンパクト形式に再エンコード
        compact_json = json.dumps(data, separators=(',', ':'))
        
        # Base64エンコードされたHMAC計算
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=token.encode('utf-8'),
                msg=compact_json.encode('utf-8'),
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        # 署名を比較
        is_valid = hmac.compare_digest(computed_hmac_b64, signature)
        if is_valid:
            logger.debug(f"コンパクトJSON(Base64)での検証に成功しました")
        return is_valid
    except Exception as e:
        logger.debug(f"JSONコンパクト化での検証中にエラー: {str(e)}")
        return False

def verify_json_normalized(request_data, signature, token):
    """JSONを正規化して検証"""
    try:
        # JSON文字列としてパース
        if isinstance(request_data, bytes):
            data = json.loads(request_data.decode('utf-8'))
        else:
            data = json.loads(request_data)
        
        # 文字列の場合はJSONに変換し、順序を正規化して文字列に戻す
        # キーを辞書順にソートしてJSON文字列を生成
        normalized_json = json.dumps(data, sort_keys=True)
        
        # Base64エンコードされたHMAC計算
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=token.encode('utf-8'),
                msg=normalized_json.encode('utf-8'),
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        is_valid = hmac.compare_digest(computed_hmac_b64, signature)
        if is_valid:
            logger.debug(f"正規化JSON(Base64)での検証に成功しました")
        return is_valid
    except Exception as e:
        logger.debug(f"JSON正規化での検証中にエラー: {str(e)}")
        return False

def verify_canonical_json(request_data, signature, token):
    """正規化JSONをカノニカル形式に変換して検証"""
    try:
        # JSON文字列としてパース
        if isinstance(request_data, bytes):
            data = json.loads(request_data.decode('utf-8'))
        else:
            data = json.loads(request_data)
        
        # カノニカルJSON形式（キーをソート＋コンパクト形式）
        canonical_json = json.dumps(data, sort_keys=True, separators=(',', ':'))
        
        # Base64エンコードされたHMAC計算
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=token.encode('utf-8'),
                msg=canonical_json.encode('utf-8'),
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        is_valid = hmac.compare_digest(computed_hmac_b64, signature)
        if is_valid:
            logger.debug(f"カノニカルJSON(Base64)での検証に成功しました")
        return is_valid
    except Exception as e:
        logger.debug(f"カノニカルJSONでの検証中にエラー: {str(e)}")
        return False

def verify_utf8_encoded_json(request_data, signature, token):
    """UTF-8エンコードしたJSONで検証"""
    try:
        # JSON文字列を取得
        if isinstance(request_data, bytes):
            json_str = request_data.decode('utf-8')
        else:
            json_str = request_data
        
        # 明示的にUTF-8エンコード
        json_bytes = json_str.encode('utf-8')
        
        # Base64エンコードされたHMAC計算
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=token.encode('utf-8'),
                msg=json_bytes,
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        is_valid = hmac.compare_digest(computed_hmac_b64, signature)
        if is_valid:
            logger.debug(f"UTF-8エンコードJSON(Base64)での検証に成功しました")
        return is_valid
    except Exception as e:
        logger.debug(f"UTF-8エンコードでの検証中にエラー: {str(e)}")
        return False

def verify_json_no_whitespace(request_data, signature, token):
    """空白を除去したJSONで検証"""
    try:
        # JSON文字列を取得
        if isinstance(request_data, bytes):
            json_str = request_data.decode('utf-8')
        else:
            json_str = request_data
        
        # 空白を除去
        no_whitespace = re.sub(r'\s', '', json_str)
        
        # Base64エンコードされたHMAC計算
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=token.encode('utf-8'),
                msg=no_whitespace.encode('utf-8'),
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        is_valid = hmac.compare_digest(computed_hmac_b64, signature)
        if is_valid:
            logger.debug(f"空白除去JSON(Base64)での検証に成功しました")
        return is_valid
    except Exception as e:
        logger.debug(f"空白除去での検証中にエラー: {str(e)}")
        return False

def verify_body_only(request_data, signature, token):
    """body部分だけを抽出して検証（Microsoft Teams特有の処理）"""
    try:
        # JSON文字列としてパース
        if isinstance(request_data, bytes):
            data = json.loads(request_data.decode('utf-8'))
        else:
            data = json.loads(request_data)
        
        # Teamsの場合、bodyフィールドのみを使う可能性がある
        if 'body' in data:
            body_json = json.dumps(data['body'], separators=(',', ':'))
            
            # Base64エンコードされたHMAC計算
            computed_hmac_b64 = base64.b64encode(
                hmac.new(
                    key=token.encode('utf-8'),
                    msg=body_json.encode('utf-8'),
                    digestmod=hashlib.sha256
                ).digest()
            ).decode('utf-8')
            
            is_valid = hmac.compare_digest(computed_hmac_b64, signature)
            if is_valid:
                logger.debug(f"bodyフィールドのみ(Base64)での検証に成功しました")
            return is_valid
    except Exception as e:
        logger.debug(f"bodyフィールド検証中にエラー: {str(e)}")
    
    return False

def verify_payload_string(request_data, signature, token):
    """文字列形式のペイロードをそのまま使用（Microsoft Teams特有の処理）"""
    try:
        # リクエストデータを文字列として扱う
        payload_str = request_data
        if isinstance(request_data, bytes):
            payload_str = request_data.decode('utf-8')
        
        # 直接文字列を使用してHMACを計算
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=token.encode('utf-8'),
                msg=payload_str.encode('utf-8'),
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        is_valid = hmac.compare_digest(computed_hmac_b64, signature)
        if is_valid:
            logger.debug(f"ペイロード文字列(Base64)での検証に成功しました")
        return is_valid
    except Exception as e:
        logger.debug(f"ペイロード文字列での検証中にエラー: {str(e)}")
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
            str_data = request_data.decode('utf-8', errors='replace')
        else:
            str_data = request_data
            
        str_b64_signature = base64.b64encode(
            hmac.new(
                key=token_bytes,
                msg=str_data.encode('utf-8'),
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        results["string_base64"] = str_b64_signature
        
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
            
        except json.JSONDecodeError:
            results["json"] = "リクエストデータはJSON形式ではありません"
        except Exception as json_err:
            results["json_error"] = str(json_err)
        
    except Exception as e:
        results["error"] = str(e)
        
    return results
