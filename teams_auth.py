# teams_auth.py - Teams認証関連の機能（署名検証改善版）
import logging
import hmac
import hashlib
import base64
import json

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

    try:
        # バイナリに変換
        if isinstance(request_data, str):
            request_data = request_data.encode('utf-8')
        
        # HMAC-SHA256を使用した署名の計算方法1
        computed_hmac_hex = hmac.new(
            key=teams_outgoing_token.encode('utf-8'),
            msg=request_data,
            digestmod=hashlib.sha256
        ).hexdigest()
        
        # HMAC-SHA256を使用した署名の計算方法2 (Base64エンコード)
        computed_hmac_b64 = base64.b64encode(
            hmac.new(
                key=teams_outgoing_token.encode('utf-8'),
                msg=request_data,
                digestmod=hashlib.sha256
            ).digest()
        ).decode('utf-8')
        
        logger.debug(f"計算したHMAC(hex): {computed_hmac_hex}")
        logger.debug(f"計算したHMAC(b64): {computed_hmac_b64}")
        
        # 両方の方法で検証
        is_valid_hex = hmac.compare_digest(computed_hmac_hex, clean_signature)
        is_valid_b64 = hmac.compare_digest(computed_hmac_b64, clean_signature)
        
        if is_valid_hex:
            logger.info("16進数HMAC署名の検証に成功しました")
            return True
        elif is_valid_b64:
            logger.info("Base64 HMAC署名の検証に成功しました")
            return True
        else:
            # ログに詳細を記録
            logger.warning(f"署名検証失敗: 受信={clean_signature}, 計算(hex)={computed_hmac_hex}, 計算(b64)={computed_hmac_b64}")
            
            # Teamsが送る形式でJSON文字列を再計算してみる
            try:
                # JSON文字列としてパース
                json_data = json.loads(request_data)
                # 再エンコード（JSONフォーマットの違いを吸収）
                json_str = json.dumps(json_data, separators=(',', ':'))
                
                # 再計算したHMACの生成
                computed_hmac_json_hex = hmac.new(
                    key=teams_outgoing_token.encode('utf-8'),
                    msg=json_str.encode('utf-8'),
                    digestmod=hashlib.sha256
                ).hexdigest()
                
                computed_hmac_json_b64 = base64.b64encode(
                    hmac.new(
                        key=teams_outgoing_token.encode('utf-8'),
                        msg=json_str.encode('utf-8'),
                        digestmod=hashlib.sha256
                    ).digest()
                ).decode('utf-8')
                
                logger.debug(f"JSON再エンコード後のHMAC(hex): {computed_hmac_json_hex}")
                logger.debug(f"JSON再エンコード後のHMAC(b64): {computed_hmac_json_b64}")
                
                # 再エンコード後の検証
                is_valid_json_hex = hmac.compare_digest(computed_hmac_json_hex, clean_signature)
                is_valid_json_b64 = hmac.compare_digest(computed_hmac_json_b64, clean_signature)
                
                if is_valid_json_hex or is_valid_json_b64:
                    logger.info("JSON再エンコード後の署名検証に成功しました")
                    return True
            except json.JSONDecodeError:
                logger.debug("リクエストデータはJSON形式ではありません")
            except Exception as json_err:
                logger.debug(f"JSON検証中にエラー: {str(json_err)}")
            
            return False

    except Exception as e:
        logger.error(f"署名検証中にエラーが発生しました: {str(e)}")
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
        
        # JSON処理を試みる
        try:
            json_data = json.loads(data_bytes)
            # 様々なJSON形式で試す
            json_formats = {
                "compact": json.dumps(json_data, separators=(',', ':')),
                "pretty": json.dumps(json_data, indent=2),
                "no_spaces": json.dumps(json_data),
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
            
        except json.JSONDecodeError:
            results["json"] = "リクエストデータはJSON形式ではありません"
        
    except Exception as e:
        results["error"] = str(e)
        
    return results
