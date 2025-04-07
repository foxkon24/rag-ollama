# routes.py内の署名検証部分を改善

# 変更対象箇所: '/webhook' ルート内の署名検証部分
@app.route('/webhook', methods=['POST'])
def teams_webhook_handler():
    """
    Teamsからのwebhookリクエストを処理する（OneDrive検索機能付き）
    """
    try:
        logger.info("Webhookリクエストを受信しました")
        logger.info(f"リクエストヘッダー: {request.headers}")
        logger.debug(f"リクエストデータ: {request.get_data()}")  # debugレベルに変更

        # 署名検証をスキップするかどうか（.env設定より）
        skip_verification = config['SKIP_VERIFICATION']

        # Teamsからの署名を検証
        signature = request.headers.get('Authorization')
        logger.info(f"受信した認証ヘッダー: {signature}")
        
        if not skip_verification:
            # リクエストデータを取得
            request_data = request.get_data()
            logger.debug(f"検証するデータ: {request_data[:100]}...")
            
            # 署名検証（改善された検証ロジック）
            from teams_auth import verify_teams_token, debug_teams_signature
            
            # デバッグ情報を記録
            debug_info = debug_teams_signature(request_data, config['TEAMS_OUTGOING_TOKEN'])
            logger.debug(f"デバッグ署名情報: {debug_info}")
            
            # 署名を検証
            if not verify_teams_token(request_data, signature, config['TEAMS_OUTGOING_TOKEN']):
                # 本番環境では認証失敗時に403エラーを返す
                if not config['DEBUG']:
                    logger.error("署名の検証に失敗しました。アクセスを拒否します。")
                    return jsonify({"error": "Unauthorized"}), 403
                else:
                    # 開発環境では警告を出して続行
                    logger.warning("署名の検証に失敗しました。デバッグモードで続行します。")
                    logger.warning("本番環境では .env の SKIP_VERIFICATION=1 を設定してください。")
