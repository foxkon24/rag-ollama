# logger.py - ロギング設定
import os
import logging
import sys

def setup_logger(script_dir):
    """
    ロギングを設定する（エンコーディング問題解決版）
    """
    # Windows環境でUTF-8を強制
    if sys.platform == 'win32':
        import codecs
        try:
            # stdoutとstderrをUTF-8に設定
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer)
            sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer)
        except AttributeError:
            # Python 2系またはバッファがない場合の対応
            pass
    
    # ログファイルパス
    log_file = os.path.join(script_dir, "ollama_system.log")
    
    # ロギング設定
    handlers = [
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
    
    # ロガーの設定
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )
    
    # エンコーディングの設定をログに出力
    logger = logging.getLogger(__name__)
    logger.info(f"ロガーを初期化しました。ファイル: {log_file}, エンコーディング: utf-8")
    logger.info(f"Python バージョン: {sys.version}")
    logger.info(f"システムのデフォルトエンコーディング: {sys.getdefaultencoding()}")
    
    return logger
