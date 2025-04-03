# logger.py - ロギング設定
import os
import logging

def setup_logger(script_dir):
    """
    ロギングを設定する

    Args:
        script_dir: スクリプトのディレクトリパス

    Returns:
        logging.Logger: 設定されたロガーインスタンス
    """
    # ロギング設定
    logging.basicConfig(
        level=logging.DEBUG,  # 最初はDEBUGレベルで開始
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(script_dir, "ollama_system.log")),
            logging.StreamHandler()
        ]
    )

    return logging.getLogger(__name__)
