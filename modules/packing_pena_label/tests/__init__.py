"""試験のあいだ、梱包明細の共有設定(管理者パスワードの控え)を使い捨ての場所へ向ける。

管理者パスワードは梱包明細の共有のもの1つで確かめる(統合 1.2.75)。既定の置き場所は社内の共有
フォルダなので、そのままだと試験が**本物の共有を見に行く**(Windows ではネットワークの場所へ、
Linux ではアプリのフォルダにその名のフォルダを作る)。試験のプロセスに1つ、空の場所を用意する
(一度も変えていない現場と同じ形: 既定のパスワードで通る)。梱包明細の試験は自分で差し替えて戻す。
"""
import atexit
import shutil
import tempfile
from pathlib import Path

from modules.packing_details.meisai import config as _details_config
from modules.packing_details.meisai import shared_settings as _shared_settings

_ROOT = Path(tempfile.mkdtemp(prefix="cpt-shared-pw-"))
atexit.register(shutil.rmtree, _ROOT, True)
(_ROOT / "share").mkdir()
_details_config.USER_CONFIG_PATH = _ROOT / "pc" / "user_config.json"
_details_config.SHARED_DIR = _ROOT / "share" / "【■】_参照用ファイル"
_shared_settings.reset_for_tests()
