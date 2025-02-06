# hook-pydantic.py
from PyInstaller.utils.hooks import collect_submodules

# pydantic v2 では 'compiled' 属性が削除されたため、
# 単純に pydantic の全サブモジュールを隠しインポートします。
hiddenimports = collect_submodules("pydantic")
