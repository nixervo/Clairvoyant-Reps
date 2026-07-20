@echo off
pushd "%~dp0"
python -m nuitka --standalone --onefile ^
  --enable-plugin=anti-bloat ^
  --include-module=pyamf ^
  --include-module=rich ^
  --include-module=pyamf.amf3 ^
  --include-module=pyamf.amf0 ^
  --include-module=pyamf.remoting ^
  --include-module=pyamf.util.pure ^
  --output-dir=dist ^
  tui.py
popd
