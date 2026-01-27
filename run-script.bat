@echo off
REM Excel処理スクリプト実行用バッチファイル
REM タスクスケジューラから実行する場合に使用（非表示実行）

REM スクリプトのディレクトリに移動
cd /d "%~dp0"

REM VBScriptラッパーを使用して非表示で実行
cscript //nologo "%~dp0run-script-hidden.vbs"

REM エラーレベルを返す
exit /b %ERRORLEVEL%

