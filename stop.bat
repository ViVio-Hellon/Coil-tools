@echo off
rem  --- keep this file in CP932 (Shift-JIS) with CRLF line endings ---
rem  Keep everything above the chcp line ASCII; see start.bat for why.
chcp 932 >nul 2>&1
rem ===================================================================
rem  コイル梱包ツール (Web版) 停止
rem
rem  このアプリだけを止めます(3つの機能とも終わります)。同じPCで動く
rem  他の Python アプリは影響を受けません(基盤仕様書 2.8)。
rem  資材計算の取り込みが走っているときは止めずに知らせます。
rem  中断してよければ  stop.bat --force
rem
rem  戻り値(業務ツール統合ランチャー 1.8.9 の決まり。統合 1.2.64):
rem      0 = 止めた・もともと動いていない
rem      3 = 処理中・保存していない入力があるので止めない(出力の1行目が理由)
rem      1 = そのほかで止められない(デスクトップ版が動いている・応答しない など)
rem  /nopause を付けるか、環境変数 STOP_NOPAUSE=1 なら、最後の一時停止を飛ばす
rem  (ランチャーは STOP_NOPAUSE=1 を付けて呼ぶ)。
rem  フォルダの名前に空白・全角の空白・半角の括弧があっても動くよう、
rem  道はかならず " で囲み、括弧の中には書かない。
rem ===================================================================
setlocal

set "NOPAUSE=0"
if "%STOP_NOPAUSE%"=="1" set "NOPAUSE=1"
for %%A in (%*) do if /i "%%~A"=="/nopause" set "NOPAUSE=1"
set "CODE=1"

rem  共有フォルダに置かれていても動くよう pushd を使う(start.bat と同じ)
pushd "%~dp0"
if errorlevel 1 goto :nofolder
title コイル梱包ツール - 停止

python --version >nul 2>&1
if errorlevel 1 goto :nopython

python process_manager.py %*
set "RC=%errorlevel%"
if "%RC%"=="0" goto :stopped
if "%RC%"=="3" set "CODE=3"
echo.
echo 止められませんでした。上のメッセージを確認してください。
echo 実行中の処理があるときは、中断してよければ次を実行してください:
echo     stop.bat --force
echo.
goto :failed

:stopped
popd
endlocal
exit /b 0

:nopython
echo [エラー] Python が見つかりません。
goto :failed

:nofolder
echo [エラー] このフォルダに移動できませんでした: "%~dp0"
if not "%NOPAUSE%"=="1" pause
endlocal & exit /b 1

:failed
if not "%NOPAUSE%"=="1" pause
popd
endlocal & exit /b %CODE%
