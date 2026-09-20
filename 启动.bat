@echo off
rem ============================================
rem   简历定制器 一键启动
rem   双击本文件即可：起服务 + 自动打开浏览器
rem   关闭这个黑窗口 = 停止服务
rem ============================================
rem
rem   【编码说明 · 别乱改】
rem   本文件是 GBK 编码，因为中文 Windows 的 cmd 默认代码页是 936(GBK)。
rem   如果存成 UTF-8，cmd 会把中文读成乱码，乱码字符还会把命令行截断，
rem   导致 launch.py 被切成 "hon:" / "unch.py" 这种错误的命令。
rem   （已实测踩过这个坑，不要再改回 UTF-8。）
rem
title 简历定制器
echo.
echo ============================================
echo   简历定制器 启动中...
echo   地址: http://127.0.0.1:5050
echo   服务就绪后自动打开浏览器
echo   关闭本窗口 = 停止服务
echo ============================================
echo.

rem Python 输出按 GBK 编码，避免中文乱码
set PYTHONIOENCODING=gbk

rem ---- 找 Python ----
rem 注意：这台机器系统 PATH 里没有 python，所以必须挨个常见位置试。
rem 找到哪个算哪个，剩下交给 launch.py 检查依赖。

where python >nul 2>&1
if not errorlevel 1 (
    set "PYEXE=python"
    goto :run
)

where py >nul 2>&1
if not errorlevel 1 (
    set "PYEXE=py"
    goto :run
)

rem 本机专属的 Python 路径写进 python_local.txt（一行一个绝对路径）。
rem 那个文件不入库——在这里写死绝对路径，会把你的 Windows 用户名带进公开仓库。
if not exist "python_local.txt" goto :try_common
set /p PYEXE=<python_local.txt
if exist "%PYEXE%" goto :run
echo [提示] python_local.txt 里的路径找不到，改用常见位置： %PYEXE%

:try_common
if exist "C:\ProgramData\anaconda3\python.exe" (
    set "PYEXE=C:\ProgramData\anaconda3\python.exe"
    goto :run
)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    goto :run
)
if exist "%USERPROFILE%\anaconda3\python.exe" (
    set "PYEXE=%USERPROFILE%\anaconda3\python.exe"
    goto :run
)

echo [没找到 Python]
echo.
echo 这台机器的 PATH 里没有 Python。请先安装 Python 3.10 以上版本：
echo   https://www.python.org/downloads/
echo 安装时记得勾选 "Add python.exe to PATH"。
echo.
pause
exit /b 1

:run
echo 使用 Python: %PYEXE%
echo.
"%PYEXE%" launch.py
if errorlevel 1 (
    echo.
    echo 启动失败，详见上方报错。
)

echo.
pause
