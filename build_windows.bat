@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === 客服总结：打包为 Windows exe（PyInstaller）===
echo 需已安装 Python 3.10+，且建议 64 位。
echo.

if not exist ".venv_build\" (
  echo 创建虚拟环境 .venv_build ...
  py -3.11 -m venv .venv_build 2>nul
  if errorlevel 1 py -3.10 -m venv .venv_build
  if errorlevel 1 python -m venv .venv_build
)
call ".venv_build\Scripts\activate.bat"
python -m pip install -U pip
pip install -r requirements.txt pyinstaller

echo.
echo 开始打包（首次较慢，约数分钟）...
pyinstaller "%~dp0客服总结.spec" --noconfirm
if errorlevel 1 (
  echo 打包失败，请把上方报错复制发我。
  pause
  exit /b 1
)

echo.
echo 完成。
echo 可执行文件：  dist\客服总结\客服总结.exe
echo 分发时请拷贝整个文件夹  dist\客服总结\  （内含依赖 dll），不要只拷单个 exe。
echo 用户数据与密钥保存在 exe 同目录下的 data\ 文件夹。
echo.
pause
