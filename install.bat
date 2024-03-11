@echo off
cls
setlocal enabledelayedexpansion
set PYTHON_VERSION=3.11.4
set PYTHON_ARCH=amd64
set PYTHON_EXE=python-%PYTHON_VERSION%-%PYTHON_ARCH%.exe
set PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/%PYTHON_EXE%

:: 检查Python是否安装
echo ---------
echo Checking Python installation
set PYTHON_FLG=1
:: 尝试运行python --version命令并获取其输出
for /f "tokens=*" %%a in ('python --version 2^>^&1') do (
    set "python_output=%%a"
)
  
:: 检查python_output变量是否包含有效的Python版本信息
if not "!python_output!" == "" (
    :: 解析python_output以获取版本号
    for /f "tokens=2 delims=." %%v in ("!python_output!") do (
        set "major_version=%%v"
    )
    :: 检查主版本号是否大于或等于3
    if !major_version! GEQ 3 (
        echo Python 3 or higher is installed. Version Information: !python_output!
    ) else (
	    set PYTHON_FLG=0
        echo Python is installed, but the version is lower than 3. Version Information: !python_output!
    )
) else (
    set PYTHON_FLG=0
    echo Python is not installed.
)
echo ---------

:: 如果Python版本低或未安装，打开下载链接
if !PYTHON_FLG!==0 (
    echo The webpage will be opened to download Python-!PYTHON_VERSION!, press any key to continue...
    pause
    start %PYTHON_URL%
    echo After installing Python, press any key to continue…
    pause
    echo ---------
)

:: 安装依赖
echo Installing dependencies
pip install -r requirements.txt
echo ---------

echo After installation, press any key to close!
endlocal
pause
exit

















::
echo 即将打开网页下载Python

:: 下载Python安装程序
start %PYTHON_URL%
  
:: 静默安装Python  
::start /wait %PYTHON_EXE% /quiet InstallAllUsers=1 PrependPath=1 Include_test=0  
  
:: 验证Python是否安装成功  
python --version  
  
:: 如果需要，您可以添加安装pip的步骤，但pip通常与Python一起安装  
  
:: 清理下载的Python安装程序（可选）  
::del %PYTHON_EXE%  
  
echo Python %PYTHON_VERSION% 安装完成！  
pause