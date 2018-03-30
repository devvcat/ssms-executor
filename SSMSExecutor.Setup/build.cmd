@echo off
setlocal

set VS_VERSION=2017
set VS_PRODUCT=Community
set VS_DEVCMD=C:\Program Files (x86)\Microsoft Visual Studio\%VS_VERSION%\%VS_PRODUCT%\Common7\Tools
set OUT_DIR=%CD%\Temp\Bin
set INNO_COMPILER_DIR=C:\Program Files (x86)\Inno Setup 5

pushd %VS_DEVCMD%
call VsDevCmd.bat
popd

cls
msbuild /target:Clean,Rebuild /p:Configuration=Release,OutDir=%OUT_DIR% /nologo /ds ..\SSMSExecutor\SSMSExecutor.csproj

call "%INNO_COMPILER_DIR%\ISCC.exe" /O"Output\" /Qp Setup.iss

endlocal
@echo on