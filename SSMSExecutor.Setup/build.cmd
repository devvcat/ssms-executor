@echo off
setlocal

set APP_VERSION=2.0.2-alpha
set VS_VERSION=2017
set VS_PRODUCT=Community
set VS_DEVCMD=C:\Program Files (x86)\Microsoft Visual Studio\%VS_VERSION%\%VS_PRODUCT%\Common7\Tools
set OUT_DIR=%CD%\Temp\SSMSExecutor
set INNO_COMPILER_DIR=C:\Program Files (x86)\Inno Setup 5
set ZIP_DIR=C:\Program Files\7-Zip

set PATH=%PATH%;%ZIP_DIR%
set PATH=%PATH%;%INNO_COMPILER_DIR%

rmdir /S /Q %OUT_DIR%
del Temp\files.rsp

pushd %VS_DEVCMD%
call VsDevCmd.bat
popd
cls

msbuild /target:Clean,Rebuild /p:Configuration=Release,OutDir=%OUT_DIR% /nologo /ds ..\SSMSExecutor\SSMSExecutor.csproj
copy build_files.rsp Temp\

call ISCC.exe /O"Output\" /Qp /DAppVersion=%APP_VERSION% Setup.iss

pushd Temp
call 7z.exe a -bt -bb3 SSMSExecutor-%APP_VERSION%.zip @build_files.rsp
move SSMSExecutor-%APP_VERSION%.zip ..\Output
popd

endlocal
@echo on