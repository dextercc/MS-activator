@echo off
:ADMIN
openfiles >nul 2>nul ||(
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
"%temp%\getadmin.vbs" >nul 2>&1
goto:eof
)
del /f /q "%temp%\getadmin.vbs" >nul 2>nul

for /f "tokens=6 delims=[]. " %%G in ('ver') do set win=%%G

pushd "%~dp0"
Title KMS activator  ---v1.1 by dexter


set servername=xykz.f3322.org
for %%a in (1 2 3 4 5) do set go%%a=y

:zM
echo         ********************************************************
echo                    Microsoft KMS 激活工具   ---by dexter
echo         *********************************************************
echo                          1.Microsoft Windows激活(32位和64位)
echo                          2.Microsoft Office激活(32位)
echo                          3.Microsoft Office激活(64位)
echo                          4.遇到激活失败输入4更换服务器
echo                          0.退出

set choice=
set /p choice=    请输入选项（1/2/3/4/0）：
if defined go%choice% @ goto :z%choice%
if /i "%choice%" equ "0" @exit
echo 无效输入！！
goto :zM


:z1
echo         ********************************************************
echo                    Microsoft KMS 激活工具   ---by dexter
echo                            ***Windows激活***
echo                             请选择Windows版本
echo         *********************************************************
echo                          1.Windows 10 Professional
echo                          2.Windows 10 Home
echo                          3.Windows 8.1 Professional
echo                          4.Windows 7 Professional
echo                          0.回主菜单

set choice=
set /p choice=请选择Windows版本：
if defined go%choice% @goto :s1%choice%
if /i "%choice%" equ "0" @ goto :zM
echo 无效输入！！
goto :z1

:z2
echo         ********************************************************
echo                    Microsoft KMS 激活工具   ---by dexter
echo                            ***Office激活***
echo                         请选择Office版本(32位)
echo         *********************************************************
echo                          1.Office 2016 Professional Plus
echo                          2.Visio 2016 Professional
echo                          3.Office 2013 Professional Plus
echo                          4.Visio 2013 Professional
echo                          0.回主菜单

set bit6432= (x86)
set choice=
set /p choice=请选择Office版本：
if defined go%choice% @goto :s2%choice%
if /i "%choice%" equ "0" @ goto :zM
echo 无效输入！！
goto :z2

:z3
echo         ********************************************************
echo                    Microsoft KMS 激活工具   ---by dexter
echo                            ***Office激活***
echo                         请选择Office版本(64位)
echo         *********************************************************
echo                          1.Office 2016 Professional Plus
echo                          2.Visio 2016 Professional
echo                          3.Office 2013 Professional Plus
echo                          4.Visio 2013 Professional
echo                          0.回主菜单

set bit6432=
set choice=
set /p choice=请选择Office版本：
if defined go%choice% @goto :s2%choice%
if /i "%choice%" equ "0" @ goto :zM
echo 无效输入！！
goto :z3

:z4
echo         ********************************************************
echo                    Microsoft KMS 激活工具   ---by dexter
echo                            ***服务器列表***
echo                              ***请选择***
echo         *********************************************************
echo                          1.win10.co
echo                          2.kms.chinancce.com
echo                          3.kms.lotro.cc
echo                          4.www.zgbs.cc
echo                          5.kms-win.msdn123.com
echo	                        0.返回主菜单

set choice=
set /p choice=请选择激活服务器：
if defined go%choice% @goto :s4%choice%
if /i "%choice%" equ "0" @ goto :zM
echo 无效输入！！
goto :z4

:s41
set servername=win10.co
echo 已更换服务器到win10.co
goto :zM
:s42
set servername=kms.chinancce.com
echo 已更换服务器到kms.chinancce.com
goto :zM
:s43
set servername=kms.lotro.cc
echo 已更换服务器到kms.lotro.cc
goto :zM
:s44
set servername=www.zgbs.cc
echo 已更换服务器到www.zgbs.cc
goto :zM
:s45
set servername=kms-win.msdn123.com
echo 已更换服务器到kms-win.msdn123.com
goto :zM


:s11
echo 遇到对话框的时候点击确定
echo 正在激活Windows 10 Professional...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
slmgr.vbs /ato
slmgr.vbs /xpr
echo 程序运行完成，按任意键退出
pause
exit


:s12
echo 遇到对话框的时候点击确定
echo 正在激活Windows 10 Home...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
slmgr.vbs /ato
slmgr.vbs /xpr
echo 程序运行完成，按任意键退出
pause
exit

:s13
echo 遇到对话框的时候点击确定
echo 正在激活Windows 8.1 Professional...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
slmgr.vbs /ato
slmgr.vbs /xpr
echo 程序运行完成，按任意键退出
pause
exit

:s14
echo 遇到对话框的时候点击确定
echo 正在激活Windows 7 Professional...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
slmgr.vbs /ato
slmgr.vbs /xpr
echo 程序运行完成，按任意键退出
pause
exit


:s21
cd C:\Program Files%bit6432%\Microsoft Office\Office16
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=请输入当前Key的后5位，回车结束：
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo 程序运行完成，按任意键退出
pause
exit

:s22
cd C:\Program Files%bit6432%\Microsoft Office\Office16
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=请输入当前Key的后5位，回车结束：
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo 程序运行完成，按任意键退出
pause
exit

:s23
cd C:\Program Files%bit6432%\Microsoft Office\Office15
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=请输入当前Key的后5位，回车结束：
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo 程序运行完成，按任意键退出
pause
exit

:s24
cd C:\Program Files%bit6432%\Microsoft Office\Office15
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=请输入当前Key的后5位，回车结束：
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:J484Y-4NKBF-W2HMG-DBMJC-PGWR7
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo 程序运行完成，按任意键退出
pause
exit
