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
echo                    Microsoft KMS �����   ---by dexter
echo         *********************************************************
echo                          1.Microsoft Windows����(32λ��64λ)
echo                          2.Microsoft Office����(32λ)
echo                          3.Microsoft Office����(64λ)
echo                          4.��������ʧ������4����������
echo                          0.�˳�

set choice=
set /p choice=    ������ѡ�1/2/3/4/0����
if defined go%choice% @ goto :z%choice%
if /i "%choice%" equ "0" @exit
echo ��Ч���룡��
goto :zM


:z1
echo         ********************************************************
echo                    Microsoft KMS �����   ---by dexter
echo                            ***Windows����***
echo                             ��ѡ��Windows�汾
echo         *********************************************************
echo                          1.Windows 10 Professional
echo                          2.Windows 10 Home
echo                          3.Windows 8.1 Professional
echo                          4.Windows 7 Professional
echo                          0.�����˵�

set choice=
set /p choice=��ѡ��Windows�汾��
if defined go%choice% @goto :s1%choice%
if /i "%choice%" equ "0" @ goto :zM
echo ��Ч���룡��
goto :z1

:z2
echo         ********************************************************
echo                    Microsoft KMS �����   ---by dexter
echo                            ***Office����***
echo                         ��ѡ��Office�汾(32λ)
echo         *********************************************************
echo                          1.Office 2016 Professional Plus
echo                          2.Visio 2016 Professional
echo                          3.Office 2013 Professional Plus
echo                          4.Visio 2013 Professional
echo                          0.�����˵�

set bit6432= (x86)
set choice=
set /p choice=��ѡ��Office�汾��
if defined go%choice% @goto :s2%choice%
if /i "%choice%" equ "0" @ goto :zM
echo ��Ч���룡��
goto :z2

:z3
echo         ********************************************************
echo                    Microsoft KMS �����   ---by dexter
echo                            ***Office����***
echo                         ��ѡ��Office�汾(64λ)
echo         *********************************************************
echo                          1.Office 2016 Professional Plus
echo                          2.Visio 2016 Professional
echo                          3.Office 2013 Professional Plus
echo                          4.Visio 2013 Professional
echo                          0.�����˵�

set bit6432=
set choice=
set /p choice=��ѡ��Office�汾��
if defined go%choice% @goto :s2%choice%
if /i "%choice%" equ "0" @ goto :zM
echo ��Ч���룡��
goto :z3

:z4
echo         ********************************************************
echo                    Microsoft KMS �����   ---by dexter
echo                            ***�������б�***
echo                              ***��ѡ��***
echo         *********************************************************
echo                          1.win10.co
echo                          2.kms.chinancce.com
echo                          3.kms.lotro.cc
echo                          4.www.zgbs.cc
echo                          5.kms-win.msdn123.com
echo	                      0.�������˵�

set choice=
set /p choice=��ѡ�񼤻��������
if defined go%choice% @goto :s4%choice%
if /i "%choice%" equ "0" @ goto :zM
echo ��Ч���룡��
goto :z4

:s41
set servername=win10.co
echo �Ѹ�����������win10.co
goto :zM
:s42
set servername=kms.chinancce.com
echo �Ѹ�����������kms.chinancce.com
goto :zM
:s43
set servername=kms.lotro.cc
echo �Ѹ�����������kms.lotro.cc
goto :zM
:s44
set servername=www.zgbs.cc
echo �Ѹ�����������www.zgbs.cc
goto :zM
:s45
set servername=kms-win.msdn123.com
echo �Ѹ�����������kms-win.msdn123.com
goto :zM


:s11
echo �����Ի����ʱ����ȷ��
echo ���ڼ���Windows 10 Professional...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
slmgr.vbs /ato
slmgr.vbs /xpr
echo ����������ɣ���������˳�
pause
exit


:s12
echo �����Ի����ʱ����ȷ��
echo ���ڼ���Windows 10 Home...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
slmgr.vbs /ato
slmgr.vbs /xpr
echo ����������ɣ���������˳�
pause
exit

:s13
echo �����Ի����ʱ����ȷ��
echo ���ڼ���Windows 8.1 Professional...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
slmgr.vbs /ato
slmgr.vbs /xpr
echo ����������ɣ���������˳�
pause
exit

:s14
echo �����Ի����ʱ����ȷ��
echo ���ڼ���Windows 7 Professional...
slmgr.vbs /upk
slmgr.vbs /skms %servername%
slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
slmgr.vbs /ato
slmgr.vbs /xpr
echo ����������ɣ���������˳�
pause
exit


:s21
cd C:\Program Files%bit6432%\Microsoft Office\Office16
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=�����뵱ǰKey�ĺ�5λ���س�������
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo ����������ɣ���������˳�
pause
exit

:s22
cd C:\Program Files%bit6432%\Microsoft Office\Office16
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=�����뵱ǰKey�ĺ�5λ���س�������
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo ����������ɣ���������˳�
pause
exit

:s23
cd C:\Program Files%bit6432%\Microsoft Office\Office15
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=�����뵱ǰKey�ĺ�5λ���س�������
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo ����������ɣ���������˳�
pause
exit

:s24
cd C:\Program Files%bit6432%\Microsoft Office\Office15
cscript ospp.vbs /sethst:%servername%
cscript ospp.vbs /dstatus
set existKey=
set /p existKey=�����뵱ǰKey�ĺ�5λ���س�������
cscript ospp.vbs /unpkey:%existKey%
cscript ospp.vbs /inpkey:J484Y-4NKBF-W2HMG-DBMJC-PGWR7
cscript ospp.vbs /act
cscript ospp.vbs /dstatus
echo ����������ɣ���������˳�
pause
exit
