@echo off
cls
echo ��ӭʹ�������Զ��Ǽǹ���
pause
cls
echo ��ȷ��fileoutĿ¼�¡�����Ǽ�.xls���͡������ʼ��Ǽ�.xls���ļ������ڴ�״̬��
set /p=�Ƿ����(y or n)�� <nul
set /p confirm=
if %confirm%==y (
cls
echo ׼����������...
"jre7/bin/java.exe" -Dfile.encoding=UTF-8 -jar ManuscriptCollector.jar
)