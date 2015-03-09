@echo off
cls
echo 欢迎使用来稿自动登记工具
pause
cls
echo 请确认fileout目录下“来稿登记.xls”和“其他邮件登记.xls”文件不处于打开状态！
set /p=是否继续(y or n)？ <nul
set /p confirm=
if %confirm%==y (
cls
echo 准备启动工具...
java -Dfile.encoding=UTF-8 -jar ManuscriptCollector.jar
)