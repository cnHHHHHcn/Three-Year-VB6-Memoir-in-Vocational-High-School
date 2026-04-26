@echo off
set /p strpath1 = 请输入要剪切的文件
set /p strpath2 = 请输入路径
copy %strpath1% %strpath2% 
del %strpath1%
pause