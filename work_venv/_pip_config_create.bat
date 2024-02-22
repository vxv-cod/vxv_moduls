@echo off 
call python -V

call pip config set global.trusted-host "pypi.python.org pypi.org files.pythonhosted.org"
call pip config set global.user false

@REM global.proxy нужна для работы pip (проверенная версия от 24.1.1)
call pip config set global.proxy http://tmn-tnnc-proxy.rosneft.ru:9090
@REM call python -m pip config debug
