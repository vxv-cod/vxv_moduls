@ECHO off 
CALL python -V

@REM Создаем переменную content с адресом proxy
call set /p content=< proxy_port.ini
@REM Создаем переменную среду сеанса CMD pip_proxy (назвать можно и так: https_proxy, PIP_PROXY)
CALL set pip_proxy=%content%

@REM Справка: Команда set (SET) читает переменные сеансы CMD и переменные среды пользователя

@REM Варианты указания прокси:
@REM ECHO Устанавка переменной среды пользователя PIP_PROXY через SETX ...
@REM SETX PIP_PROXY http://tmn-tnnc-proxy.rosneft.ru:9090
@REM ECHO Устанавка переменной среды PIP_CONFIG_FILE с ссылкой на pip.ini из текущей папки...
@REM CALL SETX PIP_CONFIG_FILE %cd%\pip.ini

@REM Создание файла конфигурации %APPDATA%\pip\pip.ini для установщика PIP
IF NOT EXIST "%APPDATA%\pip\pip.ini" (
	call pip config set global.trusted-host "pypi.python.org pypi.org files.pythonhosted.org"
	call pip config set global.user false
	@REM global.proxy нужна для работы pip (проверенная версия от 24.1.1)
	call pip config set global.proxy %content%
	@REM call python -m pip config debug	
)

IF NOT EXIST "%cd%\venv" (
	ECHO Создается виртуального окрружения: "%cd%\venv" ... 
	CALL python -m venv venv
)

ECHO Активируем venv ...
CALL venv\Scripts\activate.bat 
ECHO Виртуальное окружение активировано: "%cd%\venv" ...

ECHO Обновление модуля pip ...
CALL venv\Scripts\python.exe -m pip install --upgrade pip

ECHO Обновление модуля setuptools ...
CALL venv\Scripts\python.exe -m pip install --upgrade setuptools

ECHO Текущее состояние пакетов ...
CALL pip list

@REM Устанавка пакетв из файла requirements.txt
IF EXIST "%cd%\requirements.txt" (
	ECHO Установка пакетов из файла requirements.txt ...
	CALL pip install -r requirements.txt
	CALL pip list
	ECHO Установка пакетов из файла requirements.txt завершена.
) ELSE (
	ECHO Для установки пакетов в виртуальное окружение создайте файл requirements.txt в текущей папке, с указанием названий нужных пакетов ...
	ECHO Для ручной установки пропишите команду: pip install Имя_пакета
)

@REM CALL pause
@REM CALL deactivate
cmd 

