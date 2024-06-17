import os

def save():
    os.system('pip freeze > requirements.txt')
    print('Зависимости сохранены в файл requirements.txt')
    os.system('pause')

def venv():
    commands = [
        # '@echo on',
        'echo Установка виртуального окружения в папку VENV . . .',
        r'call py -m venv vemv',
        'echo Виртуальная срежа создана . . . . . . . . . . . . ',
        'echo Обновление pip . . . . . . . . . . . . . . . . . .',
        r'call .\vemv\Scripts\python.exe -m pip install --upgrade pip',
        'echo Обновление PIP завершено . . . . . . . . . . . . .',
        'echo Установка пакетов . . . . . . . . . . . . . . . . .',
        r'call .\vemv\Scripts\pip.exe install -r requirements.txt',
        'echo Загружены пакеты в venv из файла requirements.txt ...',
        'echo  -------------------------------------------------',
        'pause'
    ]

    text = ' & '.join(commands)
    os.system(text)


if __name__ == '__main__':
    venv()

