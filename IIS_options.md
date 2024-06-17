Редактируем разрешения для папок:
1.	Глобальной папки Python3*:  
•	Правой кнопкой по папке => Свойства => Безопасность
•	Группы или пользователи  => Изменить   => Разрешения для группы
•	Разрешения для группы  => Добавить => Выбор «Пользователя» или «группы»
•	Выбор «Пользователя» или «группы»  => Размещение  => Имя_ПК  => ОК
•	Выбор «Пользователя» или «группы»  => Размещение  => Дополнительно  => Поиск   => Все => ОК => ОК.
2.	Папки с проектом SiteName:  
•	…
•	Выбор «Пользователя» или «группы»  => Размещение  => Имя_ПК  => ОК
•	Выбор «Пользователя» или «группы»  => Дополнительно  => Поиск   => IIS_IUSRS  => ОК
•	добавляем пользователя IIS AppPool\Имя_пула_приложения (IIS AppPool\SiteNamePool). Необходимо предварительно создать данный пул приложения SiteNamePool в Диспетчере служб IIS.
Настройка IIS верхняя строка в диспетчере с именем ПК:
1.	Параметры FastCGI => Добавить приложение:
•	Полный путь: C:\SiteName\venv\Scripts\python.exe
•	Аргументы: C:\SiteName\venv\Scripts\wfastcgi.exe
Добавление Пула приложения и самого сайта:
1.	Пулы приложения => Добавить пул приложений:
 	![Alt text](Pool-1.png)
2.	Добавление веб-сайта:
 ![Alt text](site-1.png)
3.	Нажимаем на появившийся сайт и выбираем Сопоставление обработчиков
4.	Добавление сопоставление модуля
•	Путь запроса: *
•	Модуль: FastCgiModule
•	Исполняемый файл: C:\SiteName\venv\Scripts\python.exe| C:\SiteName\venv\Scripts\wfastcgi.exe
•	Имя: FastApi_FastCgiModule
•	Ограничение запроса: убираем галочку

5.	Создаем в корневой папки с проектом файл web.config и заполняем его следующим:
<?xml version="1.0" encoding="UTF-8"?>
    <configuration>
        <system.webServer>
            <handlers>
                <add name="vxv_FastApi" path="*" verb="*" modules="FastCgiModule" 
                scriptProcessor="C:\SiteName\venv\Scripts\python.exe|C:\SiteName\venv\Scripts\wfastcgi.exe" 
                resourceType="Unspecified" 
                requireAccess="Script" /> 
            </handlers> 
        </system.webServer>
        <appSettings>
                <add key="PYTHONPATH" value="C:\SiteName"/>
                <!-- <add key="WSGI_HANDLER" value="main.wsgi_app" /> -->
                <add key="WSGI_HANDLER" value="main.wsgi_app" />
        </appSettings>
    </configuration>

Где wsgi_app имя переменной приложения
C:\inetpub - сюда создавать проекты про рекомендациям Максима

