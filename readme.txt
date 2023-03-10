Скрипт(программа) для расчёта премии для сотрудников отдела продаж.
Принцип работы:
	1) чтение файла zap.xlsx и выбор строк подходящих по условию заданному заказчиком;
	2) чтение файла data.xlsx и выбор строк подходящих по условию заданному заказчиком;
	3) выбор строк, которые присутствуют в выборках пунктов 1, 2;
	4) запись данных в файл report.xlsx
Использованные библиотеки: OpenPyXl, PyInstaller (для преобразования в формат «exe»).
Файлы zap.xlsx, data.xlsx загрузить не могу, так как в них содержится коммерческая тайна.
Для использовании прошу следовать следующим шагам:
	1) скачать или склонировать проект;
	2) создать виртуальное окружение;
	3) воспользоваться командой pip freeze > requirements.txt
	
-------------------------------------------------------------------------------------------

Script (program) to calculate bonuses for sales staff.
Working principle:
	1) reading the file zap.xlsx and selecting rows suitable by the condition set by the customer;
	2. Reading the file data.xlsx and selecting rows suitable for the condition set by the customer;
	3) selecting rows that are present in the selections of items 1 and 2;
	4) writing data to the report.xlsx file.
Used libraries: OpenPyXl, PyInstaller (for conversion into "exe" format).
Files zap.xlsx, data.xlsx I can not download, because they contain trade secrets.
To use please follow these steps:
	1) download or clone the project;
	2) Create a virtual environment;
	3) use the command pip freeze > requirements.txt