# OPC2XLS Simple OPC UA client for save PLC tag to EXCEL.

Для отработки предоставления моточасов из пром сети через OPC UA в офисную сеть - написана консольная утилита opc2xls.exe
Программа подключается к указанному OPC серверу и сохраняет значения всех его тегов в файл opc2xls.xlsx

### запуск без параметров
Подключится к opc.tcp://10.100.59.1:4861 и сохранит все теги сервера в файл opc2xls.xlsx расположенный в той же папке.
```
opc2xls.exe
```
### фильтрация тегов
Cохранит только теги моточасов:
```
opc2xls.exe -filter=_MH
```
Сохранит теги в имени которых есть подстрока BC15:
opc2xls.exe -filter=BC15

### End Point URL OPC UA сервера:
```
opc2xls.exe -ep_url="opc.tcp://10.100.59.1:4861"
```
### End Point URL OPC UA сервера:
```
opc2xls.exe -ep_url="opc.tcp://10.100.59.1:4861"
```


### имя файла для сохранения:
```
opc2xls.exe  -file_name="motor_h.xlsx"
```

### вызов справки 
```bash
opc2xls.exe -h
```
