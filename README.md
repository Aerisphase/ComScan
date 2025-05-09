### Установка зависимостей
pip install pyserial tqdm

### Найти устройство по имени ("Arduino")
python com_scanner.py --name "Arduino" --command "GET_FILE" --output "results.bin"

### По производителю или ID устройства
python com_scanner.py --vid 0403 --pid 6001 --command "GET_FILE" --output "results.bin"

### Сканировать и показать все доступные COM-порты и устройства
python com_scanner.py --scan-all

### Вариант команды
cd ComScan; python com_scanner.py --scan-all
