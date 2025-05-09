# Установка зависимостей
pip install pyserial tqdm

# Найти устройство по имени ("Arduino")
python com_scanner.py --name "Arduino" --command "GET_FILE" --output "data.bin"

# По производителю или ID устройства
python com_scanner.py --vid 0403 --pid 6001 --command "DOWNLOAD" --output "data.bin"

# Сканировать и показать все доступные COM-порты и устройства
python com_scanner.py --scan-all

# Вариант команды
cd ComScan; python com_scanner.py --scan-all