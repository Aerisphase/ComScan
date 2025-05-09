#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
COM port scanner лабораторная для Тян
"""
##Комменты специально на русском обычно (как и в проекте курсовой/дипломной) пишу на английском

# Исправление проблем с кодировкой в консоли Windows
import sys
import os

# Попытка установить режим UTF-8 для консоли Windows
try:
    if sys.platform == 'win32':
        os.system('chcp 65001 > nul')
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
except Exception:
    pass  # Если не получилось, продолжаем с кодировкой по умолчанию

import serial
import serial.tools.list_ports
import time
import os
import re
import argparse
import logging
from tqdm import tqdm

# Логирование, смотрим ивенты

# Функция для логирования с поддержкой кириллицы (9ое мая все таки)
def safe_log(logger_func, message):
    try:
        logger_func(message)
    except UnicodeEncodeError:
        # Если произошла ошибка кодировки, просто выводим в консоль
        try:
            print(message)
        except:
            print("[Log message with encoding issues]")

# Настройка логгера только для файла
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("com_scanner.log", encoding='utf-8'),
    ]
)
logger = logging.getLogger(__name__)

class ComScanner:
    def __init__(self, vendor_id=None, product_id=None, device_name=None, baud_rate=9600, timeout=1):
        """
        Инициализация сканера COM-портов.
        
        Args:
            vendor_id (str, опционально): ID производителя целевого устройства в шестнадцатеричном формате (например, '0403')
            product_id (str, опционально): ID продукта целевого устройства в шестнадцатеричном формате (например, '6001')
            device_name (str, опционально): Имя или часть имени целевого устройства
            baud_rate (int, опционально): Скорость передачи данных для последовательной связи
            timeout (int, опционально): Тайм-аут для операций с последовательным портом в секундах
        """
        self.vendor_id = vendor_id
        self.product_id = product_id
        self.device_name = device_name
        self.baud_rate = baud_rate
        self.timeout = timeout
        self.connected_port = None
        self.serial_conn = None
    
    def scan_ports(self):
        """
        Сканирование доступных COM-портов и возврат списка информации о портах.
        
        Returns:
            list: Список доступных COM-портов с их подробной информацией
        """
        safe_log(logger.info, "Сканирование доступных COM-портов...")
        ports = list(serial.tools.list_ports.comports())
        
        if not ports:
            safe_log(logger.warning, "COM-порты не найдены!")
            return []
        
        safe_log(logger.info, f"Найдено {len(ports)} COM-порт(ов)")
        for port in ports:
            safe_log(logger.info, f"Порт: {port.device}, Описание: {port.description}, "
                      f"Аппаратный ID: {port.hwid}")
        
        return ports
    
    def find_device(self):
        """
        Поиск целевого устройства среди доступных COM-портов на основе указанных критериев.
        
        Returns:
            str: Имя COM-порта найденного устройства или None, если устройство не найдено
        """
        ports = self.scan_ports()
        
        if not ports:
            return None
        
        for port in ports:
            # Проверка, соответствует ли порт каким-либо из указанных критериев
            if self.vendor_id and self.product_id:
                # Проверка аппаратного ID на наличие VID и PID
                hwid = port.hwid.lower()
                if f"vid_{self.vendor_id.lower()}" in hwid and f"pid_{self.product_id.lower()}" in hwid:
                    safe_log(logger.info, f"Найдено устройство с VID:{self.vendor_id}, PID:{self.product_id} на порту {port.device}")
                    self.connected_port = port.device
                    return port.device
            
            if self.device_name:
                # Проверка, содержится ли имя устройства в описании
                if self.device_name.lower() in port.description.lower():
                    safe_log(logger.info, f"Найдено устройство с именем '{self.device_name}' на порту {port.device}")
                    self.connected_port = port.device
                    return port.device
        
        safe_log(logger.warning, "Целевое устройство не найдено!")
        return None
    
    def connect(self, port=None):
        """
        Подключение к указанному COM-порту.
        
        Args:
            port (str, опционально): COM-порт для подключения. Если None, использует порт найденного устройства.
            
        Returns:
            bool: True, если подключение успешно, False в противном случае
        """
        if not port and not self.connected_port:
            safe_log(logger.error, "Порт не указан и устройство не найдено!")
            return False
        
        target_port = port or self.connected_port
        
        try:
            safe_log(logger.info, f"Подключение к {target_port} со скоростью {self.baud_rate} бод...")
            self.serial_conn = serial.Serial(
                port=target_port,
                baudrate=self.baud_rate,
                timeout=self.timeout
            )
            safe_log(logger.info, f"Успешно подключено к {target_port}")
            return True
        except serial.SerialException as e:
            safe_log(logger.error, f"Не удалось подключиться к {target_port}: {str(e)}")
            return False
    
    def disconnect(self):
        """
        Отключение от текущего последовательного соединения.
        """
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
            safe_log(logger.info, f"Отключено от {self.connected_port}")
            self.serial_conn = None
    
    def send_command(self, command, wait_time=0.5):
        """
        Отправка команды подключенному устройству.
        
        Args:
            command (str): Команда для отправки
            wait_time (float, опционально): Время ожидания после отправки команды
            
        Returns:
            str: Ответ от устройства или None в случае неудачи
        """
        if not self.serial_conn or not self.serial_conn.is_open:
            logger.error("Not connected to any device!")
            return None
        
        try:
            # Убедиться, что команда заканчивается символом новой строки
            if not command.endswith('\n'):
                command += '\n'
            
            # Send command
            self.serial_conn.write(command.encode())
            safe_log(logger.debug, f"Отправлена команда: {command.strip()}")
            
            # Ожидание обработки устройством
            time.sleep(wait_time)
            
            # Чтение ответа
            response = ""
            while self.serial_conn.in_waiting:
                response += self.serial_conn.read(self.serial_conn.in_waiting).decode(errors='replace')
            
            return response
        except Exception as e:
            safe_log(logger.error, f"Ошибка отправки команды: {str(e)}")
            return None
    
    def extract_file(self, file_command, output_path, file_size=None, chunk_size=1024):
        """
        Извлечение файла из подключенного устройства.
        
        Args:
            file_command (str): Команда для запроса файла с устройства
            output_path (str): Путь, по которому будет сохранен извлеченный файл
            file_size (int, опционально): Ожидаемый размер файла в байтах для отслеживания прогресса
            chunk_size (int, опционально): Размер фрагментов для чтения за один раз
            
        Returns:
            bool: True, если извлечение файла успешно, False в противном случае
        """
        if not self.serial_conn or not self.serial_conn.is_open:
            logger.error("Not connected to any device!")
            return False
        
        try:
            # Создание выходного каталога, если он не существует
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            
            # Отправка команды для запроса файла
            safe_log(logger.info, f"Запрос файла с помощью команды: {file_command}")
            self.serial_conn.reset_input_buffer()
            response = self.send_command(file_command, wait_time=1.0)
            
            if not response:
                safe_log(logger.error, "Нет ответа от устройства после запроса файла")
                return False
            
            safe_log(logger.info, "Начало передачи файла...")
            
            # Открытие файла для записи
            with open(output_path, 'wb') as f:
                bytes_received = 0
                start_time = time.time()
                
                # Настройка индикатора прогресса, если известен размер файла
                pbar = None
                if file_size:
                    pbar = tqdm(total=file_size, unit='B', unit_scale=True, desc="Загрузка")
                
                # Чтение данных фрагментами
                while True:
                    if self.serial_conn.in_waiting:
                        chunk = self.serial_conn.read(min(chunk_size, self.serial_conn.in_waiting))
                        if not chunk:
                            break
                        
                        f.write(chunk)
                        bytes_received += len(chunk)
                        
                        if pbar:
                            pbar.update(len(chunk))
                    else:
                        # Проверка, получили ли мы ожидаемый размер файла
                        if file_size and bytes_received >= file_size:
                            break
                        
                        # Проверка на тайм-аут
                        if time.time() - start_time > 30:  # 30 секунд тайм-аут
                            if bytes_received > 0:
                                safe_log(logger.info, "Передача, похоже, завершена (достигнут тайм-аут)")
                                break
                            else:
                                safe_log(logger.error, "Тайм-аут ожидания данных")
                                return False
                        
                        time.sleep(0.1)  # Короткая задержка чтоб не сгорел у меня проц
                
                if pbar:
                    pbar.close()
            
            transfer_time = time.time() - start_time
            safe_log(logger.info, f"Передача файла завершена. Получено {bytes_received} байт за {transfer_time:.2f} секунд")
            safe_log(logger.info, f"Файл сохранен в: {output_path}")
            
            return True
        
        except Exception as e:
            safe_log(logger.error, f"Ошибка извлечения файла: {str(e)}")
            return False


def scan_all_devices():
    """
    Сканирует и выводит информацию о всех доступных COM-портах и подключенных устройствах.
    """
    safe_log(logger.info, "Сканирование всех доступных COM-портов...")
    ports = list(serial.tools.list_ports.comports())
    
    if not ports:
        safe_log(logger.warning, "COM-порты не найдены!")
        return False
    
    print("\n=== Найденные устройства на COM-портах ===")
    print("{:<10} {:<50} {:<30}".format("Порт", "Описание", "Аппаратный ID"))
    print("-" * 90)
    
    for i, port in enumerate(ports):
        print("{:<10} {:<50} {:<30}".format(
            port.device,
            port.description[:47] + "..." if len(port.description) > 50 else port.description,
            port.hwid[:27] + "..." if len(port.hwid) > 30 else port.hwid
        ))
        
        # Извлечение VID и PID из hwid, если они есть
        hwid = port.hwid.lower()
        vid_match = re.search(r"vid_([0-9a-f]{4})", hwid)
        pid_match = re.search(r"pid_([0-9a-f]{4})", hwid)
        
        if vid_match and pid_match:
            vid = vid_match.group(1)
            pid = pid_match.group(1)
            print("    └─ VID:PID = {}:{}".format(vid.upper(), pid.upper()))
        
        # Дополнительная информация, если доступна
        if hasattr(port, 'manufacturer') and port.manufacturer:
            print("    └─ Производитель: {}".format(port.manufacturer))
        if hasattr(port, 'product') and port.product:
            print("    └─ Продукт: {}".format(port.product))
        if hasattr(port, 'serial_number') and port.serial_number:
            print("    └─ Серийный номер: {}".format(port.serial_number))
        
        # Пустая строка между устройствами для лучшей читаемости
        if i < len(ports) - 1:
            print()
    
    print("\nВсего найдено {} устройств".format(len(ports)))
    return True

def main():
    """
    Основная функция для анализа аргументов и выполнения сканирования COM-порта и извлечения файла.
    """
    parser = argparse.ArgumentParser(description="Сканирование COM-портов, поиск определенного устройства и извлечение из него файла")
    
    # Параметры идентификации устройства
    device_group = parser.add_argument_group("Идентификация устройства")
    device_group.add_argument("--scan-all", action="store_true", help="Сканировать и показать все доступные COM-порты и устройства")
    device_group.add_argument("--vid", help="ID производителя целевого устройства (например, '0403')")
    device_group.add_argument("--pid", help="ID продукта целевого устройства (например, '6001')")
    device_group.add_argument("--name", help="Имя или часть имени целевого устройства")
    device_group.add_argument("--port", help="Прямое указание COM-порта для использования (например, 'COM3')")
    
    # Параметры подключения
    conn_group = parser.add_argument_group("Параметры подключения")
    conn_group.add_argument("--baud", type=int, default=9600, help="Скорость передачи данных для последовательной связи (по умолчанию: 9600)")
    conn_group.add_argument("--timeout", type=int, default=1, help="Тайм-аут для операций с последовательным портом в секундах (по умолчанию: 1)")
    
    # Параметры извлечения файла
    file_group = parser.add_argument_group("Извлечение файла")
    file_group.add_argument("--command", help="Команда для отправки запроса файла")
    file_group.add_argument("--output", help="Путь, по которому будет сохранен извлеченный файл")
    file_group.add_argument("--size", type=int, help="Ожидаемый размер файла в байтах (для отслеживания прогресса)")
    
    # Анализ аргументов
    args = parser.parse_args()
    
    # Если указан параметр --scan-all, просто сканируем и выводим все устройства
    if args.scan_all:
        return 0 if scan_all_devices() else 1
    
    # Проверка аргументов для режима извлечения файла
    if not args.command or not args.output:
        parser.error("Для извлечения файла необходимо указать параметры --command и --output")
        
    if not (args.vid and args.pid) and not args.name and not args.port:
        parser.error("Вы должны указать либо --vid и --pid, либо --name, либо --port")
    
    # Создание экземпляра сканера
    scanner = ComScanner(
        vendor_id=args.vid,
        product_id=args.pid,
        device_name=args.name,
        baud_rate=args.baud,
        timeout=args.timeout
    )
    
    try:
        # Подключение к устройству
        if args.port:
            # Прямое подключение к указанному порту
            if not scanner.connect(args.port):
                return 1
        else:
            # Поиск и подключение к устройству
            port = scanner.find_device()
            if not port:
                return 1
            
            if not scanner.connect():
                return 1
        
        # Извлечение файла
        success = scanner.extract_file(
            file_command=args.command,
            output_path=args.output,
            file_size=args.size
        )
        
        return 0 if success else 1
    
    finally:
        # Гарантированное отключение
        scanner.disconnect()


if __name__ == "__main__":
    exit(main())
