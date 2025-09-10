import os
import re
from dataclasses import dataclass
from typing import Dict, List, Optional
from dotenv import load_dotenv
from datetime import datetime

@dataclass
class ViewConfig:
    view_name: str
    template_name: str
    output_pattern: str

@dataclass
class AppConfig:
    db_host: str
    db_port: str
    db_name: str
    db_user: str
    db_password: str
    templates_folder: str
    output_folder: str
    views_config: Dict[str, ViewConfig]
    default_start_row: int
    auto_adjust_columns: bool
    preserve_formatting: bool

class ConfigLoader:
    def __init__(self, env_path: str = "view_export.env"):
        self.env_path = env_path
        self.config = None
    
    def load_config(self) -> AppConfig:
        """Загрузка конфигурации из .env файла"""
        load_dotenv(self.env_path)
        
        # Базовые настройки
        db_host = os.getenv("DB_HOST", "localhost")
        db_port = os.getenv("DB_PORT", "5432")
        db_name = os.getenv("DB_NAME", "")
        db_user = os.getenv("DB_USER", "")
        db_password = os.getenv("DB_PASSWORD", "")
        
        templates_folder = os.getenv("TEMPLATES_FOLDER", "./templates")

        output_folder = os.getenv("OUTPUT_FOLDER", "./output")
        
        # Параметры экспорта
        default_start_row = int(os.getenv("DEFAULT_START_ROW", "2"))
        auto_adjust_columns = os.getenv("AUTO_ADJUST_COLUMNS", "true").lower() == "true"
        preserve_formatting = os.getenv("PRESERVE_FORMATTING", "true").lower() == "true"
        
        # Загрузка конфигурации view
        views_config = self._load_views_config()
        
        # Создание папок если не существуют
        os.makedirs(templates_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)
        
        self.config = AppConfig(
            db_host=db_host,
            db_port=db_port,
            db_name=db_name,
            db_user=db_user,
            db_password=db_password,
            templates_folder=templates_folder,
            output_folder=output_folder,
            views_config=views_config,
            default_start_row=default_start_row,
            auto_adjust_columns=auto_adjust_columns,
            preserve_formatting=preserve_formatting
        )
        
        return self.config
    
    def _load_views_config(self) -> Dict[str, ViewConfig]:
        """Загрузка конфигурации view из переменных окружения"""
        views_config = {}
        
        # Читаем напрямую из .env файла, чтобы избежать системных переменных
        try:
            with open(self.env_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    # Пропускаем пустые строки и комментарии
                    if not line or line.startswith('#'):
                        continue
                    
                    # Разбираем строку на ключ и значение
                    if '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        # Пропускаем служебные переменные
                        if key.startswith(("DB_", "TEMPLATES_", "OUTPUT_", "DEFAULT_", "AUTO_", "PRESERVE_")):
                            continue
                        
                        # Проверяем, что значение содержит разделитель : и выглядит как конфигурация view
                        if self._is_valid_view_config(value):
                            try:
                                template_name, output_pattern = value.split(':', 1)
                                template_name = template_name.strip()
                                output_pattern = output_pattern.strip()
                                
                                # Дополнительная проверка на валидность
                                if template_name and output_pattern:
                                    view_config = ViewConfig(
                                        view_name=key,
                                        template_name=template_name,
                                        output_pattern=output_pattern
                                    )
                                    views_config[key] = view_config
                                    print(f"✅ Загружена конфигурация для view: {key} = {template_name}:{output_pattern}")
                            except ValueError:
                                print(f"⚠️ Неверный формат конфигурации для {key}: {value}")
        except FileNotFoundError:
            print(f"⚠️ Файл {self.env_path} не найден")
        except Exception as e:
            print(f"⚠️ Ошибка чтения файла {self.env_path}: {e}")
        
        if not views_config:
            print("⚠️ Не найдено ни одной конфигурации view в .env файле")
            print("ℹ️ Пример формата: SALES_REPORT=sales_template.xlsx:sales_report_{date}.xlsx")
        
        return views_config
    
    def _is_valid_view_config(self, value: str) -> bool:
        """Проверяет, является ли значение валидной конфигурацией view"""
        if ':' not in value:
            return False
        
        # Проверяем, что значение не похоже на системный путь
        if any(char in value for char in ['/', '\\', '$', '%']):
            return False
        
        # Проверяем, что значение содержит расширение .xlsx
        if '.xlsx' not in value:
            return False
        
        # Проверяем, что после разделителя есть какой-то паттерн
        parts = value.split(':', 1)
        if len(parts) != 2:
            return False
        
        template_part, output_part = parts
        return bool(template_part.strip() and output_part.strip())    
    
    def get_available_views(self) -> List[str]:
        """Получение списка доступных view"""
        if not self.config:
            self.load_config()
        return list(self.config.views_config.keys())
    
    def get_view_config(self, view_name: str) -> Optional[ViewConfig]:
        """Получение конфигурации для конкретного view"""
        if not self.config:
            self.load_config()
        return self.config.views_config.get(view_name)
    
    def generate_output_filename(self, view_name: str) -> str:
        """Генерация имени выходного файла на основе паттерна"""
        config = self.get_view_config(view_name)
        if not config:
            raise ValueError(f"Конфигурация для view '{view_name}' не найдена")
        
        # Замена плейсхолдеров
        filename = config.output_pattern
        filename = filename.replace('{date}', datetime.now().strftime('%Y-%m-%d'))
        filename = filename.replace('{timestamp}', datetime.now().strftime('%Y%m%d_%H%M%S'))
        filename = filename.replace('{time}', datetime.now().strftime('%H%M%S'))
        
        return filename