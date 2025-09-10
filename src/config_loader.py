import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from dotenv import load_dotenv
from datetime import datetime

@dataclass
class ViewConfig:
    view_name: str
    template_name: str
    output_pattern: str
    start_row: int
    start_col: int

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
    default_start_col: int
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
        db_name = os.getenv("DB_NAME", "test_db")
        db_user = os.getenv("DB_USER", "postgres")
        db_password = os.getenv("DB_PASSWORD", "")
        
        templates_folder = os.getenv("TEMPLATES_FOLDER", "./templates")
        output_folder = os.getenv("OUTPUT_FOLDER", "./output")
        
        # Параметры экспорта
        default_start_row = int(os.getenv("DEFAULT_START_ROW", "2"))
        default_start_col = int(os.getenv("DEFAULT_START_COL", "1"))
        auto_adjust_columns = os.getenv("AUTO_ADJUST_COLUMNS", "true").lower() == "true"
        preserve_formatting = os.getenv("PRESERVE_FORMATTING", "true").lower() == "true"
        
        # Загрузка конфигурации view
        views_config = self._load_views_config(default_start_row, default_start_col)
        
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
            default_start_col=default_start_col,
            auto_adjust_columns=auto_adjust_columns,
            preserve_formatting=preserve_formatting
        )
        
        return self.config
    
    def _parse_position(self, position_str: str, default_row: int, default_col: int) -> Tuple[int, int]:
        """Парсит строку позиции в формате 'row,col' или 'cell'"""
        if not position_str:
            return default_row, default_col
        
        # Формат: "2,3" (строка, столбец)
        if ',' in position_str:
            try:
                row, col = position_str.split(',', 1)
                return int(row.strip()), int(col.strip())
            except:
                return default_row, default_col
        
        # Формат: "B3" (Excel-style)
        elif position_str and any(c.isalpha() for c in position_str) and any(c.isdigit() for c in position_str):
            try:
                from openpyxl.utils import coordinate_to_tuple
                row, col = coordinate_to_tuple(position_str.upper())
                return row, col
            except:
                return default_row, default_col
        
        return default_row, default_col
    
    def _load_views_config(self, default_row: int, default_col: int) -> Dict[str, ViewConfig]:
        """Загрузка конфигурации view из .env файла"""
        views_config = {}
        
        # Читаем напрямую из .env файла, чтобы избежать системных переменных
        try:
            if not os.path.exists(self.env_path):
                print(f"⚠️ Файл {self.env_path} не найден")
                return views_config
                
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
                        if key.startswith(('DB_', 'TEMPLATES_', 'OUTPUT_', 'DEFAULT_', 'AUTO_', 'PRESERVE_')):
                            continue
                        
                        # Пропускаем системные переменные (все заглавные)
                        if key.isupper() and len(key) > 3:
                            continue
                        
                        # Проверяем, что значение содержит разделитель : и выглядит как конфигурация view
                        if value and ':' in value and '.xlsx' in value:
                            try:
                                # Форматы:
                                # 1. template.xlsx:output_pattern
                                # 2. template.xlsx:output_pattern:start_row,start_col
                                # 3. template.xlsx:output_pattern:B3 (Excel-style)
                                
                                parts = value.split(':')
                                template_name = parts[0].strip()
                                output_pattern = parts[1].strip()
                                
                                # Проверяем, что template_name заканчивается на .xlsx
                                if not template_name.lower().endswith('.xlsx'):
                                    continue
                                
                                # Парсим позицию (если указана)
                                start_row, start_col = default_row, default_col
                                if len(parts) >= 3:
                                    start_row, start_col = self._parse_position(parts[2].strip(), default_row, default_col)
                                
                                if template_name and output_pattern:
                                    views_config[key] = ViewConfig(
                                        view_name=key,
                                        template_name=template_name,
                                        output_pattern=output_pattern,
                                        start_row=start_row,
                                        start_col=start_col
                                    )
                                    print(f"✅ Загружена конфигурация: {key} = {template_name}:{output_pattern} (позиция: {start_row},{start_col})")
                            except Exception as e:
                                print(f"⚠️ Ошибка парсинга конфигурации {key}: {value} - {e}")
                                continue
        except Exception as e:
            print(f"⚠️ Ошибка чтения файла {self.env_path}: {e}")
        
        print(f"📊 Всего загружено view: {len(views_config)}")
        return views_config
    
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
            return f"{view_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        filename = config.output_pattern
        filename = filename.replace('{date}', datetime.now().strftime('%Y-%m-%d'))
        filename = filename.replace('{timestamp}', datetime.now().strftime('%Y%m%d_%H%M%S'))
        return filename