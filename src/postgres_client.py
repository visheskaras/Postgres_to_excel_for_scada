import psycopg2
import pandas as pd
from typing import Optional, Dict, Any, List
from dataclasses import dataclass
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@dataclass
class PostgreSQLConfig:
    """Конфигурация подключения к PostgreSQL"""
    host: str
    database: str
    user: str
    password: str
    port: int = 5432
    schema: str = "public"

class PostgreSQLClient:
    """Класс для работы с PostgreSQL базой данных"""
    
    def __init__(self, config: PostgreSQLConfig):
        self.config = config
        self.connection: Optional[psycopg2.extensions.connection] = None
        self.is_connected = False
    
    def connect(self) -> bool:
        """Установка подключения к базе данных"""
        try:
            self.connection = psycopg2.connect(
                host=self.config.host,
                port=self.config.port,
                database=self.config.database,
                user=self.config.user,
                password=self.config.password
            )
            self.is_connected = True
            logger.info("Успешное подключение к PostgreSQL")
            return True
            
        except psycopg2.Error as e:
            logger.error(f"Ошибка подключения к PostgreSQL: {e}")
            self.is_connected = False
            return False
    
    def disconnect(self):
        """Закрытие подключения к базе данных"""
        if self.connection and not self.connection.closed:
            self.connection.close()
            self.is_connected = False
            logger.info("Подключение к PostgreSQL закрыто")
    
    def test_connection(self) -> bool:
        """Проверка подключения к базе данных"""
        if not self.is_connected:
            return self.connect()
        return True
    
    def execute_query(self, query: str, params: Optional[tuple] = None) -> Optional[pd.DataFrame]:
        """
        Выполнение SQL запроса и возврат результата в виде DataFrame
        
        Args:
            query: SQL запрос
            params: Параметры для запроса
            
        Returns:
            DataFrame с результатами или None в случае ошибки
        """
        if not self.test_connection():
            return None
        
        try:
            df = pd.read_sql_query(query, self.connection, params=params)
            logger.info(f"Запрос выполнен успешно. Получено {len(df)} строк")
            return df
            
        except psycopg2.Error as e:
            logger.error(f"Ошибка выполнения запроса: {e}")
            return None
        except Exception as e:
            logger.error(f"Неожиданная ошибка: {e}")
            return None
    
    def get_view_data(self, view_name: str, schema: Optional[str] = None) -> Optional[pd.DataFrame]:
        """
        Получение данных из view
        
        Args:
            view_name: Имя view
            schema: Схема (если не указана, используется из конфига)
            
        Returns:
            DataFrame с данными view или None в случае ошибки
        """
        schema_to_use = schema or self.config.schema
        query = f'SELECT * FROM "{schema_to_use}"."{view_name}"'
        
        logger.info(f"Получение данных из view: {schema_to_use}.{view_name}")
        return self.execute_query(query)
    
    def get_view_columns(self, view_name: str, schema: Optional[str] = None) -> Optional[List[str]]:
        """
        Получение списка колонок view
        
        Args:
            view_name: Имя view
            schema: Схема
            
        Returns:
            Список имен колонок или None в случае ошибки
        """
        schema_to_use = schema or self.config.schema
        query = """
            SELECT column_name 
            FROM information_schema.columns 
            WHERE table_schema = %s AND table_name = %s 
            ORDER BY ordinal_position
        """
        
        result = self.execute_query(query, (schema_to_use, view_name))
        if result is not None:
            return result['column_name'].tolist()
        return None
    
    def get_available_views(self, schema: Optional[str] = None) -> Optional[List[str]]:
        """
        Получение списка доступных view в схеме
        
        Args:
            schema: Схема для поиска
            
        Returns:
            Список имен view или None в случае ошибки
        """
        schema_to_use = schema or self.config.schema
        query = """
            SELECT table_name 
            FROM information_schema.views 
            WHERE table_schema = %s 
            ORDER BY table_name
        """
        
        result = self.execute_query(query, (schema_to_use,))
        if result is not None:
            return result['table_name'].tolist()
        return None
    
    def __enter__(self):
        """Контекстный менеджер для использования with"""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Контекстный менеджер - автоматическое закрытие подключения"""
        self.disconnect()