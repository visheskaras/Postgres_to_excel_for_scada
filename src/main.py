import flet as ft
from postgres_client import PostgreSQLClient, PostgreSQLConfig
from excel_exporter import ExcelExporter, ExcelExportConfig
from config_loader import ConfigLoader
import pandas as pd
import os
from datetime import datetime
from typing import List

class PostgreSQLExporterApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.config_loader = ConfigLoader()
        self.config = None
        self.available_views: List[str] = []
        
        # Refs
        self.db_host = ft.Ref[ft.TextField]()
        self.db_port = ft.Ref[ft.TextField]()
        self.db_name = ft.Ref[ft.TextField]()
        self.db_user = ft.Ref[ft.TextField]()
        self.db_password = ft.Ref[ft.TextField]()
        self.templates_folder = ft.Ref[ft.TextField]()
        self.output_folder = ft.Ref[ft.TextField]()
        self.view_dropdown = ft.Ref[ft.Dropdown]()
        self.status_text = ft.Ref[ft.Text]()
        self.progress_ring = ft.Ref[ft.ProgressRing]()
        self.export_button = ft.Ref[ft.ElevatedButton]()
        
        self.file_picker_templates = ft.FilePicker(on_result=self.pick_templates_folder_result)
        self.file_picker_output = ft.FilePicker(on_result=self.pick_output_folder_result)
        self.page.overlay.extend([self.file_picker_templates, self.file_picker_output])
        
        self.setup_ui()
    
    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        self.page.title = "PostgreSQL View Exporter"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 20
        self.page.scroll = ft.ScrollMode.AUTO
        
               
        self.page.add(
            ft.Text("PostgreSQL View → Excel Экспортер", 
                   size=24, weight=ft.FontWeight.BOLD),
            
            ft.Divider(),
            
            ft.Text("Настройки подключения к PostgreSQL", 
                   size=18, weight=ft.FontWeight.BOLD),
            
            ft.Row([
                ft.TextField(ref=self.db_host, label="Хост", width=200),
                ft.TextField(ref=self.db_port, label="Порт", width=100),
            ]),
            
            ft.Row([
                ft.TextField(ref=self.db_name, label="База данных", width=200),
                ft.TextField(ref=self.db_user, label="Пользователь", width=200),
            ]),
            
            ft.TextField(ref=self.db_password, label="Пароль", 
                        password=True, can_reveal_password=True, width=400),
            
            ft.Divider(),
            
            ft.Text("Настройки путей", size=18, weight=ft.FontWeight.BOLD),
            
            ft.Row([
                ft.TextField(ref=self.templates_folder, label="Папка с шаблонами", 
                            width=300, read_only=True),
                ft.ElevatedButton(
                    "Выбрать",
                    on_click=lambda _: self.file_picker_templates.get_directory_path(
                        dialog_title="Выберите папку с шаблонами"
                    )
                )
            ]),
            
            ft.Row([
                ft.TextField(ref=self.output_folder, label="Папка для результатов", 
                            width=300, read_only=True),
                ft.ElevatedButton(
                    "Выбрать",
                    on_click=lambda _: self.file_picker_output.get_directory_path(
                        dialog_title="Выберите папку для результатов"
                    )
                )
            ]),
            
            ft.Divider(),
            
            ft.Text("Выбор View для экспорта", size=18, weight=ft.FontWeight.BOLD),
            
            ft.Dropdown(
                ref=self.view_dropdown,
                label="Выберите View",
                width=400,
                # options=self.get_view_options(),
                on_change=self.on_view_selected
            ),
            
            ft.Divider(),
            
            ft.ElevatedButton(
                ref=self.export_button,
                text="🚀 Экспортировать выбранный View",
                on_click=self.export_data,
                style=ft.ButtonStyle(
                    bgcolor=ft.Colors.BLUE,
                    color=ft.Colors.WHITE,
                    padding=20
                ),
                width=400,
                disabled=True
            ),
            
            ft.Divider(),
            
            ft.Row([
                ft.ProgressRing(ref=self.progress_ring, visible=False, width=20, height=20),
                ft.Text(ref=self.status_text, size=16, selectable=True, 
                       text_align=ft.TextAlign.CENTER, expand=True),
            ], alignment=ft.MainAxisAlignment.CENTER),
            
            ft.Text("Конфигурация view и шаблонов задается в файле .env", 
                   size=12, color=ft.Colors.GREY)
        )
        # Загрузка конфигурации
        self.load_configuration()
        self.view_dropdown.current.options =self.get_view_options()
        self.page.update()
    
    def load_configuration(self):
        """Загрузка конфигурации из .env файла"""
        try:
            self.config = self.config_loader.load_config()
            self.available_views = self.config_loader.get_available_views()
            
            # Установка значений по умолчанию
            self.db_host.current.value = self.config.db_host
            self.db_port.current.value = self.config.db_port
            self.db_name.current.value = self.config.db_name
            self.db_user.current.value = self.config.db_user
            self.db_password.current.value = self.config.db_password
            self.templates_folder.current.value = self.config.templates_folder
            self.output_folder.current.value = self.config.output_folder
            
            self.update_status("✅ Конфигурация загружена из .env файла", ft.Colors.GREEN)
            
        except Exception as e:
            self.update_status(f"❌ Ошибка загрузки конфигурации: {e}", ft.Colors.RED)
    
    def get_view_options(self) -> List[ft.dropdown.Option]:
        """Получение опций для dropdown с view"""
        options = []
        for view_name in self.available_views:
            view_config = self.config_loader.get_view_config(view_name)
            if view_config:
                options.append(ft.dropdown.Option(
                    key=view_name,
                    text=f"{view_name} → {view_config.template_name}"
                ))
        return options
    
    def on_view_selected(self, e):
        """Обработчик выбора view"""
        if self.view_dropdown.current.value:
            self.export_button.current.disabled = False
        else:
            self.export_button.current.disabled = True
        self.page.update()
    
    def pick_templates_folder_result(self, e: ft.FilePickerResultEvent):
        """Обработчик выбора папки с шаблонами"""
        if e.path:
            self.templates_folder.current.value = e.path
            self.templates_folder.current.update()
    
    def pick_output_folder_result(self, e: ft.FilePickerResultEvent):
        """Обработчик выбора папки для результатов"""
        if e.path:
            self.output_folder.current.value = e.path
            self.output_folder.current.update()
    
    def update_status(self, message: str, color: str = ft.Colors.BLACK, show_progress: bool = False):
        """Обновление статуса"""
        self.status_text.current.value = message
        self.status_text.current.color = color
        self.progress_ring.current.visible = show_progress
        self.page.update()
    
    def export_data(self, e):
        """Экспорт данных"""
        selected_view = self.view_dropdown.current.value
        if not selected_view:
            self.update_status("❌ Выберите View для экспорта", ft.Colors.RED)
            return
        
        # Получение текущих значений
        host = self.db_host.current.value.strip()
        port = self.db_port.current.value.strip()
        database = self.db_name.current.value.strip()
        username = self.db_user.current.value.strip()
        password = self.db_password.current.value.strip()
        templates_folder = self.templates_folder.current.value.strip()
        output_folder = self.output_folder.current.value.strip()
        
        # Валидация
        if not all([host, database, username, password]):
            self.update_status("❌ Заполните все поля подключения к БД", ft.Colors.RED)
            return
        
        if not all([templates_folder, output_folder]):
            self.update_status("❌ Укажите папки для шаблонов и результатов", ft.Colors.RED)
            return
        
        try:
            self.update_status(f"⏳ Подготовка к экспорту {selected_view}...", ft.Colors.ORANGE, True)
            
            # Получение конфигурации view
            view_config = self.config_loader.get_view_config(selected_view)
            if not view_config:
                self.update_status(f"❌ Конфигурация для {selected_view} не найдена", ft.Colors.RED)
                return
            
            # Проверка существования шаблона
            template_path = os.path.join(templates_folder, view_config.template_name)
            if not os.path.exists(template_path):
                self.update_status(f"❌ Шаблон не найден: {view_config.template_name}", ft.Colors.RED)
                return
            
            # Генерация имени выходного файла
            output_filename = self.config_loader.generate_output_filename(selected_view)
            output_path = os.path.join(output_folder, output_filename)
            
            # Подключение к БД и экспорт
            db_config = PostgreSQLConfig(
                host=host,
                database=database,
                user=username,
                password=password,
                port=int(port) if port else 5432
            )
            
            with PostgreSQLClient(db_config) as db_client:
                self.update_status(f"⏳ Загрузка данных из {selected_view}...", ft.Colors.ORANGE, True)
                
                df = db_client.get_view_data(selected_view)
                if df is None or df.empty:
                    self.update_status(f"⚠️ {selected_view} не содержит данных", ft.Colors.ORANGE)
                    return
                
                self.update_status(f"⏳ Экспорт в {output_filename}...", ft.Colors.ORANGE, True)
                
                # Экспорт в Excel
                export_config = ExcelExportConfig(
                    template_path=template_path,
                    output_path=output_path,
                    sheet_name="Data",
                    start_row=self.config.default_start_row,
                    auto_adjust_columns=self.config.auto_adjust_columns,
                    preserve_formatting=self.config.preserve_formatting
                )
                
                exporter = ExcelExporter(export_config)
                result = exporter.export_dataframe_to_template(df, clear_existing=True, include_headers=False)
                
                if result["success"]:
                    self.update_status(
                        f"✅ {selected_view} успешно экспортирован!\n"
                        f"Файл: {output_filename}\n"
                        f"Записей: {result['records_count']}",
                        ft.Colors.GREEN
                    )
                else:
                    self.update_status(f"❌ Ошибка экспорта: {result['message']}", ft.Colors.RED)
                    
        except Exception as e:
            self.update_status(f"❌ Ошибка: {str(e)}", ft.Colors.RED)

def main(page: ft.Page):
    app = PostgreSQLExporterApp(page)

if __name__ == "__main__":
    ft.app(target=main)