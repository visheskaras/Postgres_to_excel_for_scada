import flet as ft
import os
import threading
from config_loader import ConfigLoader
from postgres_client import PostgreSQLClient, PostgreSQLConfig
from excel_exporter import ExcelExporter, ExcelExportConfig
import pandas as pd

def main(page: ft.Page):
    page.title = "PostgreSQL View Exporter"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 20
    page.scroll = ft.ScrollMode.AUTO
    
    # Загрузка конфигурации
    config_loader = ConfigLoader()
    config = config_loader.load_config()
    available_views = config_loader.get_available_views()
    
    # Элементы UI
    db_host = ft.TextField(label="Хост", value=config.db_host, width=200)
    db_port = ft.TextField(label="Порт", value=config.db_port, width=100)
    db_name = ft.TextField(label="База данных", value=config.db_name, width=200)
    db_user = ft.TextField(label="Пользователь", value=config.db_user, width=200)
    db_password = ft.TextField(label="Пароль", value=config.db_password,
                              password=True, can_reveal_password=True, width=400)
    
    templates_folder = ft.TextField(label="Папка с шаблонами", 
                                   value=config.templates_folder, width=300, read_only=True)
    output_folder = ft.TextField(label="Папка для результатов", 
                                value=config.output_folder, width=300, read_only=True)
    
    # Checkbox для выбора всех
    select_all_checkbox = ft.Checkbox(
        label="Выбрать все",
        value=False,
        on_change=lambda e: toggle_select_all(e)
    )
    
    # Checkbox для каждого view с отображением позиции
    view_checkboxes = []
    for view in available_views:
        view_config = config_loader.get_view_config(view)
        position_info = f" ({view_config.start_row},{view_config.start_col})" if view_config else ""
        view_checkboxes.append(
            ft.Checkbox(
                label=f"{view}{position_info}",
                value=False,
                data=view
            )
        )
    
    status_text = ft.Text("Готово к работе", size=16, color=ft.Colors.GREEN)
    progress_ring = ft.ProgressRing(visible=False, width=20, height=20)
    progress_bar = ft.ProgressBar(visible=False, width=400)
    
    # File pickers
    def pick_templates_result(e: ft.FilePickerResultEvent):
        if e.path:
            templates_folder.value = e.path
            templates_folder.update()
    
    def pick_output_result(e: ft.FilePickerResultEvent):
        if e.path:
            output_folder.value = e.path
            output_folder.update()
    
    file_picker_templates = ft.FilePicker(on_result=pick_templates_result)
    file_picker_output = ft.FilePicker(on_result=pick_output_result)
    page.overlay.extend([file_picker_templates, file_picker_output])
    
    # Функция для выбора/снятия всех
    def toggle_select_all(e):
        for checkbox in view_checkboxes:
            checkbox.value = select_all_checkbox.value
        page.update()
    
    # Функция получения выбранных view
    def get_selected_views():
        return [checkbox.data for checkbox in view_checkboxes if checkbox.value]
    
    # Функция экспорта одного view
    def export_single_view(view_name, templates_path, output_path):
        try:
            # Подключение к БД
            db_config = PostgreSQLConfig(
                host=db_host.value.strip(),
                database=db_name.value.strip(),
                user=db_user.value.strip(),
                password=db_password.value.strip(),
                port=int(db_port.value.strip()) if db_port.value.strip() else 5432
            )
            
            with PostgreSQLClient(db_config) as db_client:
                # Получение данных
                df = db_client.get_view_data(view_name)
                if df is None or df.empty:
                    return f"❌ {view_name}: нет данных"
                
                # Конфигурация view
                view_config = config_loader.get_view_config(view_name)
                if not view_config:
                    return f"❌ {view_name}: конфигурация не найдена"
                
                # Проверка шаблона
                template_path = os.path.join(templates_path, view_config.template_name)
                if not os.path.exists(template_path):
                    return f"❌ {view_name}: шаблон не найден"
                
                # Генерация имени файла
                output_filename = config_loader.generate_output_filename(view_name)
                output_filepath = os.path.join(output_path, output_filename)
                
                # Экспорт в Excel с указанной позицией
                export_config = ExcelExportConfig(
                    template_path=template_path,
                    output_path=output_filepath,
                    sheet_name="Data",
                    start_row=view_config.start_row,
                    start_col=view_config.start_col,
                    auto_adjust_columns=config.auto_adjust_columns,
                    preserve_formatting=config.preserve_formatting
                )
                
                exporter = ExcelExporter(export_config)
                result = exporter.export_dataframe_to_template(df, clear_existing=True, include_headers=False)
                
                if result["success"]:
                    position_info = f" (позиция: {view_config.start_row},{view_config.start_col})"
                    return f"✅ {view_name}: успешно {position_info} ({result['records_count']} записей)"
                else:
                    return f"❌ {view_name}: ошибка экспорта"
                    
        except Exception as e:
            return f"❌ {view_name}: {str(e)}"
    
    # Функции для обновления UI
    def update_ui_status(message, color, show_progress=False, progress_value=None):
        status_text.value = message
        status_text.color = color
        progress_ring.visible = show_progress
        progress_bar.visible = show_progress
        if progress_value is not None:
            progress_bar.value = progress_value
        export_button.disabled = show_progress
        page.update()
    
    # Функция экспорта в отдельном потоке
    def export_in_thread():
        selected_views = get_selected_views()
        if not selected_views:
            update_ui_status("❌ Выберите хотя бы один View для экспорта", ft.Colors.RED)
            return
        
        templates_path = templates_folder.value.strip()
        output_path = output_folder.value.strip()
        
        if not all([templates_path, output_path]):
            update_ui_status("❌ Укажите папки для шаблонов и результатов", ft.Colors.RED)
            return
        
        # Настройка UI для процесса экспорта
        update_ui_status(f"⏳ Начинаем экспорт {len(selected_views)} view...", ft.Colors.ORANGE, True, 0)
        
        results = []
        total = len(selected_views)
        
        # Экспорт каждого view
        for i, view_name in enumerate(selected_views):
            progress = (i + 1) / total
            update_ui_status(f"⏳ Экспортируем {view_name} ({i+1}/{total})...", ft.Colors.ORANGE, True, progress)
            
            result = export_single_view(view_name, templates_path, output_path)
            results.append(result)
        
        # Завершение
        success_count = sum(1 for r in results if r.startswith("✅"))
        error_count = sum(1 for r in results if r.startswith("❌"))
        
        result_text = f"✅ Готово! Успешно: {success_count}, Ошибок: {error_count}"
        color = ft.Colors.GREEN if error_count == 0 else ft.Colors.ORANGE
        
        update_ui_status(result_text, color, False, 1)
        show_results_dialog("\n".join(results))
    
    # Диалог с результатами
    def show_results_dialog(results_text):
        def close_dialog(e):
            results_dialog.open = False
            page.update()
        
        results_dialog = ft.AlertDialog(
            title=ft.Text("Результаты экспорта"),
            content=ft.Column(
                [
                    ft.Text("Детали выполнения:", weight=ft.FontWeight.BOLD),
                    ft.Text(results_text, selectable=True, size=14),
                ],
                tight=True,
                scroll=ft.ScrollMode.AUTO,
                height=300
            ),
            actions=[
                ft.TextButton("Закрыть", on_click=close_dialog)
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )
        
        page.dialog = results_dialog
        results_dialog.open = True
        page.update()
    
    # Основная функция экспорта
    def start_export(e):
        # Запускаем экспорт в отдельном потоке, чтобы не блокировать UI
        thread = threading.Thread(target=export_in_thread, daemon=True)
        thread.start()
    
    # Кнопка экспорта
    export_button = ft.ElevatedButton(
        "🚀 Экспортировать выбранные",
        on_click=start_export,
        style=ft.ButtonStyle(
            bgcolor=ft.Colors.BLUE,
            color=ft.Colors.WHITE,
            padding=20
        ),
        width=400
    )
    
    # Сборка интерфейса
    page.add(
        ft.Text("PostgreSQL View → Excel Экспортер", 
               size=24, weight=ft.FontWeight.BOLD),
        
        ft.Divider(),
        
        ft.Text("Настройки подключения", size=18, weight=ft.FontWeight.BOLD),
        
        ft.Row([db_host, db_port]),
        ft.Row([db_name, db_user]),
        db_password,
        
        ft.Divider(),
        
        ft.Text("Настройки путей", size=18, weight=ft.FontWeight.BOLD),
        
        ft.Row([
            templates_folder,
            ft.ElevatedButton(
                "Выбрать",
                on_click=lambda _: file_picker_templates.get_directory_path()
            )
        ]),
        
        ft.Row([
            output_folder,
            ft.ElevatedButton(
                "Выбрать",
                on_click=lambda _: file_picker_output.get_directory_path()
            )
        ]),
        
        ft.Divider(),
        
        ft.Text("Выбор View для экспорта", size=18, weight=ft.FontWeight.BOLD),
        
        ft.Row([
            select_all_checkbox,
            ft.Text(f"Доступно: {len(available_views)} view", color=ft.Colors.GREY)
        ]),
        
        ft.Column(
            view_checkboxes,
            scroll=ft.ScrollMode.AUTO,
            height=200,
            width=400
        ),
        
        ft.Divider(),
        
        export_button,
        
        ft.Divider(),
        
        ft.Column([
            ft.Row([progress_ring, status_text], alignment=ft.MainAxisAlignment.CENTER),
            progress_bar
        ]),
        
        ft.Text("В скобках указана стартовая позиция (строка,столбец) для вставки данных", 
               size=12, color=ft.Colors.GREY)
    )

if __name__ == "__main__":
    ft.app(target=main)