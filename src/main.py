import flet as ft
from postgres_client import PostgreSQLClient, PostgreSQLConfig
from excel_exporter import ExcelExporter, ExcelExportConfig
import pandas as pd
import os
from datetime import datetime

def main(page: ft.Page):
    page.title = "PostgreSQL View to Excel Exporter"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.vertical_alignment = ft.MainAxisAlignment.START
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.padding = 20
    page.scroll = ft.ScrollMode.AUTO

    # Refs для элементов UI
    db_host = ft.Ref[ft.TextField]()
    db_port = ft.Ref[ft.TextField]()
    db_name = ft.Ref[ft.TextField]()
    db_user = ft.Ref[ft.TextField]()
    db_password = ft.Ref[ft.TextField]()
    view_name = ft.Ref[ft.TextField]()
    template_path = ft.Ref[ft.TextField]()
    output_path = ft.Ref[ft.TextField]()
    status_text = ft.Ref[ft.Text]()
    progress_ring = ft.Ref[ft.ProgressRing]()

    # File pickers
    def pick_template_result(e: ft.FilePickerResultEvent):
        if e.files:
            template_path.current.value = e.files[0].path
            template_path.current.update()

    def pick_output_result(e: ft.FilePickerResultEvent):
        if e.files:
            output_path.current.value = e.files[0].path
            output_path.current.update()
        elif e.path:
            output_path.current.value = e.path
            output_path.current.update()

    file_picker_template = ft.FilePicker(on_result=pick_template_result)
    file_picker_output = ft.FilePicker(on_result=pick_output_result)
    page.overlay.extend([file_picker_template, file_picker_output])

    def update_status(message: str, color: str = ft.Colors.BLACK, show_progress: bool = False):
        status_text.current.value = message
        status_text.current.color = color
        progress_ring.current.visible = show_progress
        page.update()

    def export_data(e):
        # Получение значений
        host = db_host.current.value.strip()
        port = db_port.current.value.strip() or "5432"
        database = db_name.current.value.strip()
        username = db_user.current.value.strip()
        password = db_password.current.value.strip()
        view = view_name.current.value.strip()
        template = template_path.current.value.strip()
        output = output_path.current.value.strip() or f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        # Валидация
        if not all([host, database, username, password, view, template]):
            update_status("❌ Заполните все обязательные поля!", ft.Colors.RED)
            return

        if not os.path.exists(template):
            update_status("❌ Файл шаблона не найден!", ft.Colors.RED)
            return

        try:
            update_status("⏳ Подключаемся к PostgreSQL...", ft.Colors.ORANGE, True)

            # 1. Получение данных из PostgreSQL
            db_config = PostgreSQLConfig(
                host=host,
                database=database,
                user=username,
                password=password,
                port=int(port)
            )

            with PostgreSQLClient(db_config) as db_client:
                update_status("⏳ Загружаем данные из view...", ft.Colors.ORANGE, True)
                df = db_client.get_view_data(view)

                if df is None or df.empty:
                    update_status("⚠️ Не удалось получить данные или view пусто", ft.Colors.ORANGE)
                    return

                update_status("⏳ Экспортируем в Excel...", ft.Colors.ORANGE, True)

                # 2. Экспорт в Excel
                export_config = ExcelExportConfig(
                    template_path=template,
                    output_path=output,
                    sheet_name="Data",
                    start_row=2
                )

                exporter = ExcelExporter(export_config)
                result = exporter.export_dataframe_to_template(df, clear_existing=True, include_headers=False)

                if result["success"]:
                    update_status(
                        f"✅ Успешно!\nФайл: {result['output_path']}\nЗаписей: {result['records_count']}",
                        ft.Colors.GREEN
                    )
                else:
                    update_status(f"❌ {result['message']}", ft.Colors.RED)

        except Exception as e:
            update_status(f"❌ Ошибка: {str(e)}", ft.Colors.RED)

    # UI
    page.add(
        ft.Text("PostgreSQL → Excel Экспортер", size=24, weight=ft.FontWeight.BOLD),
        ft.Divider(),
        
        ft.Text("Настройки PostgreSQL", size=18, weight=ft.FontWeight.BOLD),
        ft.Row([ft.TextField(ref=db_host, label="Хост*", width=200, value="localhost"),
                ft.TextField(ref=db_port, label="Порт", width=100, value="5432")]),
        ft.Row([ft.TextField(ref=db_name, label="База данных*", width=200),
                ft.TextField(ref=db_user, label="Пользователь*", width=200)]),
        ft.Row([ft.TextField(ref=db_password, label="Пароль*", password=True, width=400)]),
        
        ft.Divider(),
        ft.Text("Параметры экспорта", size=18, weight=ft.FontWeight.BOLD),
        ft.Row([ft.TextField(ref=view_name, label="Имя View*", width=400)]),
        
        ft.Row([ft.TextField(ref=template_path, label="Шаблон Excel*", width=300, read_only=True),
                ft.ElevatedButton("Выбрать", on_click=lambda _: file_picker_template.pick_files(
                    allowed_extensions=["xlsx"], dialog_title="Выберите шаблон Excel"))]),
        
        ft.Row([ft.TextField(ref=output_path, label="Куда сохранить", width=300),
                ft.ElevatedButton("Выбрать", on_click=lambda _: file_picker_output.get_directory_path())]),
        
        ft.Divider(),
        ft.ElevatedButton("🚀 Экспортировать", on_click=export_data, width=400,
                         style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE, color=ft.Colors.WHITE, padding=20)),
        
        ft.Divider(),
        ft.Row([ft.ProgressRing(ref=progress_ring, visible=False, width=20, height=20),
                ft.Text(ref=status_text, size=16, selectable=True, expand=True)]),
        
        ft.Text("* - обязательные поля", size=12, color=ft.Colors.GREY)
    )

if __name__ == "__main__":
    ft.app(target=main)