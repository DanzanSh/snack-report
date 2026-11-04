"""Создаёт config/template.xlsx: пустую форму отчёта (список товаров переписан с фото).

Запуск: python tools/make_template.py
После этого откройте template.xlsx и один раз впишите «ОстУтро» на текущий день.
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from snack_report.cash_report import SNACK_SHEET, ensure_sheets  # noqa: E402

ITEMS = [
    "Батончик 1000", "Батончик 1200", "Батончик 800", "Батончик 700", "Батончик 1500",
    "печенье барное", "Козинаки", "торт Сникерс", "Конфеты штучно", "Банан",
    "Вода 0,5", "Вода 1", "Вода 1,5", "Американо", "экспанента", "энергетик вейкапп",
    "Пепси кола", "Ред Булл", "Фаренгейт", "Энергетик Бомбар Вит мин", "Сок Грацио стекло 0.2",
    "Чизкейк", "Сок Дада пластик бут 0.5", "Бомбар молочный коктейль", "Гейнер", "Протеин в пачках штучно", "Протеин порц",
    "BCAA порц", "Креатин порц", "Медовик", "Шот Гуарана", "Футболки 777", "каллоген порция",
    "тапочки одноразовые", "резинка для волос", "Носки", "Полотенца", "калоген?", "Оптимен",
    "Психотик", "конфеты пачка", "Витамин D3 180 капсул", "Омега", "Протеин оптом",
    "Креатин опт", "Шампунь", "Зефир 750", "Зефир 1000",
]
SECTION = "Прочее"
FOOD = ["Сэндвич", "омлет с курицей", "Рис с яйцами", "Пита", "Рис курица",
        "Разовое Бокс", "Разовые посещения зал"]
SERVICE = ["Утеря карт", "Наличие карт", "Администратор", "Тех персонал", "Охрана"]
HEADERS = ["Наименование", "Приход", "ОстУтро", "Продано", "Долг", "Имя", "Списание", "Причина", "ОстВечер"]
WIDTHS = [32, 9, 10, 10, 8, 14, 10, 18, 10]


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SNACK_SHEET
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    beige = PatternFill("solid", fgColor="F2DCCB")

    ws["A1"].number_format = r"dd\.mm\.yyyy"
    ws["A1"].font = Font(bold=True)
    ws["B1"] = "Ежедневный отчет о наличии снеков и напитков"
    ws["B1"].font = Font(bold=True)
    ws["B1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("B1:I1")

    ws.append(HEADERS)
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
        cell = ws.cell(2, i)
        cell.font = Font(bold=True)
        cell.fill = beige
        cell.alignment = Alignment(horizontal="center")

    for name in ITEMS + [SECTION] + FOOD + SERVICE:
        ws.append([name])
    last = ws.max_row
    for row in ws.iter_rows(min_row=1, max_row=last, max_col=len(HEADERS)):
        for cell in row:
            cell.border = border
            if cell.column > 1 and cell.row > 2:
                cell.alignment = Alignment(horizontal="center")
    section_row = 3 + len(ITEMS)
    for c in range(1, len(HEADERS) + 1):
        ws.cell(section_row, c).fill = beige

    # Отрицательный остаток на вечер подсвечивается красным
    red = PatternFill("solid", fgColor="FFC7CE")
    ws.conditional_formatting.add(
        f"I3:I{last}", CellIsRule(operator="lessThan", formula=["0"], fill=red, font=Font(color="9C0006"))
    )
    ws.freeze_panes = "B3"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 1

    ensure_sheets(wb)  # первым добавляется кассовый лист

    out = Path(__file__).resolve().parents[1] / "config" / "template.xlsx"
    wb.save(out)
    print(f"Шаблон сохранён: {out}")


if __name__ == "__main__":
    main()
