# SnackReport

Converts the UCS Shelter export «Сводка по доходу» (`.xls`) into the «Ежедневный отчет о наличии снеков и напитков» daily report (`.xlsx`).

Each daily report has two sheets:
1. **Ежедневный кассовый отчет**: date, «Дорожные» = 1000 (always), «Остаток наличных на начала дня» (`наличные`) is taken from the previous report's «Остаток наличных на конец дня». The rest (other expenses and their notes, cash, non-cash, приход, предоплаты, инкасация, администратор, тех персонал, охрана) is cleared every day and filled in by hand.
2. **Наличие снеков и напитков**: the stock report (Excel limits sheet names to 31 characters, so the full title «Ежедневный отчет о наличии снеков и напитков» is the heading inside the sheet).

If a previous report has no cash sheet (older format), one is added and the opening cash is entered by hand.

What the script does for the stock sheet:
- **ОстУтро**: today's opening stock (ОстУтро) is copied from yesterday's report (its ОстВечер column). If yesterday's ОстВечер is negative, ОстУтро starts at 0.
- **Продано**: the sold count comes from the «Услуги» block of the UCS export. Payments are ignored.
- **ОстВечер**: a formula, `ОстУтро + Приход − Продано − Долг − Списание`. It recalculates when you enter Приход, Долг or Списание by hand. Negative values are highlighted in red.
- **Приход, Долг, Имя, Списание, Причина**: cleared every day and filled in by hand.
- **Staff rows** (Утеря карт, Наличие карт, Администратор, Тех персонал, Охрана): cleared every day.

## Usage (Windows)

```
any folder\
  SnackReport.exe
  report\                     <- UCS exports and reports go here
    Сводка_28.10.xls
    Отчёт_28.10.2026.xlsx     <- daily report, created by the program
    Месяц_10.2026.xlsx        <- monthly report, updated on every launch
```

1. Put `SnackReport.exe` in any folder. On first launch it creates a `report` folder next to itself.
2. Every day, save the UCS export (`.xls`, any file name) into `report` and double-click `SnackReport.exe`.
3. For every export whose date has no report yet, the program creates `Отчёт_ДД.ММ.ГГГГ.xlsx`
   in the same folder, in date order. ОстУтро is taken from the previous report.
   Exports that already have a report are left alone.
4. **First report:** the template is built into the exe. In the first report, enter ОстУтро by hand;
   ОстВечер recalculates on its own, and the next day takes its stock from there.

## Monthly report

On every launch the program also updates `report\Месяц_ММ.ГГГГ.xlsx`. It is built from that month's daily reports, so it includes Приход, Долг and Списание entered by hand:

| Товар | ОстНачало | Приход | Продано | Долг | Списание | ОстКонец |
|---|---|---|---|---|---|---|

- ОстНачало: ОстУтро of the month's first daily report. ОстКонец: ОстВечер of its last.
- Приход, Продано, Долг, Списание: totals for the month. An ИТОГО row sits at the bottom.
- A monthly report is rebuilt when any daily report of that month is newer than it,
  e.g. after you entered Приход and saved the daily file.
- Days without a report (weekends, holidays) are not an error and are not flagged. The next day's ОстУтро is taken from the latest earlier report.

You can also drag a single `.xls` onto the exe: the report goes into `report` (it asks before overwriting).
Errors are written to `snack-report.log` next to the exe.

## Adding a new item

Just sell it through UCS: the row appears in the report automatically.
To add an item ahead of time, add a row to the latest report in `report`, named exactly as in UCS.
The newest report serves as the template for the next day.

Which rows count as sections and which as staff rows is set in `settings.ini` next to the exe (sample: `config/settings.ini`).

## Command line

```
SnackReport.exe [ucs.xls] [--prev отчёт.xlsx] [--out папка] [--force] [--console]
```
Without arguments it processes the `report` folder. To change settings, place `settings.ini` next to the exe
(sample: `config/settings.ini`).
```
```

## Development

```
pip install -r requirements-dev.txt
python -m pytest tests
python run_snack_report.py tests\fixtures\test.xls --console
build.bat                       # build dist\SnackReport.exe on Windows
tools/build_exe_docker.sh       # same on Linux via Docker + Wine
```
`tools/make_template.py` regenerates `config/template.xlsx` from scratch. The item list was transcribed from a photo, so check it.
