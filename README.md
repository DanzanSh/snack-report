# SnackReport

Converts the UCS Shelter export «Сводка по доходу» (`.xls`) into the «Ежедневный отчет о наличии снеков и напитков» daily report (`.xlsx`).

Each daily report has two sheets:
1. **Ежедневный кассовый отчет**: date, «Дорожные» = 1000 (always), «Остаток наличных на начала дня» (`наличные`) is taken from the previous report's «Остаток наличных на конец дня». The rest (other expenses and their notes, cash, non-cash, приход, предоплаты, инкасация, администратор, тех персонал, охрана) is cleared every day and filled in by hand.
2. **Наличие снеков и напитков**: the stock report (Excel limits sheet names to 31 characters, so the full title «Ежедневный отчет о наличии снеков и напитков» is the heading inside the sheet).

If a previous report has no cash sheet (older format), one is added and the opening cash is entered by hand.

What the script does for the stock sheet:
- **ОстУтро**: today's opening stock (ОстУтро) is copied from yesterday's report (its ОстВечер column). If yesterday's ОстВечер is negative, ОстУтро starts at 0.
- **Продано**: the sold count comes from the «Услуги» block of the UCS export. Payments are ignored.
  When UCS names an item differently from the report (e.g. «Вода 1 л» → «Вода 1»), the sale goes into the report's row; see «Item names» below.
- **ОстВечер**: a formula, `ОстУтро + Приход − Продано − Списание` (Долг is not subtracted, as in the reference report). It recalculates when you enter Приход or Списание by hand. Negative values are highlighted in red.
- **Приход, Долг, Имя, Списание, Причина**: cleared every day and filled in by hand.
- **Staff rows** (Утеря карт, Наличие карт, Администратор, Тех персонал, Охрана): cleared every day.

## Usage (Windows)

```
any folder\
  SnackReport.exe
  report\                         <- put new UCS exports here
    Сводка_04.11.xls              <- not processed yet
    2026-10 Октябрь\              <- one folder per month, created by the program
      Отчёт_2026-10-28.xlsx       <- daily report
      Месяц_10.2026.xlsx          <- monthly report, updated on every launch
      UCS\
        Сводка_28.10.xls          <- processed exports are moved here
    2026-11 Ноябрь\
```

1. Put `SnackReport.exe` in any folder. On first launch it creates a `report` folder next to itself.
2. Every day, save the UCS export (`.xls`, any file name) into `report` and double-click `SnackReport.exe`.
3. For every export whose date has no report yet, the program creates `Отчёт_ГГГГ-ММ-ДД.xlsx`
   in that month's folder, in date order. ОстУтро is taken from the previous report, even if it is in last month's folder.
   If a report for that date already exists, it is rebuilt (see «Rebuilding a report» below).
4. After that, every export with a report is moved to `<month>\UCS`. If the same name is already there,
   ` (2)` is added. Files the program cannot read stay in `report`, so whatever is left there needs attention.
   Reports from older versions (`Отчёт_ДД.ММ.ГГГГ.xlsx` directly in `report`) are moved into month folders
   automatically on the next launch. The old file name format is still read.
5. **First report:** the template is built into the exe. In the first report, enter ОстУтро by hand;
   ОстВечер recalculates on its own, and the next day takes its stock from there.

## Rebuilding a report

If a report for a day is already made and you get a new UCS export for that day, put the new export into `report` and launch the program
(or drag it onto the exe). The report is rebuilt:
- Продано is taken from the new export; new items get their own rows.
- Everything entered by hand stays: ОстУтро, Приход, Долг, Имя, Списание, Причина, staff rows and the whole cash sheet.
- Reports for later days are updated in a chain: their ОстУтро (from the previous day's ОстВечер) and «Остаток наличных на начала дня».
  Everything else in them stays. If ОстУтро was corrected by hand in a later report, that correction is overwritten.
- If the report is open in Excel, nothing is saved: close the file and launch again (the export stays in `report`).

## Quarterly cleanup

When a new quarter starts (January, April, July, October), the program offers to delete daily reports and UCS exports
for the earlier quarters. Monthly reports and month folders stay. Nothing is deleted without «Да».
If you answer «Нет», it asks again next quarter (the answer is kept in `report\.cleanup`).
The month with the latest daily report is never cleaned: the next day takes its ОстУтро from that report.
In `--console` mode the program only mentions old files and never deletes them.

## Monthly report

On every launch the program also updates `Месяц_ММ.ГГГГ.xlsx` in each month's folder. It is built from that month's daily reports, so it includes Приход, Долг and Списание entered by hand:

| Товар | ОстНачало | Приход | Продано | Долг | Списание | ОстКонец |
|---|---|---|---|---|---|---|

- ОстНачало: ОстУтро of the month's first daily report. ОстКонец: ОстВечер of its last.
- Приход, Продано, Долг, Списание: totals for the month. An ИТОГО row sits at the bottom.
- A monthly report is rebuilt when any daily report of that month is newer than it,
  e.g. after you entered Приход and saved the daily file.
- Days without a report (weekends, holidays) are not an error and are not flagged. The next day's ОстУтро is taken from the latest earlier report.

You can also drag a single `.xls` onto the exe: the report goes into its month folder in `report` (an existing report for that date is rebuilt). The dragged file itself is not moved.
Errors are written to `snack-report.log` next to the exe.

## Adding a new item

Just sell it through UCS: the row appears in the report automatically (before «Прочее»), and the program warns about it.
To add an item ahead of time, add a row to the latest daily report in the current month's folder, named exactly as in UCS.
The newest report serves as the template for the next day.

### Item names

Some items are named differently in UCS and in the report. The program has a built-in list
(«Вода 1 л» → «Вода 1», «RedBull 0.25» → «Ред Булл», «Протеин 30гр» → «Протеин в пачках штучно» and others, see `config/settings.ini`).
To add or change a pair, put a `[aliases]` section into `settings.ini` next to the exe:
```
[aliases]
Название в UCS = Строка отчёта
```
Several UCS names can point to one row; their sales are added up.

Subscriptions and visits from UCS are not counted in the report: every item of the UCS group «Посещения»
(«12 посещений», «Абонемент MIX 1мес», «Бойцовская зона группа 12пос»…) is skipped.
The groups are `ignored_groups` in the `[report]` section of `settings.ini`; single items can be skipped with
`ignored` (comma-separated, `*` matches any text).

Exceptions inside «Посещения»:
- «Разовые зал Бокс» is split by price: 4000 per visit goes to «Разовые посещения зал», 5000 to «Разовое Бокс»
  (a mixed row, e.g. 3 pcs for 13000, becomes 2 + 1). If the sum can't be split, the program says so and
  the row must be filled by hand. The rules are in the `[price_split]` section of `settings.ini`:
  `Название в UCS = 4000: Строка; 5000: Строка`.
- «утеря карточки» goes to «Продано» of the staff row «Утеря карт». Any UCS item with an alias in `[aliases]`
  is taken even from a skipped group.
«Разовые посещения зал» is a report row and is still filled.

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
