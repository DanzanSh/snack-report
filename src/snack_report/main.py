r"""SnackReport: выгрузки UCS «Сводка по доходу» → ежедневные отчёты о наличии снеков и напитков.

Запуск без аргументов: берёт все выгрузки UCS (.xls) из папки `report` рядом с exe
и для каждой даты, по которой ещё нет отчёта, создаёт report\ГГГГ-ММ Месяц\Отчёт_ГГГГ-ММ-ДД.xlsx.
Обработанные выгрузки переносятся в report\ГГГГ-ММ Месяц\UCS.
Новая выгрузка за дату, по которой отчёт уже есть, пересоздаёт его: меняется «Продано»,
вписанное вручную остаётся, у следующих дней пересчитываются остатки на утро.
"""
from __future__ import annotations

import argparse
import configparser
import datetime as dt
import logging
import sys
from dataclasses import dataclass
from pathlib import Path

from .report_builder import (
    DEFAULT_ALIASES,
    DEFAULT_IGNORED,
    DEFAULT_SECTION_ROWS,
    DEFAULT_STAFF_ROWS,
    BuildResult,
    build_report,
    find_previous_report,
    find_reports,
    rebuild_report,
    refresh_opening,
)
from .cleanup import declined, find_old, remember_declined, run_cleanup
from .layout import move_into, organize, report_path, ucs_archive_dir
from .monthly import MonthResult, update_monthly
from .ucs_reader import DEFAULT_IGNORED_GROUPS, DEFAULT_PRICE_SPLIT, UcsFormatError, UcsReport, read_ucs

log = logging.getLogger("snack_report")
DEFAULT_REPORT_DIR = "report"


def base_dir() -> Path:
    """Папка рядом с exe (или корень проекта при запуске из исходников)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def bundled_template() -> Path:
    """Шаблон, зашитый в exe (или config/template.xlsx при запуске из исходников)."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS")) / "template.xlsx"
    return Path(__file__).resolve().parents[2] / "config" / "template.xlsx"


def load_settings(base: Path) -> dict:
    """Необязательный settings.ini рядом с exe (или в config/)."""
    cfg = configparser.ConfigParser(delimiters=("=",))
    cfg.optionxform = str  # названия товаров с сохранением регистра
    cfg.read([base / "config" / "settings.ini", base / "settings.ini"], encoding="utf-8-sig")
    resolve = lambda v: (base / v) if not Path(v).is_absolute() else Path(v)  # noqa: E731

    def rows(option: str, default: tuple[str, ...]) -> tuple[str, ...]:
        v = cfg.get("report", option, fallback=None)
        return tuple(x.strip() for x in v.split(",") if x.strip()) if v is not None else default

    template = cfg.get("paths", "template", fallback=None)
    aliases = dict(DEFAULT_ALIASES)
    if cfg.has_section("aliases"):
        aliases.update({k.strip(): v.strip() for k, v in cfg.items("aliases") if v.strip()})
    price_split = {k: dict(v) for k, v in DEFAULT_PRICE_SPLIT.items()}
    if cfg.has_section("price_split"):
        for name, spec in cfg.items("price_split"):
            prices = _parse_prices(spec)
            if prices:
                price_split[name.strip()] = prices
    return {
        "reports_dir": resolve(cfg.get("paths", "reports_dir", fallback=DEFAULT_REPORT_DIR)),
        "template": resolve(template) if template else bundled_template(),
        "section_rows": rows("section_rows", DEFAULT_SECTION_ROWS),
        "staff_rows": rows("staff_rows", DEFAULT_STAFF_ROWS),
        "aliases": aliases,
        "ignored": rows("ignored", DEFAULT_IGNORED),
        "ignored_groups": rows("ignored_groups", DEFAULT_IGNORED_GROUPS),
        "price_split": price_split,
    }


def _parse_prices(spec: str) -> dict[float, str]:
    """«4000: Разовые посещения зал; 5000: Разовое Бокс» -> {4000.0: ..., 5000.0: ...}."""
    prices = {}
    for part in spec.split(";"):
        price, sep, row = part.partition(":")
        if sep and row.strip():
            prices[float(price.strip().replace(",", ".").replace(" ", ""))] = row.strip()
    return prices


def _read(path: Path, s: dict) -> UcsReport:
    return read_ucs(path, s["ignored_groups"], s["price_split"], keep=s["aliases"])


class Ui:
    """Окна (по умолчанию) или консоль (--console)."""

    def __init__(self, gui: bool):
        self.gui = gui
        if gui:
            import tkinter as tk
            from tkinter import messagebox

            self._root = tk.Tk()
            self._root.withdraw()
            self._mb = messagebox

    def confirm(self, text: str) -> bool:
        return self._mb.askyesno("SnackReport", text) if self.gui else False

    def info(self, text: str) -> None:
        self._mb.showinfo("SnackReport", text) if self.gui else print(text)

    def warn(self, text: str) -> None:
        self._mb.showwarning("SnackReport", text) if self.gui else print(text)

    def error(self, text: str) -> None:
        self._mb.showerror("SnackReport", text) if self.gui else print(text, file=sys.stderr)


def run(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="SnackReport", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ucs", nargs="?", type=Path,
                    help="одна выгрузка UCS (.xls); без неё обрабатывается вся папка report")
    ap.add_argument("--prev", type=Path, help="вчерашний отчёт для нового отчёта (по умолчанию — самый свежий в папке)")
    ap.add_argument("--out", type=Path, help="папка с выгрузками и отчётами (по умолчанию report рядом с exe)")
    ap.add_argument("--force", action="store_true", help="(устарело, ни на что не влияет: отчёт пересоздаётся всегда)")
    ap.add_argument("--console", action="store_true", help="без окон, вывод в консоль")
    args = ap.parse_args(argv)

    base = base_dir()
    logging.basicConfig(
        filename=base / "snack-report.log",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )
    ui = Ui(gui=not args.console)
    try:
        s = load_settings(base)
        folder: Path = args.out or s["reports_dir"]
        if args.ucs:
            return _run_single(args, s, folder, ui)
        return _run_folder(s, folder, ui)
    except Exception as e:  # показываем пользователю любую ошибку, полный текст — в лог
        log.exception("Ошибка")
        ui.error(f"Ошибка: {e}\n\nПодробности в {base / 'snack-report.log'}")
        return 1


def _monthly(s: dict, folder: Path) -> list[str]:
    """Пересобирает месячные отчёты. Возвращает строки для итогового окна."""
    results: list[MonthResult] = update_monthly(folder, s["section_rows"], s["staff_rows"])
    lines = []
    for m in results:
        log.info("Месячный отчёт %s: дней %d", m.output.name, m.days)
        lines.append(f"{m.output.name}: дневных отчётов {m.days}")
    return lines


@dataclass
class Job:
    ucs_path: Path
    ucs: UcsReport


def _build(job: Job, s: dict, folder: Path, out: Path, prev: Path | None = None) -> BuildResult:
    if prev:
        source, from_template = prev, prev.resolve() == s["template"].resolve()
    else:
        found = find_previous_report(folder, job.ucs.date)
        source, from_template = (found, False) if found else (s["template"], True)
    if not source.exists():
        raise FileNotFoundError(f"Не найден файл {source}")
    log.info("%s → %s (источник %s, шаблон: %s)", job.ucs_path.name, out.name, source.name, from_template)
    res = build_report(
        date=job.ucs.date,
        sales=job.ucs.sales,
        source=source,
        from_template=from_template,
        output=out,
        section_rows=s["section_rows"],
        staff_rows=s["staff_rows"],
        aliases=s["aliases"],
        ignored=s["ignored"],
    )
    res.ignored = job.ucs.ignored + res.ignored
    res.unsplit = list(job.ucs.unsplit)
    log.info("Сохранено %s; заполнено %d; новых строк %d", out, len(res.filled), len(res.added))
    return res


def _rebuild(job: Job, s: dict, existing: Path) -> BuildResult:
    log.info("%s → пересоздание %s", job.ucs_path.name, existing.name)
    res = rebuild_report(
        date=job.ucs.date,
        sales=job.ucs.sales,
        existing=existing,
        section_rows=s["section_rows"],
        staff_rows=s["staff_rows"],
        aliases=s["aliases"],
        ignored=s["ignored"],
    )
    res.ignored = job.ucs.ignored + res.ignored
    res.unsplit = list(job.ucs.unsplit)
    log.info("Пересоздан %s; заполнено %d; новых строк %d", existing, len(res.filled), len(res.added))
    return res


LOCKED = "закройте файл в Excel и запустите программу снова"


def _chain(s: dict, folder: Path, start: dt.date) -> tuple[list[str], list[str]]:
    """Обновляет ОстУтро и остаток наличных в отчётах после даты `start`, по цепочке.
    Возвращает (строки для окна, проблемы)."""
    reports = find_reports(folder)
    dates = sorted(reports)
    lines: list[str] = []
    for prev_d, d in zip(dates, dates[1:]):
        if d <= start:
            continue
        try:
            changed = refresh_opening(reports[d], reports[prev_d], s["section_rows"], s["staff_rows"])
        except PermissionError:
            log.warning("Не удалось обновить %s: файл занят", reports[d])
            return lines, [f"{reports[d].name}: остатки на утро не обновлены, {LOCKED}"]
        if changed:
            log.info("Обновлены остатки на утро %s: %s", reports[d].name, ", ".join(changed))
            lines.append(f"{reports[d].name}: {', '.join(changed)}")
    return lines, []


def _today() -> dt.date:
    return dt.date.today()


def _cleanup(folder: Path, ui: Ui) -> list[str]:
    """Раз в квартал предлагает удалить дневные отчёты и выгрузки UCS за прошлые кварталы."""
    today = _today()
    plan = find_old(folder, today)
    if not plan.files or declined(folder, today):
        return []
    months = plan.describe_months()
    if not ui.gui:  # без окон не спрашиваем и отказ не запоминаем
        return [f"Можно удалить старые дневные отчёты и выгрузки UCS ({months}): запустите программу без --console.", ""]
    question = (f"Начался новый квартал.\n\nУдалить дневные отчёты и выгрузки UCS за: {months}?\n"
                f"Файлов: {len(plan.files)}. Месячные отчёты останутся.")
    if not ui.confirm(question):
        remember_declined(folder, today)
        log.info("Очистка за %s: отказ", months)
        return []
    deleted, errors = run_cleanup(folder, plan)
    for p in deleted:
        log.info("Удалён %s", p)
    return [f"Удалено старых файлов: {len(deleted)} ({months})", ""] + _bullets("Не удалены:", errors)


def _describe(res: BuildResult, ucs_name: str) -> list[str]:
    fmt = lambda x: f"{x:g}"  # noqa: E731
    if res.rebuilt:
        lines = [f"{res.output.name}  ← {ucs_name}  (пересоздан, вписанное вручную сохранено)"]
    else:
        lines = [f"{res.output.name}  ← {ucs_name}"]
        if res.from_template:
            lines.append("  Первый отчёт: впишите в нём «ОстУтро» вручную, «ОстВечер» пересчитается сам.")
        else:
            lines.append(f"  Остатки на утро из: {res.source.name}")
        if res.cash_opening is not None:
            lines.append(f"  Остаток наличных на начало дня: {fmt(res.cash_opening)}")
        else:
            lines.append("  Кассовый отчёт: впишите «Остаток наличных на начала дня» вручную.")
    lines.append(f"  Продано: {', '.join(f'{n} {fmt(q)}' for n, q in res.filled) or 'нет продаж'}")
    if res.ignored:
        lines.append(f"  Не учитываются (абонементы): {', '.join(res.ignored)}")
    if res.unsplit:
        lines.append(f"  Не разнесено по цене (впишите вручную): {', '.join(res.unsplit)}")
    if res.added:
        lines.append(f"  Новые товары (впишите им «ОстУтро»): {', '.join(res.added)}")
    if res.negative and not res.from_template:
        lines.append(f"  Отрицательный остаток: {', '.join(f'{n} {fmt(q)}' for n, q in res.negative)}")
    return lines


def _bullets(title: str, items: list[str]) -> list[str]:
    return [title] + [f"  • {x}" for x in items] + [""] if items else []


def _run_folder(s: dict, folder: Path, ui: Ui) -> int:
    if not folder.is_dir():
        folder.mkdir(parents=True)
        ui.warn(f"Создана папка {folder}.\nПоложите в неё выгрузку UCS (.xls) и запустите программу снова.")
        return 2

    problems = organize(folder)  # отчёты из корня (старая раскладка) — по папкам месяцев
    jobs: dict = {}
    skipped: list[str] = []
    older: list[Job] = []  # выгрузки, для даты которых есть файл новее
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.suffix.lower() != ".xls" or p.name.startswith("~$"):
            continue
        try:
            ucs = _read(p, s)
        except UcsFormatError as e:
            log.warning("Пропущен %s: %s", p.name, e)
            skipped.append(f"{p.name}: {e}")
            continue
        job, other = Job(p, ucs), jobs.get(ucs.date)
        if other is None or p.stat().st_mtime > other.ucs_path.stat().st_mtime:
            if other:
                skipped.append(f"{other.ucs_path.name}: за {ucs.date:%d.%m.%Y} есть файл новее ({p.name})")
                older.append(other)
            jobs[ucs.date] = job
        else:
            skipped.append(f"{p.name}: за {ucs.date:%d.%m.%Y} есть файл новее ({other.ucs_path.name})")
            older.append(job)

    existing = find_reports(folder)
    created = [d for d in jobs if d not in existing]
    rebuilt = [d for d in jobs if d in existing]
    lines: list[str] = []
    warn = bool(skipped or problems)
    if jobs:
        counts = [f"создано отчётов: {len(created)}"] if created else []
        counts += [f"пересоздано: {len(rebuilt)}"] if rebuilt else []
        lines += [f"{', '.join(counts).capitalize()} (папка {folder})", ""]
    else:
        lines += [f"Новых выгрузок нет. Положите выгрузку UCS (.xls) в папку {folder}.", ""]
        warn = True

    failed: set[dt.date] = set()
    for d in sorted(jobs):  # по возрастанию даты: каждый следующий берёт остатки из предыдущего
        job = jobs[d]
        try:
            if d in existing:
                res = _rebuild(job, s, existing[d])
            else:
                res = _build(job, s, folder, report_path(folder, d))
        except PermissionError:
            log.warning("Не удалось сохранить отчёт за %s: файл занят", d)
            problems.append(f"{job.ucs_path.name}: отчёт за {d:%d.%m.%Y} не сохранён, {LOCKED}")
            failed.add(d)
            warn = True
            continue
        lines += _describe(res, job.ucs_path.name) + [""]
        warn |= bool(res.added) or bool(res.unsplit) or bool(res.negative and not res.from_template)

    done = [d for d in jobs if d not in failed]
    if done:  # отчёты за более поздние даты берут остатки на утро из обновлённых
        chain, chain_problems = _chain(s, folder, min(done))
        lines += _bullets("Обновлены остатки на утро:", chain)
        problems += chain_problems
        warn |= bool(chain_problems)

    # отчёт по выгрузке есть — переносим её в архив месяца, в корне остаются только необработанные
    for job in [*(jobs[d] for d in done), *older]:
        try:
            dst = move_into(job.ucs_path, ucs_archive_dir(folder, job.ucs.date))
            log.info("%s перенесён в %s", job.ucs_path.name, dst)
        except OSError as e:
            log.warning("Не удалось перенести %s: %s", job.ucs_path, e)
            problems.append(f"{job.ucs_path.name}: не удалось перенести в архив ({e})")
            warn = True

    lines += _bullets("Пропущены:", skipped) + _bullets("Не обработаны:", problems)
    month_lines = _monthly(s, folder)
    if month_lines:
        lines += ["Месячные отчёты:"] + month_lines + [""]
    lines += _cleanup(folder, ui)
    (ui.warn if warn else ui.info)("\n".join(lines).rstrip())
    return 0 if jobs or existing else 2


def _run_single(args, s: dict, folder: Path, ui: Ui) -> int:
    job = Job(args.ucs, _read(args.ucs, s))
    existing = find_reports(folder).get(job.ucs.date)
    try:
        if existing:  # отчёт уже есть — пересоздаём, вписанное вручную сохраняется
            res = _rebuild(job, s, existing)
        else:
            res = _build(job, s, folder, report_path(folder, job.ucs.date), args.prev)
    except PermissionError:
        ui.error(f"Отчёт за {job.ucs.date:%d.%m.%Y} не сохранён: {LOCKED}.")
        return 3
    text = "Готово:\n" + "\n".join(_describe(res, job.ucs_path.name))
    chain, chain_problems = _chain(s, folder, job.ucs.date)
    if chain:
        text += "\n\nОбновлены остатки на утро:\n" + "\n".join(f"  • {x}" for x in chain)
    if chain_problems:
        text += "\n\n" + "\n".join(chain_problems)
    month_lines = _monthly(s, folder)
    if month_lines:
        text += "\n\nМесячные отчёты:\n" + "\n".join(month_lines)
    (ui.warn if res.added or chain_problems else ui.info)(text)
    return 0


if __name__ == "__main__":
    sys.exit(run())
