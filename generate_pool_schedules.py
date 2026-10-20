
from datetime import date, datetime, timedelta
from pathlib import Path
import argparse
import re
import subprocess
import sys

from pptx import Presentation
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://recreation.rutgers.edu/recreation-list-calendar?calendar_filter=122"
BASE = Path(__file__).resolve().parent
OUT = BASE / "output"

TEMPLATES = {
    "Olympic Pool": BASE / "Olympic Pool Spring 2026 Hours.pptx",
    "Patio Pool": BASE / "Patio Pool Spring 2026 Hours.pptx",
}

ALIASES = {
    "Olympic Pool": ("olympic pool",),
    "Patio Pool": ("patio pool", "fixed depth patio pool"),
}

MONTHS = (
    "January|February|March|April|May|June|July|August|"
    "September|October|November|December"
)
DATE_RE = re.compile(
    rf"^({MONTHS})\s+(\d{{1,2}}),\s*(\d{{4}})", re.I
)
EVENT_RE = re.compile(
    r"^(\d{1,2}:\d{2}\s*(?:am|pm))\s*-\s*"
    r"(\d{1,2}:\d{2}\s*(?:am|pm))\s*\|\s*\|\s*(.+)$", re.I
)


def scrape_week(monday):
    """Read calendar text and keep entries for the requested week."""
    collected = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except PlaywrightTimeoutError:
                pass

            for selector in (
                "button.fc-listWeek-button",
                "button[title*='list']",
                "button[aria-label*='list']",
            ):
                try:
                    button = page.locator(selector).first
                    if button.count() and button.is_visible():
                        button.click(timeout=1500)
                        page.wait_for_timeout(500)
                        break
                except Exception:
                    pass

            def body_text():
                return page.locator("body").inner_text()

            def calendar_title():
                for selector in (".fc-toolbar-title", "h2", "h1"):
                    try:
                        value = page.locator(selector).first.inner_text(timeout=700)
                        if re.search(MONTHS, value, re.I):
                            return value.strip()
                    except Exception:
                        pass
                return ""

            for _ in range(24):
                match = re.search(
                    rf"({MONTHS})\s+(\d{{4}})", calendar_title(), re.I
                )
                if not match:
                    break
                shown = datetime.strptime(match.group(0), "%B %Y")
                if (shown.year, shown.month) == (monday.year, monday.month):
                    break
                selector = (
                    "button.fc-prev-button"
                    if (shown.year, shown.month) > (monday.year, monday.month)
                    else "button.fc-next-button"
                )
                try:
                    page.locator(selector).click(timeout=1500)
                    page.wait_for_timeout(350)
                except Exception:
                    break

            collected.append(body_text())

            target = re.compile(
                rf"{re.escape(monday.strftime('%B'))}\s+{monday.day},\s*"
                rf"{monday.year}", re.I
            )

            if not any(target.search(t) for t in collected):
                for _ in range(10):
                    try:
                        page.locator("button.fc-prev-button").click(timeout=1000)
                        page.wait_for_timeout(250)
                        collected.append(body_text())
                        if any(target.search(t) for t in collected):
                            break
                    except Exception:
                        break

                for _ in range(12):
                    try:
                        page.locator("button.fc-next-button").click(timeout=1000)
                        page.wait_for_timeout(250)
                        collected.append(body_text())
                        if any(target.search(t) for t in collected):
                            break
                    except Exception:
                        break
        finally:
            browser.close()

    events = []
    for text in collected:
        current_day = None
        for raw in text.splitlines():
            line = " ".join(raw.split()).strip()
            dm = DATE_RE.match(line)
            if dm:
                try:
                    current_day = datetime.strptime(
                        f"{dm.group(1)} {dm.group(2)} {dm.group(3)}",
                        "%B %d %Y",
                    ).date()
                except ValueError:
                    current_day = None
                continue

            em = EVENT_RE.match(line)
            if not em or current_day is None:
                continue
            if not monday <= current_day <= monday + timedelta(days=6):
                continue

            detail = em.group(3).strip()
            low = detail.casefold()
            if "swrc" not in low or not ("pool" in low or "swim" in low):
                continue

            events.append({
                "date": current_day,
                "start": em.group(1).lower(),
                "end": em.group(2).lower(),
                "detail": detail,
            })

    unique = {
        (e["date"], e["start"], e["end"], e["detail"]): e
        for e in events
    }
    return sorted(
        unique.values(),
        key=lambda e: (e["date"], e["start"], e["detail"].casefold()),
    )


def fmt_time(value):
    try:
        return datetime.strptime(value.upper(), "%I:%M %p").strftime("%I:%M%p").lstrip("0").replace(":00", "")
    except ValueError:
        return value.upper().replace(" ", "")


def pool_entries(events, pool):
    matches = []
    for event in events:
        low = event["detail"].casefold()
        if any(alias in low for alias in ALIASES[pool]):
            matches.append(event)
    return matches


def write_pdf(pool, monday, events):
    template = TEMPLATES[pool]
    if not template.exists():
        raise RuntimeError(f"Template not found: {template.name}")

    presentation = Presentation(str(template))
    slide = presentation.slides[0]
    slots = []

    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        content = shape.text.strip()
        if not content:
            continue
        if content.casefold() in ("olympic pool", "patio pool"):
            shape.text = pool.upper()
        elif content.casefold() == "swim meet" or re.search(
            r"\b\d{1,2}\s*(?:am|pm)\b", content, re.I
        ):
            slots.append(shape)

    slots.sort(key=lambda shape: shape.top)
    if len(slots) != 7:
        raise RuntimeError(
            f"Expected 7 schedule boxes in {template.name}; found {len(slots)}."
        )

    for index, shape in enumerate(slots):
        day = monday + timedelta(days=index)
        todays = [e for e in events if e["date"] == day]
        lines = []

        for event in todays:
            line = f"{fmt_time(event['start'])} - {fmt_time(event['end'])}"
            if "swim meet" in event["detail"].casefold():
                line += "  Swim Meet"
            lines.append(line)

        if not lines:
            lines = ["Hours not listed"]

        frame = shape.text_frame
        frame.clear()
        for i, line in enumerate(lines):
            paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            paragraph.text = line
            paragraph.alignment = 1
            for run in paragraph.runs:
                run.font.name = "Arial"

    OUT.mkdir(exist_ok=True)
    stem = f"{pool.replace(' ', '_')}_{monday.isoformat()}"
    pptx_path = OUT / f"{stem}.pptx"
    presentation.save(pptx_path)

    result = subprocess.run(
        [
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", str(OUT), str(pptx_path),
        ],
        capture_output=True, text=True, timeout=90,
    )
    pdf_path = OUT / f"{stem}.pdf"
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(f"PDF export failed: {result.stderr[-1000:]}")
    return pdf_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--week-start", required=True)
    parser.add_argument("--olympic", required=True)
    parser.add_argument("--patio", required=True)
    args = parser.parse_args()

    monday = datetime.strptime(args.week_start, "%Y-%m-%d").date()
    if monday.weekday() != 0:
        raise ValueError("The selected week must start on a Monday.")

    selected = []
    if args.olympic.lower() == "true":
        selected.append("Olympic Pool")
    if args.patio.lower() == "true":
        selected.append("Patio Pool")
    if not selected:
        raise ValueError("Select at least one pool.")

    events = scrape_week(monday)
    if not events:
        raise RuntimeError(
            "No SWRC pool events were read. Check the official calendar; "
            "the calendar layout may have changed."
        )

    for pool in selected:
        entries = pool_entries(events, pool)
        pdf = write_pdf(pool, monday, entries)
        print(f"Created {pdf}")

    report = OUT / "generation-report.txt"
    report.write_text(
        f"Week starting: {monday.isoformat()}\n"
        f"Official source: {URL}\n"
        "Any 'Hours not listed' row means no matching entry was found by "
        "the scraper. Verify missing hours and closures against the official "
        "calendar before using the PDFs.\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
