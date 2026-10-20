
#!/usr/bin/env python3
"""
Rutgers SWRC Pool Schedule Generator

Reads pool-only hours from the official Rutgers Recreation calendar,
updates the supplied PowerPoint templates, and exports separate PDFs.

Rules:
- No entries for a pool on a day means CLOSED.
- Overlapping and back-to-back sessions are merged.
- Separate sessions stay on one line with clear spacing.
- All time entries use one consistent font size, calculated from
  the longest line across the selected pools.
- CLOSED remains large and prominent.
- Text is centered horizontally and vertically.
- Stops if no dated SWRC pool events can be parsed.
"""

import argparse
import re
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from playwright.sync_api import sync_playwright


CALENDAR_URL = (
    "https://recreation.rutgers.edu/"
    "recreation-list-calendar?calendar_filter=122"
)

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"

# Font settings, in points.
MAX_TIME_FONT_SIZE = 36
MIN_TIME_FONT_SIZE = 18
CLOSED_FONT_SIZE = 36

# Approximate character width used to estimate whether a line fits.
# The font size is capped conservatively to avoid clipping.
ESTIMATED_CHARACTER_WIDTH = 0.52

# Extra space between separate sessions on the same line.
SESSION_SEPARATOR = "     "

POOLS = {
    "olympic": {
        "label": "Olympic Pool",
        "template": ROOT / "Olympic Pool Spring 2026 Hours.pptx",
        "match": re.compile(r"SWRC\s+Olympic\s+Pool", re.I),
    },
    "patio": {
        "label": "Patio Pool",
        "template": ROOT / "Patio Pool Spring 2026 Hours.pptx",
        "match": re.compile(
            r"SWRC\s+(?:Fixed\s+Depth\s+)?Patio\s+Pool", re.I
        ),
    },
}

TIME_RANGE_RE = re.compile(
    r"(?P<start>\d{1,2}:\d{2}\s*[ap]\.?m\.?)"
    r"\s*-\s*"
    r"(?P<end>\d{1,2}:\d{2}\s*[ap]\.?m\.?)",
    re.I,
)


def normalize_time(value):
    """Normalize a calendar time to a consistent readable format."""
    value = re.sub(r"\s+", " ", value.strip().lower())
    value = value.replace(".", "")
    parsed = datetime.strptime(value, "%I:%M %p")
    return parsed.strftime("%I:%M %p").lstrip("0")


def parse_time_range(line):
    match = TIME_RANGE_RE.search(line)
    if not match:
        return None

    try:
        start = normalize_time(match.group("start"))
        end = normalize_time(match.group("end"))
    except ValueError:
        return None

    return f"{start} - {end}"


def parse_date_from_line(line, default_year):
    """Return a date if the line contains a recognizable date."""
    iso = re.search(r"\b(\d{4})-(\d{2})-(\d{2})\b", line)

    if iso:
        try:
            return date(
                int(iso.group(1)),
                int(iso.group(2)),
                int(iso.group(3)),
            )
        except ValueError:
            return None

    month_match = re.search(
        r"\b(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+"
        r"(\d{1,2})(?:,?\s+(\d{4}))?\b",
        line,
        re.I,
    )

    if month_match:
        month = datetime.strptime(
            month_match.group(1), "%B"
        ).month
        day = int(month_match.group(2))
        year = int(month_match.group(3) or default_year)

        try:
            return date(year, month, day)
        except ValueError:
            return None

    return None


def fetch_calendar_text():
    """
    Try the rendered page first because the calendar may use JavaScript.
    Fall back to directly returned HTML if necessary.
    """
    rendered_text = ""
    browser = None

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch()
            page = browser.new_page()
            page.goto(
                CALENDAR_URL,
                wait_until="domcontentloaded",
                timeout=60000,
            )
            page.wait_for_timeout(2500)
            rendered_text = page.locator("body").inner_text(
                timeout=15000
            )
            browser.close()
            browser = None

    except Exception as exc:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass

        print(
            "Notice: rendered-calendar retrieval failed; "
            f"trying direct HTML instead ({exc})"
        )

    pool_names_found = (
        "SWRC Olympic Pool" in rendered_text
        or "SWRC Fixed Depth Patio Pool" in rendered_text
        or "SWRC Patio Pool" in rendered_text
    )

    if rendered_text and pool_names_found:
        return rendered_text

    response = requests.get(
        CALENDAR_URL,
        timeout=45,
        headers={
            "User-Agent": "RutgersPoolScheduleGenerator/1.0"
        },
    )
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    direct_text = soup.get_text("\n", strip=True)

    if direct_text:
        return direct_text

    raise RuntimeError(
        "The Rutgers calendar returned no readable text."
    )


def parse_calendar(text, week_start):
    """
    Parse dated calendar sections and pool event lines.

    Events without a reliably associated date are not assigned to a
    day by guessing.
    """
    events = []
    current_date = None
    year = week_start.year

    lines = [
        re.sub(r"\s+", " ", line).strip()
        for line in text.splitlines()
        if line.strip()
    ]

    for line in lines:
        parsed_date = parse_date_from_line(line, year)

        if parsed_date:
            current_date = parsed_date

        if not TIME_RANGE_RE.search(line):
            continue

        event_date = parsed_date or current_date

        if event_date is None:
            continue

        for pool_key, pool_info in POOLS.items():
            if not pool_info["match"].search(line):
                continue

            time_range = parse_time_range(line)

            if not time_range:
                continue

            events.append({
                "date": event_date,
                "pool": pool_key,
                "hours": time_range,
                "source_line": line,
            })

    # Deduplicate identical events in overlapping calendar sections.
    unique = {}

    for event in events:
        key = (
            event["date"],
            event["pool"],
            event["hours"],
            event["source_line"],
        )
        unique[key] = event

    return list(unique.values())


def monday_date(value):
    try:
        result = date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            "Enter the week start as YYYY-MM-DD."
        ) from exc

    if result.weekday() != 0:
        raise argparse.ArgumentTypeError(
            f"{result.isoformat()} is not a Monday. "
            "Choose the Monday of the schedule week."
        )

    return result


def pool_events_for_week(events, pool_key, week_start):
    end_date = week_start + timedelta(days=6)

    return [
        event
        for event in events
        if event["pool"] == pool_key
        and week_start <= event["date"] <= end_date
    ]


def time_to_minutes(value):
    """Convert a time such as '10:00 AM' to minutes after midnight."""
    parsed = datetime.strptime(value.strip().upper(), "%I:%M %p")
    return parsed.hour * 60 + parsed.minute


def format_display_time(minutes):
    """Format time compactly, for example 10AM or 4:30PM."""
    minutes = minutes % (24 * 60)
    hour24, minute = divmod(minutes, 60)
    suffix = "AM" if hour24 < 12 else "PM"
    hour12 = hour24 % 12 or 12

    if minute:
        return f"{hour12}:{minute:02d}{suffix}"

    return f"{hour12}{suffix}"


def merge_time_ranges(hours):
    """
    Merge overlapping or back-to-back ranges.

    Genuine gaps remain separate sessions.
    """
    ranges = []

    for value in hours:
        match = TIME_RANGE_RE.search(value)

        if not match:
            continue

        try:
            start = time_to_minutes(
                normalize_time(match.group("start"))
            )
            end = time_to_minutes(
                normalize_time(match.group("end"))
            )
        except ValueError:
            continue

        if end < start:
            end += 24 * 60

        ranges.append((start, end))

    if not ranges:
        return []

    ranges.sort()
    merged = [list(ranges[0])]

    for start, end in ranges[1:]:
        if start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])

    return [
        (
            f"{format_display_time(start)} - "
            f"{format_display_time(end)}"
        )
        for start, end in merged
    ]


def get_daily_display(pool_key, day, events):
    """Build one day's display string, merging sessions as needed."""
    day_events = [
        event
        for event in events
        if event.get("pool") == pool_key
        and event.get("date") == day
        and event.get("hours")
    ]

    merged_hours = merge_time_ranges(
        [event["hours"] for event in day_events]
    )

    if not merged_hours:
        return "CLOSED"

    return SESSION_SEPARATOR.join(merged_hours)


def get_hours_box(presentation):
    """Find the seven existing hours boxes in a template."""
    if len(presentation.slides) != 1:
        raise RuntimeError(
            "Expected one slide in the template; "
            f"found {len(presentation.slides)}."
        )

    slide = presentation.slides[0]

    expected_names = {
        f"Google Shape;{number};p1"
        for number in range(91, 98)
    }

    boxes = {
        shape.name: shape
        for shape in slide.shapes
        if shape.name in expected_names and shape.has_text_frame
    }

    missing = expected_names - set(boxes)

    if missing:
        raise RuntimeError(
            "Template is missing expected hours boxes: "
            f"{sorted(missing)}"
        )

    return [
        boxes[f"Google Shape;{number};p1"]
        for number in range(91, 98)
    ]


def calculate_shared_time_font_size(selected, week_start, events):
    """
    Calculate one font size for all time entries across the selected
    pools. The longest line in the narrowest hours box determines size.
    """
    longest_line = ""
    narrowest_width_pt = None

    days = [
        week_start + timedelta(days=i)
        for i in range(7)
    ]

    for pool_key in selected:
        template = POOLS[pool_key]["template"]

        if not template.is_file():
            raise RuntimeError(
                f"Template not found: {template.name}"
            )

        presentation = Presentation(str(template))
        boxes = get_hours_box(presentation)

        for day, shape in zip(days, boxes):
            display = get_daily_display(pool_key, day, events)

            if display == "CLOSED":
                continue

            if len(display) > len(longest_line):
                longest_line = display

            available_width_pt = (
                shape.width
                - Inches(0.12)
            ) / 12700

            if (
                narrowest_width_pt is None
                or available_width_pt < narrowest_width_pt
            ):
                narrowest_width_pt = available_width_pt

    if not longest_line or not narrowest_width_pt:
        return MAX_TIME_FONT_SIZE

    estimated_size = narrowest_width_pt / (
        len(longest_line) * ESTIMATED_CHARACTER_WIDTH
    )

    chosen_size = min(
        MAX_TIME_FONT_SIZE,
        estimated_size,
    )

    chosen_size = max(
        MIN_TIME_FONT_SIZE,
        chosen_size,
    )

    print(
        "Shared time font size: "
        f"{chosen_size:.1f} pt "
        f"(longest line: {longest_line})"
    )

    return chosen_size


def set_centered_text(shape, replacement, font_size):
    """
    Replace the text while retaining the first run's font styling.
    Center the text horizontally and vertically.
    """
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.word_wrap = False
    text_frame.margin_left = Inches(0.06)
    text_frame.margin_right = Inches(0.06)
    text_frame.margin_top = Inches(0.02)
    text_frame.margin_bottom = Inches(0.02)

    paragraphs = list(text_frame.paragraphs)

    if not paragraphs:
        paragraph = text_frame.add_paragraph()
    else:
        paragraph = paragraphs[0]

    all_runs = [
        run
        for item in paragraphs
        for run in item.runs
    ]

    if all_runs:
        first_run = all_runs[0]
        first_run.text = replacement

        for run in all_runs[1:]:
            run.text = ""
    else:
        first_run = paragraph.add_run()
        first_run.text = replacement

    # Remove leftover paragraphs, keeping the text on a single line.
    for item in paragraphs[1:]:
        for run in item.runs:
            run.text = ""

    paragraph.alignment = PP_ALIGN.CENTER
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    paragraph.line_spacing = 1.0

    # CLOSED stays large; all time entries share one calculated size.
    if replacement == "CLOSED":
        first_run.font.size = Pt(CLOSED_FONT_SIZE)
    else:
        first_run.font.size = Pt(font_size)


def update_template(
    pool_key,
    week_start,
    events,
    shared_time_font_size,
):
    """Update the seven hours boxes without redesigning the template."""
    if pool_key not in POOLS:
        raise RuntimeError(f"Unknown pool key: {pool_key}")

    pool_info = POOLS[pool_key]
    label = pool_info["label"]
    template = pool_info["template"]

    if not template.is_file():
        raise RuntimeError(
            f"Template not found: {template.name}"
        )

    presentation = Presentation(str(template))
    boxes = get_hours_box(presentation)

    days = [
        week_start + timedelta(days=i)
        for i in range(7)
    ]

    for day, shape in zip(days, boxes):
        display = get_daily_display(pool_key, day, events)

        set_centered_text(
            shape,
            display,
            shared_time_font_size,
        )

        print(f"{label} — {day.strftime('%A')}: {display}")

    OUTPUT.mkdir(parents=True, exist_ok=True)

    pptx_path = OUTPUT / (
        f"{label.replace(' ', '_')}_{week_start.isoformat()}.pptx"
    )

    presentation.save(str(pptx_path))
    return pptx_path


def export_pdf(pptx_path):
    """Export a PowerPoint file to PDF using LibreOffice."""
    result = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(OUTPUT),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )

    pdf_path = OUTPUT / f"{pptx_path.stem}.pdf"

    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            "LibreOffice could not export the PDF.\n"
            f"stdout: {result.stdout}\n"
            f"stderr: {result.stderr}"
        )

    return pdf_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--week-start", required=True)
    parser.add_argument("--olympic", default="true")
    parser.add_argument("--patio", default="true")
    args = parser.parse_args()

    week_start = monday_date(args.week_start)

    selected = []

    if args.olympic.lower() in ("true", "1", "yes"):
        selected.append("olympic")

    if args.patio.lower() in ("true", "1", "yes"):
        selected.append("patio")

    if not selected:
        raise RuntimeError("No pool was selected for generation.")

    print(f"Reading official calendar: {CALENDAR_URL}")
    print(
        f"Requested week: {week_start} through "
        f"{week_start + timedelta(days=6)}"
    )

    calendar_text = fetch_calendar_text()
    events = parse_calendar(calendar_text, week_start)

    if not events:
        raise RuntimeError(
            "No dated SWRC pool events were parsed. The calendar may "
            "have changed, or dates may not be associated with event "
            "lines. No schedules were created."
        )

    print(f"Parsed {len(events)} dated SWRC pool event(s).")
    OUTPUT.mkdir(parents=True, exist_ok=True)

    # Calculate the same time font size before creating either PDF.
    shared_time_font_size = calculate_shared_time_font_size(
        selected,
        week_start,
        events,
    )

    for pool_key in selected:
        week_events = pool_events_for_week(
            events,
            pool_key,
            week_start,
        )

        print(
            f"{POOLS[pool_key]['label']}: "
            f"{len(week_events)} event(s) in requested week"
        )

        pptx_path = update_template(
            pool_key,
            week_start,
            events,
            shared_time_font_size,
        )

        pdf_path = export_pdf(pptx_path)

        print(f"Created PowerPoint: {pptx_path}")
        print(f"Created PDF: {pdf_path}")

    print(
        "Generation finished. Review both PDFs before posting."
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
