
#!/usr/bin/env python3
"""
Rutgers SWRC Pool Schedule Generator

Reads pool-only hours from the official Rutgers Recreation calendar,
updates the supplied PowerPoint templates, and exports separate PDFs.

Safety rules:
- Never invent pool hours.
- Never interpret a missing event as confirmed closed hours.
- Stop if the calendar cannot be parsed or a template cannot be updated
  confidently.
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
from playwright.sync_api import sync_playwright


CALENDAR_URL = (
    "https://recreation.rutgers.edu/"
    "recreation-list-calendar?calendar_filter=122"
)

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"

POOLS = {
    "olympic": {
        "label": "Olympic Pool",
        "": ROOT / "Olympic Pool Spring 2026 Hours.pptx",
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

DATE_PATTERNS = [
    re.compile(
        r"\b(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+"
        r"(\d{1,2})(?:,?\s+(\d{4}))?\b",
        re.I,
    ),
    re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b"),
]

WEEKDAYS = [
    "Monday", "Tuesday", "Wednesday", "Thursday",
    "Friday", "Saturday", "Sunday",
]


def normalize_time(value):
    """Normalize calendar time text to a readable format."""
    value = re.sub(r"\s+", " ", value.strip().lower())
    value = value.replace(".", "")
    parsed = datetime.strptime(value, "%I:%M %p")
    return parsed.strftime("%-I:%M %p")


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
                int(iso.group(1)), int(iso.group(2)), int(iso.group(3))
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
        month = datetime.strptime(month_match.group(1), "%B").month
        day = int(month_match.group(2))
        year = int(month_match.group(3) or default_year)
        try:
            return date(year, month, day)
        except ValueError:
            return None

    return None


def fetch_calendar_text():
    """
    First try the page's rendered text, because the calendar may be
    populated by JavaScript. Fall back to the directly returned HTML.
    """
    rendered_text = ""

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
            rendered_text = page.locator("body").inner_text(timeout=15000)
            browser.close()
    except Exception as exc:
        print(
            "Notice: rendered-calendar retrieval failed; "
            f"trying direct HTML instead ({exc})"
        )

    if rendered_text and (
        "SWRC Olympic Pool" in rendered_text
        or "SWRC Fixed Depth Patio Pool" in rendered_text
        or "SWRC Patio Pool" in rendered_text
    ):
        return rendered_text

    response = requests.get(
        CALENDAR_URL,
        timeout=45,
        headers={"User-Agent": "RutgersPoolScheduleGenerator/1.0"},
    )
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    direct_text = soup.get_text("\n", strip=True)

    if direct_text:
        return direct_text

    raise RuntimeError("The Rutgers calendar returned no readable text.")


def parse_calendar(text, week_start):
    """
    Parse dated calendar sections and pool event lines.

    Rutgers entries commonly look like:
    07:00 am - 10:00 pm | | Recreation Swim -
    SWRC Fixed Depth Patio Pool

    If an event has no reliably associated date, it is not assigned
    to a day by guessing.
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

        # If the date is embedded in the same line as the event,
        # use it. Otherwise use the latest recognized date heading.
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

    # Deduplicate identical events that may appear in overlapping
    # calendar sections.
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
        event for event in events
        if event["pool"] == pool_key
        and week_start <= event["date"] <= end_date
    ]


def replace_text_in_shape(shape, replacement):
    """Replace a shape's text while retaining its first run's styling."""
    if not shape.has_text_frame:
        return False

    paragraph_runs = [
        run
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
    ]

    if paragraph_runs:
        paragraph_runs[0].text = replacement
        for run in paragraph_runs[1:]:
            run.text = ""
    else:
        shape.text_frame.text = replacement

    return True




def update_template(pool_key, week_start, events):
    """
    Update the seven existing hours text boxes in the supplied template.
    """
    if pool_key == "olympic":
        label = "Olympic Pool"
        template = ROOT / "Olympic Pool Spring 2026 Hours.pptx"
    elif pool_key == "patio":
        label = "Patio Pool"
        template = ROOT / "Patio Pool Spring 2026 Hours.pptx"
    else:
        raise RuntimeError(f"Unknown pool key: {pool_key}")

    if not template.is_file():
        raise RuntimeError(f"Template not found: {template.name}")

    presentation = Presentation(str(template))

    if len(presentation.slides) != 1:
        raise RuntimeError(
            f"Expected one slide in {template.name}; "
            f"found {len(presentation.slides)}."
        )

    slide = presentation.slides[0]

    # The seven editable hours boxes are identified by their vertical
    # positions. The background is a picture, not a text box.
    hour_boxes = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        # Hours boxes begin around 2.2 inches from the top and end
        # around 7.3 inches. This excludes the title and footer.
        top_inches = shape.top / 914400

        
    if 2.0 <= top_inches <= 7.3 and shape.text.strip():
        hour_boxes.append(shape)


    hour_boxes.sort(key=lambda shape: shape.top)

    if len(hour_boxes) != 7:
        details = [
            f"{shape.name}: top={shape.top / 914400:.2f}in, "
            f"text={shape.text.strip()!r}"
            for shape in hour_boxes
        ]
        raise RuntimeError(
            f"Expected 7 hours boxes in {template.name}, "
            f"found {len(hour_boxes)}. "
            f"Detected boxes: {details}"
        )

    days = [week_start + timedelta(days=i) for i in range(7)]

    events_by_date = {}
    for event in events:
        events_by_date.setdefault(event["date"], []).append(event)

    for day, shape in zip(days, hour_boxes):
        day_events = events_by_date.get(day, [])

        # Only use events already filtered for this pool.
        hours = sorted({
            event["hours"]
            for event in day_events
            if event.get("pool") == pool_key
        })

        if len(hours) == 1:
            display = hours[0]
        elif len(hours) > 1:
            display = "\n".join(hours) + "\nVERIFY: multiple entries"
        else:
            display = "VERIFY: hours not found"

        # Keep the existing first-run formatting where possible.
        paragraph = (
            shape.text_frame.paragraphs[0]
            if shape.text_frame.paragraphs
            else None
        )

        if paragraph and paragraph.runs:
            paragraph.runs[0].text = display
            for run in paragraph.runs[1:]:
                run.text = ""

            # Clear any remaining paragraphs without removing the
            # text frame or changing the shape's geometry.
            for extra in shape.text_frame.paragraphs[1:]:
                for run in extra.runs:
                    run.text = ""
        else:
            shape.text_frame.text = display

    OUTPUT.mkdir(parents=True, exist_ok=True)
    safe_pool_name = label.replace(" ", "_")
    pptx_path = OUTPUT / f"{safe_pool_name}_{week_start.isoformat()}.pptx"

    presentation.save(str(pptx_path))
    return pptx_path



def export_pdf(pptx_path):
    """Export a PowerPoint file to PDF using LibreOffice."""
    result = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(OUTPUT),
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
    print(f"Requested week: {week_start} through "
          f"{week_start + timedelta(days=6)}")

    text = fetch_calendar_text()
    events = parse_calendar(text, week_start)

    # Validate that the page yielded identifiable SWRC pool entries
    # before attempting to edit any template.
    if not events:
        raise RuntimeError(
            "No dated SWRC pool events were parsed. "
            "The calendar layout may have changed, or dates may not "
            "be associated with event lines. No schedules were created."
        )

    print(f"Parsed {len(events)} dated SWRC pool event(s).")
    OUTPUT.mkdir(parents=True, exist_ok=True)

    for pool_key in selected:
        week_events = pool_events_for_week(
            events, pool_key, week_start
        )

        print(
            f"{POOLS[pool_key]['label']}: "
            f"{len(week_events)} event(s) in requested week"
        )

        pptx_path = update_template(
            pool_key, week_start, week_events
        )
        pdf_path = export_pdf(pptx_path)

        print(f"Created PowerPoint: {pptx_path}")
        print(f"Created PDF: {pdf_path}")

    print(
        "Generation finished. Review every day marked "
        "'VERIFY: hours not found on calendar' before using the PDFs."
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
