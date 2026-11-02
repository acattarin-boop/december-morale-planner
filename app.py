"""Team morale event planner.

Everything that changes per event lives in the Google Sheet, not in this file:
  config tab  - title, date range, deadlines, lock behaviour
  roster tab  - the names people can pick from

Streamlit secrets required:
  sheet_id             the ID from the Google Sheet URL
  gcp_service_account  the full service account JSON, as a string
  initial_roster       optional comma-separated names, used once to fill an empty roster tab
"""
import calendar
import html
import json
import logging
import threading
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

log = logging.getLogger("planner")

# --------------------------------------------------------------------------
# Sheet layout
# --------------------------------------------------------------------------
IDEA_HEADERS = ["Submitted By", "Title", "Description", "Link", "Est. Price/Person"]
RANKS = [("1st", 3), ("2nd", 2), ("3rd", 1)]
VOTE_HEADERS = ["Name"] + [f"{kind} {rank}" for kind in ("Event", "Dining") for rank, _ in RANKS]

# Required leading headers per tab. The availability tab gains one column per
# date, added by the app the first time that date is saved.
TABS = {
    "config": ["Setting", "Value", "Notes"],
    "roster": ["Name"],
    "availability": ["Name"],
    "events": IDEA_HEADERS,
    "dining": IDEA_HEADERS,
    "votes": VOTE_HEADERS,
}

DEFAULT_CONFIG = [
    ("event_title", "December Morale Event", "Shown at the top of the app"),
    ("first_date", "2026-11-30", "First date people can choose (YYYY-MM-DD)"),
    ("last_date", "2026-12-31", "Last date people can choose (YYYY-MM-DD)"),
    ("weekdays_only", "yes", "yes = Monday to Friday only, no = include weekends"),
    ("skip_dates",
     "2026-12-21, 2026-12-22, 2026-12-23, 2026-12-24, 2026-12-25, 2026-12-28, 2026-12-29",
     "Dates to leave out, comma separated (YYYY-MM-DD)"),
    ("ideas_deadline", "2026-11-16", "Last day to add event or dining ideas (YYYY-MM-DD)"),
    ("voting_deadline", "2026-11-23", "Last day to vote (YYYY-MM-DD)"),
    ("deadline", "2026-11-23", "Last day to enter availability (YYYY-MM-DD)"),
    ("timezone", "America/Los_Angeles", "Deadlines are 11:59 PM in this time zone"),
    ("lock_after_deadline", "yes", "yes = closed after each deadline, no = just show the dates"),
]

YES, NO = "Yes", "No"
# Times offered on a Yes day. The short form is used on the calendar's name chips.
SLOTS = {"8am to noon": "8-12", "noon to 2pm": "12-2", "2pm to 5pm": "2-5"}
NO_PICK = "No pick"
TITLE_LIMIT = 80
NAME_LIMIT = 40


class SetupError(Exception):
    """A problem the organizer can fix in the sheet or in Streamlit secrets."""


# --------------------------------------------------------------------------
# Google Sheets connection
# --------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_workbook():
    for secret in ("sheet_id", "gcp_service_account"):
        if secret not in st.secrets:
            raise SetupError(f"Streamlit secret '{secret}' is missing. Add it under Settings > Secrets.")
    creds = Credentials.from_service_account_info(
        json.loads(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    client = gspread.authorize(creds, http_client=gspread.BackOffHTTPClient)
    try:
        return client.open_by_key(st.secrets["sheet_id"])
    except (gspread.exceptions.SpreadsheetNotFound, PermissionError) as err:
        raise SetupError(
            "The app can't open the Google Sheet. Check the sheet_id secret and share "
            "the sheet with the service account email as Editor."
        ) from err


@st.cache_resource(show_spinner=False)
def write_lock():
    """One save at a time, so two people can't race for the same row."""
    return threading.Lock()


@st.cache_resource(show_spinner=False)
def ensure_tabs():
    """Create any missing tab, write its headers, and fill in first-run defaults."""
    book = get_workbook()
    existing = {ws.title: ws for ws in book.worksheets()}
    for tab, headers in TABS.items():
        ws = existing.get(tab) or book.add_worksheet(tab, rows=200, cols=max(len(headers), 40))
        first_row = [h.strip() for h in ws.row_values(1)]
        if not first_row:
            ws.update(values=[headers], range_name="A1", value_input_option="RAW")
            seed_tab(ws, tab)
        elif first_row[: len(headers)] != headers:
            raise SetupError(
                f"The '{tab}' tab starts with {first_row[:len(headers)]} "
                f"but the app expects {headers}. Fix row 1 of that tab."
            )
        elif tab == "config":
            add_missing_settings(ws)
    return True


def add_missing_settings(ws):
    """Add settings introduced after the sheet was created. Existing rows are left alone."""
    present = {value.strip() for value in ws.col_values(1)}
    missing = [list(row) for row in DEFAULT_CONFIG if row[0] not in present]
    if missing:
        ws.append_rows(missing, value_input_option="RAW", table_range="A1")


def seed_tab(ws, tab):
    if tab == "config":
        ws.append_rows([list(row) for row in DEFAULT_CONFIG], value_input_option="RAW", table_range="A1")
    elif tab == "roster":
        raw = st.secrets.get("initial_roster", "")
        names = raw if isinstance(raw, (list, tuple)) else str(raw).split(",")
        names = [str(n).strip() for n in names if str(n).strip()]
        if names:
            ws.append_rows([[n] for n in names], value_input_option="RAW", table_range="A1")


def rows_to_frame(rows):
    """Row 1 is the header. Short rows are padded and blank header columns dropped."""
    if not rows:
        return pd.DataFrame()
    headers = [str(h).strip() for h in rows[0]]
    body = [(list(r) + [""] * len(headers))[: len(headers)] for r in rows[1:]]
    frame = pd.DataFrame(body, columns=headers, dtype=str)
    return frame.loc[:, [h for h in frame.columns if h]]


@st.cache_data(ttl=60, show_spinner=False)
def load_all():
    """Every tab in a single API call. Shared by all visitors for 60 seconds."""
    ensure_tabs()
    response = get_workbook().values_batch_get(list(TABS))
    return {
        tab: rows_to_frame(part.get("values", []))
        for tab, part in zip(TABS, response["valueRanges"])
    }


def same(a, b):
    return str(a).strip().casefold() == str(b).strip().casefold()


def upsert(tab, key_header, key, values):
    """Find the row in the live sheet, then update it, or append if it's new.

    The row is always looked up immediately before writing, and cells are matched
    by column header rather than position. Columns not in `values` are left alone.
    """
    with write_lock():
        ws = get_workbook().worksheet(tab)
        rows = ws.get_all_values()
        headers = [h.strip() for h in rows[0]]

        missing = [h for h in values if h not in headers]
        if missing:
            headers += missing
            if len(headers) > ws.col_count:
                ws.add_cols(len(headers) - ws.col_count)
            ws.update(values=[headers], range_name="A1", value_input_option="RAW")

        key_col = headers.index(key_header)
        row_number = next(
            (i for i, row in enumerate(rows[1:], start=2)
             if len(row) > key_col and same(row[key_col], key)),
            None,
        )
        if row_number is None:
            ws.append_row([values.get(h, "") for h in headers],
                          value_input_option="RAW", table_range="A1")
        else:
            old = rows[row_number - 1] + [""] * len(headers)
            row = [values.get(h, old[i]) for i, h in enumerate(headers)]
            end = rowcol_to_a1(row_number, len(headers))
            ws.update(values=[row], range_name=f"A{row_number}:{end}", value_input_option="RAW")
    load_all.clear()


def append_unique(tab, key_header, key, row):
    """Append unless the key already exists. Returns False for a duplicate."""
    with write_lock():
        ws = get_workbook().worksheet(tab)
        rows = ws.get_all_values()
        key_col = [h.strip() for h in rows[0]].index(key_header)
        if any(len(r) > key_col and same(r[key_col], key) for r in rows[1:]):
            return False
        ws.append_row(row, value_input_option="RAW", table_range="A1")
    load_all.clear()
    return True


# --------------------------------------------------------------------------
# Settings and dates
# --------------------------------------------------------------------------
def parse_date(setting, value):
    for pattern in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(value).strip(), pattern).date()
        except ValueError:
            pass
    raise SetupError(f"In the config tab, {setting} should look like 2026-12-01. It is '{value}'.")


def is_yes(value):
    return str(value).strip().lower() in ("yes", "y", "true", "1")


@dataclass(frozen=True)
class Settings:
    event_title: str
    first_date: date
    last_date: date
    weekdays_only: bool
    skip_dates: tuple
    ideas_deadline: datetime
    voting_deadline: datetime
    availability_deadline: datetime
    lock_after_deadline: bool

    def still_open(self, deadline):
        return datetime.now(deadline.tzinfo) <= deadline or not self.lock_after_deadline

    @property
    def ideas_open(self):
        return self.still_open(self.ideas_deadline)

    @property
    def voting_open(self):
        return self.still_open(self.voting_deadline)

    @property
    def availability_open(self):
        return self.still_open(self.availability_deadline)

    @property
    def any_open(self):
        return self.ideas_open or self.voting_open or self.availability_open

    def days(self):
        result = []
        day = self.first_date
        while day <= self.last_date:
            if day not in self.skip_dates and not (self.weekdays_only and day.weekday() >= 5):
                result.append(day)
            day += timedelta(days=1)
        return result


def parse_settings(frame):
    raw = {name: value for name, value, _ in DEFAULT_CONFIG}
    if not frame.empty:
        for _, row in frame.iterrows():
            name = row["Setting"].strip()
            if name:
                raw[name] = row["Value"].strip()

    try:
        zone = ZoneInfo(raw["timezone"])
    except Exception as err:
        raise SetupError(f"In the config tab, '{raw['timezone']}' isn't a time zone name.") from err

    first_date = parse_date("first_date", raw["first_date"])
    last_date = parse_date("last_date", raw["last_date"])
    if last_date < first_date:
        raise SetupError("In the config tab, last_date is earlier than first_date.")

    skip = tuple(parse_date("skip_dates", part) for part in raw["skip_dates"].split(",") if part.strip())

    def end_of(setting):
        return datetime.combine(parse_date(setting, raw[setting]), time(23, 59, 59), tzinfo=zone)

    return Settings(
        event_title=raw["event_title"] or "Morale Event",
        first_date=first_date,
        last_date=last_date,
        weekdays_only=is_yes(raw["weekdays_only"]),
        skip_dates=skip,
        ideas_deadline=end_of("ideas_deadline"),
        voting_deadline=end_of("voting_deadline"),
        availability_deadline=end_of("deadline"),
        lock_after_deadline=is_yes(raw["lock_after_deadline"]),
    )


def day_key(day):
    """The one and only format used for date columns in the sheet."""
    return day.isoformat()


def day_label(day):
    return f"{day.strftime('%a %b')} {day.day}"


def deadline_text(settings):
    """One line with each cutoff, plus a countdown on the next one to close."""
    steps = [
        ("Ideas", "close", settings.ideas_deadline, settings.ideas_open),
        ("Voting", "closes", settings.voting_deadline, settings.voting_open),
        ("Availability", "closes", settings.availability_deadline, settings.availability_open),
    ]
    parts = []
    countdown_used = False
    for label, verb, moment, is_open in sorted(steps, key=lambda step: step[2]):
        when = f"{moment.strftime('%a, %b')} {moment.day}"
        if not is_open:
            parts.append(f"{label} closed {when}")
            continue
        text = f"{label} {verb} {when}"
        if not countdown_used:
            left = (moment.date() - datetime.now(moment.tzinfo).date()).days
            if left >= 0:
                countdown = "today" if left == 0 else "1 day left" if left == 1 else f"{left} days left"
                text += f" ({countdown})"
                countdown_used = True
        parts.append(text)
    zone = settings.availability_deadline.strftime("%Z")
    return " · ".join(parts) + f". All at 11:59 PM {zone}."


# --------------------------------------------------------------------------
# Reading answers
# --------------------------------------------------------------------------
def roster_names(frame):
    seen = {}
    if not frame.empty:
        for value in frame["Name"]:
            name = value.strip()
            if name:
                seen.setdefault(name.casefold(), name)
    return sorted(seen.values(), key=str.casefold)


def parse_answer(cell):
    """Sheet cell to (status, slots), or None if unanswered.

    "Yes" means every time slot, "Yes: 8am to noon, 2pm to 5pm" means only those.
    """
    text = str(cell).strip()
    if text == NO:
        return (NO, ())
    if text == YES or text.startswith(f"{YES}:"):
        chosen = [part.strip() for part in text.partition(":")[2].split(",")]
        return (YES, tuple(s for s in SLOTS if s in chosen) or tuple(SLOTS))
    return None


def format_answer(status, slots):
    if status == NO:
        return NO
    if status != YES:
        return ""
    picked = [s for s in SLOTS if s in slots]
    return YES if len(picked) == len(SLOTS) else f"{YES}: {', '.join(picked)}"


def availability_by_person(frame, names, days):
    """{name: {day_key: (status, slots)}} for people on the roster. First row per name wins."""
    answers = {}
    if frame.empty:
        return answers
    canonical = {n.casefold(): n for n in names}
    keys = [day_key(d) for d in days]
    for _, row in frame.iterrows():
        name = canonical.get(row["Name"].strip().casefold())
        if name and name not in answers:
            parsed = {k: parse_answer(row[k]) for k in keys if k in row}
            answers[name] = {k: answer for k, answer in parsed.items() if answer}
    return answers


def votes_by_person(frame, names):
    votes = {}
    if frame.empty:
        return votes
    canonical = {n.casefold(): n for n in names}
    for _, row in frame.iterrows():
        name = canonical.get(row["Name"].strip().casefold())
        if name and name not in votes:
            votes[name] = {h: row[h].strip() for h in VOTE_HEADERS[1:] if h in row}
    return votes


def idea_titles(frame):
    if frame.empty:
        return []
    return [t.strip() for t in frame["Title"] if t.strip()]


def tally(votes, kind, titles):
    points = {t: {"Option": t, "Points": 0, "1st-place votes": 0} for t in titles}
    for vote in votes.values():
        for rank, weight in RANKS:
            pick = vote.get(f"{kind} {rank}", "")
            if pick in points:
                points[pick]["Points"] += weight
                points[pick]["1st-place votes"] += 1 if rank == "1st" else 0
    table = sorted(points.values(), key=lambda r: (-r["Points"], -r["1st-place votes"], r["Option"].casefold()))
    return pd.DataFrame(table, columns=["Option", "Points", "1st-place votes"])


def people_for(answers, names, key, status):
    return [n for n in names if answers.get(n, {}).get(key, (None,))[0] == status]


def date_summary(answers, names, days):
    """One row per date. The best dates are those where the most people share a time."""
    rows = []
    for day in days:
        key = day_key(day)
        yes = people_for(answers, names, key, YES)
        no = people_for(answers, names, key, NO)
        slot_counts = {s: sum(s in answers[n][key][1] for n in yes) for s in SLOTS}
        rows.append({
            "Date": day_label(day),
            **slot_counts,
            "Can't make it": len(no),
            "No answer": len(names) - len(yes) - len(no),
            "Who can't": ", ".join(no),
            "_best": max(slot_counts.values()),
            "_yes": len(yes),
        })
    rows.sort(key=lambda r: (-r["_best"], -r["_yes"], r["Can't make it"]))
    columns = ["Date", *SLOTS, "Can't make it", "No answer", "Who can't"]
    return pd.DataFrame(rows, columns=columns)


# --------------------------------------------------------------------------
# Page helpers
# --------------------------------------------------------------------------
def esc(text):
    """Escape for an HTML block, including $ so Streamlit doesn't render math."""
    return html.escape(str(text)).replace("$", "&#36;").replace("\n", " ")


def flash(message):
    st.session_state["flash"] = message


def show_save_error(err):
    log.exception("Save failed")
    st.error(f"That didn't save because Google Sheets returned an error ({describe(err)}). Please try again.")


def describe(err):
    response = getattr(err, "response", None)
    return f"HTTP {response.status_code}" if response is not None else type(err).__name__


def apply_style():
    st.markdown(
        """
        <style>
        :root { --pine: #1f5c45; --cranberry: #a8202e; --gold: #c9a24b; --line: #e6dfd0; }
        .block-container { max-width: 1100px; padding-top: 2.5rem; }
        h1, h2, h3, h4 { color: var(--pine) !important; }
        h1 { font-family: Georgia, "Times New Roman", serif; font-weight: 700; }
        .holiday-ribbon {
            height: 10px; border-radius: 5px; margin: .75rem 0 1.25rem;
            background: repeating-linear-gradient(115deg, var(--cranberry) 0 16px,
                #fff 16px 24px, var(--pine) 24px 40px, #fff 40px 48px);
        }
        .holiday-mark { color: var(--gold); font-size: .7em; vertical-align: middle; margin: 0 .2em; }
        .stTabs [data-baseweb="tab"] { font-weight: 600; }
        [data-testid="stColumn"] [data-testid="stVerticalBlockBorderWrapper"] {
            border-color: var(--line); background: #fffefb;
        }
        [role="radiogroup"] [role="radio"]:nth-of-type(2)[aria-checked="true"] {
            color: var(--cranberry); border-color: var(--cranberry); background: #fbecec;
        }
        .day-top { display: flex; justify-content: space-between; align-items: baseline; gap: .4rem; }
        .day-name { font-weight: 600; color: var(--pine); }
        .day-count { font-size: .75rem; color: #5f6368; white-space: nowrap; }
        .chips { display: flex; flex-wrap: wrap; gap: 3px; padding-bottom: .9rem; }
        .chip { font-size: .7rem; padding: 0 5px; border-radius: 4px; line-height: 1.5; }
        .chip-yes { background: #e4f0e8; color: #1f5c45; }
        .chip-no { background: #fae6e6; color: #a8202e; }
        .idea { border: 1px solid var(--line); border-left: 3px solid var(--gold); border-radius: 8px;
                padding: .55rem .8rem; margin-bottom: .5rem; background: #fffefb; }
        .idea-head { display: flex; justify-content: space-between; gap: .75rem; }
        .idea-title { font-weight: 600; }
        .idea-title a { color: var(--pine); }
        .idea-price { color: #5f6368; font-size: .85rem; white-space: nowrap; }
        .idea-by { color: #80868b; font-size: .75rem; }
        .idea-desc { color: #444; font-size: .85rem; margin-top: .15rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def calendar_heading(days):
    first, last = days[0], days[-1]
    if (first.year, first.month) == (last.year, last.month):
        return f"{calendar.month_name[first.month]} {first.year}"
    return f"{calendar.month_name[first.month]} to {calendar.month_name[last.month]} {last.year}"


def calendar_grid(settings, days, render_day):
    """One continuous grid, Monday first, with a bordered cell per selectable day."""
    if not days:
        st.warning("No dates are available. Check the date settings in the config tab.")
        return
    selectable = set(days)
    columns = 5 if settings.weekdays_only else 7
    st.markdown(f"#### {calendar_heading(days)}")
    monday = days[0] - timedelta(days=days[0].weekday())
    while monday <= days[-1]:
        week = [monday + timedelta(days=offset) for offset in range(columns)]
        if selectable.intersection(week):
            for column, day in zip(st.columns(columns, gap="small"), week):
                if day in selectable:
                    with column, st.container(border=True):
                        render_day(day)
        monday += timedelta(days=7)


def yes_chip(name, slots):
    """A partial day shows the person's times next to their name."""
    times = "" if len(slots) == len(SLOTS) else " " + ", ".join(SLOTS[s] for s in slots)
    return f'<span class="chip chip-yes">{esc(name)}{esc(times)}</span>'


def day_html(day, answers, names):
    key = day_key(day)
    yes = people_for(answers, names, key, YES)
    no = people_for(answers, names, key, NO)
    chips = "".join(yes_chip(n, answers[n][key][1]) for n in yes)
    chips += "".join(f'<span class="chip chip-no">{esc(n)}</span>' for n in no)
    count = f"{len(yes)} of {len(names)} yes" if yes or no else ""
    top = f'<div class="day-top"><span class="day-name">{esc(day_label(day))}</span>' \
          f'<span class="day-count">{count}</span></div>'
    return top, f'<div class="chips">{chips}</div>'


# --------------------------------------------------------------------------
# Tabs
# --------------------------------------------------------------------------
def availability_tab(settings, me, names, answers, days):
    st.subheader("When can you make it?")

    def read_only_day(day):
        top, chips = day_html(day, answers, names)
        st.markdown(top + chips, unsafe_allow_html=True)

    if not settings.availability_open:
        st.info("Availability is closed. This is what everyone said.")
        calendar_grid(settings, days, read_only_day)
    elif not me:
        st.info("Choose your name above to fill this in.")
        calendar_grid(settings, days, read_only_day)
    else:
        st.caption("Choose Yes or No for each date. Yes selects all three times, and you can tap "
                   "a time to remove it. Press your Yes or No again to clear the date. "
                   "Press Save at the bottom when you're done.")
        availability_editor(settings, me, names, answers, days)


def on_day_change(day_state, slots_state):
    if st.session_state[day_state] == YES:
        st.session_state[slots_state] = list(SLOTS)


def on_slots_change(day_state, slots_state):
    """Removing every time from a Yes day means the person can't make that day."""
    if not st.session_state[slots_state]:
        st.session_state[day_state] = NO


@st.fragment
def availability_editor(settings, me, names, answers, days):
    """Runs as a fragment so tapping a date doesn't redraw the rest of the page."""
    mine = answers.get(me, {})
    choices = {}

    def editable_day(day):
        key = day_key(day)
        day_state, slots_state = f"day_{me}_{key}", f"slots_{me}_{key}"
        status, slots = mine.get(key, (None, ()))
        if day_state not in st.session_state:
            st.session_state[day_state] = status
        if st.session_state[day_state] == YES and slots_state not in st.session_state:
            st.session_state[slots_state] = list(slots or SLOTS)

        top, chips = day_html(day, answers, names)
        st.markdown(top, unsafe_allow_html=True)
        picked = st.segmented_control(
            day_label(day), [YES, NO], key=day_state, label_visibility="collapsed", width="stretch",
            on_change=on_day_change, args=(day_state, slots_state),
        )
        picked_slots = ()
        if picked == YES:
            picked_slots = st.pills(
                f"Times on {day_label(day)}", list(SLOTS), selection_mode="multi", key=slots_state,
                label_visibility="collapsed", on_change=on_slots_change, args=(day_state, slots_state),
            )
        choices[key] = format_answer(picked, picked_slots)
        st.markdown(chips, unsafe_allow_html=True)

    calendar_grid(settings, days, editable_day)

    if st.button("Save my availability", type="primary") and settings.availability_open:
        try:
            upsert("availability", "Name", me, {"Name": me, **choices})
            flash("Availability saved")
            st.rerun()
        except gspread.exceptions.GSpreadException as err:
            show_save_error(err)


def idea_card(row):
    title = esc(row["Title"].strip())
    link = row["Link"].strip()
    if link.lower().startswith(("http://", "https://")):
        title = f'<a href="{html.escape(link, quote=True)}" target="_blank" rel="noopener">{title}</a>'
    price = row["Est. Price/Person"].strip()
    if price[:1].isdigit():
        price = f"${price}"
    price_html = f'<span class="idea-price">{esc(price)} per person</span>' if price \
        else '<span class="idea-price">Cost not listed</span>'
    description = row["Description"].strip()
    description_html = f'<div class="idea-desc">{esc(description)}</div>' if description else ""
    by = esc(row["Submitted By"].strip() or "Unknown")
    return (f'<div class="idea"><div class="idea-head"><span class="idea-title">{title}</span>{price_html}</div>'
            f'{description_html}<div class="idea-by">Suggested by {by}</div></div>')


def ideas_tab(settings, me, sheet_tab, label, frame):
    st.subheader(f"{label} ideas")
    left, right = st.columns([3, 2], gap="large")

    with left:
        rows = [] if frame.empty else [r for _, r in frame.iterrows() if r["Title"].strip()]
        if not rows:
            st.info(f"No {label.lower()} ideas yet.")
        for row in rows:
            st.markdown(idea_card(row), unsafe_allow_html=True)

    with right:
        if not settings.ideas_open:
            st.info("New ideas are closed. You can still vote on these until voting closes."
                    if settings.voting_open else "Ideas and voting are closed.")
        elif not me:
            st.info("Choose your name above to suggest an idea.")
        else:
            st.markdown(f"**Suggest {'an' if label == 'Event' else 'a'} {label.lower()} idea**")
            version = st.session_state.get(f"idea_form_{sheet_tab}", 0)
            field = f"{sheet_tab}_{version}"
            with st.form(f"idea_{field}", border=False):
                title = st.text_input("Title (required)", max_chars=TITLE_LIMIT, key=f"title_{field}")
                price = st.text_input("Cost per person (required)", placeholder="25, or Free",
                                      key=f"price_{field}")
                description = st.text_area("Description (optional)", height=80, key=f"desc_{field}")
                link = st.text_input("Link (optional)", placeholder="https://", key=f"link_{field}")
                submitted = st.form_submit_button("Add idea", type="primary")
            if submitted and settings.ideas_open:
                add_idea(sheet_tab, me, title, description, link, price)


def add_idea(sheet_tab, me, title, description, link, price):
    title, link, price = title.strip(), link.strip(), price.strip()
    if not title:
        st.warning("Add a title first.")
    elif not price:
        st.warning("Add the cost per person. Use an estimate, or Free if there's no cost.")
    elif link and not link.lower().startswith(("http://", "https://")):
        st.warning("Links need to start with http:// or https://.")
    else:
        row = [me, title, description.strip(), link, price]
        try:
            if append_unique(sheet_tab, "Title", title, row):
                version_key = f"idea_form_{sheet_tab}"
                st.session_state[version_key] = st.session_state.get(version_key, 0) + 1
                flash("Idea added")
                st.rerun()
            st.warning(f"'{title}' is already on the list.")
        except gspread.exceptions.GSpreadException as err:
            show_save_error(err)


def vote_picks(me, kind, titles, current):
    picks = {}
    options = [NO_PICK] + titles
    for (rank, weight), medal in zip(RANKS, ("1st choice", "2nd choice", "3rd choice")):
        header = f"{kind} {rank}"
        saved = current.get(header, "")
        unit = "point" if weight == 1 else "points"
        picks[header] = st.selectbox(
            f"{medal} ({weight} {unit})", options,
            index=options.index(saved) if saved in options else 0,
            key=f"vote_{me}_{header}",
        )
    return picks


def votes_tab(settings, me, votes, event_titles, dining_titles):
    st.subheader("Vote")
    st.caption("Voting is optional. Pick as many or as few as you like, in either list, "
               "and change your vote until the deadline.")
    if not settings.voting_open:
        st.info("Voting is closed. The Results tab has the final scores.")
        return
    if not me:
        st.info("Choose your name above to vote.")
        return
    if not event_titles and not dining_titles:
        st.info("There's nothing to vote on yet. Add ideas in the Events and Dining tabs first.")
        return

    current = votes.get(me, {})
    with st.form(f"votes_{me}", border=False):
        left, right = st.columns(2, gap="large")
        picks = {}
        with left:
            st.markdown("**Events**")
            if event_titles:
                picks.update(vote_picks(me, "Event", event_titles, current))
            else:
                st.info("No event ideas yet.")
        with right:
            st.markdown("**Dining**")
            if dining_titles:
                picks.update(vote_picks(me, "Dining", dining_titles, current))
            else:
                st.info("No dining ideas yet.")
        saved = st.form_submit_button("Save my votes", type="primary")

    if saved and settings.voting_open:
        chosen = {h: ("" if v == NO_PICK else v) for h, v in picks.items()}
        repeated = [
            kind for kind in ("Event", "Dining")
            if len(p := [v for h, v in chosen.items() if h.startswith(kind) and v]) != len(set(p))
        ]
        if repeated:
            st.warning("Each choice can only be used once per list. Change the repeated pick.")
            return
        try:
            upsert("votes", "Name", me, {"Name": me, **chosen})
            flash("Votes saved")
            st.rerun()
        except gspread.exceptions.GSpreadException as err:
            show_save_error(err)


def results_tab(names, answers, votes, days, event_titles, dining_titles):
    st.subheader("Results so far")
    responded = [n for n in names if answers.get(n)]
    voted = [n for n in names if any(votes.get(n, {}).values())]

    first, second = st.columns(2)
    first.metric("Availability answered", f"{len(responded)} of {len(names)}")
    second.metric("Votes in", f"{len(voted)} of {len(names)}")

    waiting_dates = [n for n in names if n not in responded]
    waiting_votes = [n for n in names if n not in voted]
    if waiting_dates:
        st.caption(f"Still need availability from: {', '.join(waiting_dates)}")
    if waiting_votes:
        st.caption(f"Still need votes from: {', '.join(waiting_votes)}")

    st.markdown("#### Best dates")
    st.caption("Each time column shows how many people can make that time. "
               "Dates are ranked by the time that suits the most people.")
    if responded:
        st.dataframe(date_summary(answers, names, days).head(5), hide_index=True, width="stretch")
    else:
        st.info("No availability saved yet.")

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("#### Events")
        if event_titles:
            st.dataframe(tally(votes, "Event", event_titles), hide_index=True, width="stretch")
        else:
            st.info("No event ideas yet.")
    with right:
        st.markdown("#### Dining")
        if dining_titles:
            st.dataframe(tally(votes, "Dining", dining_titles), hide_index=True, width="stretch")
        else:
            st.info("No dining ideas yet.")


def who_are_you(settings, names):
    if "pending_me" in st.session_state:
        st.session_state["me"] = st.session_state.pop("pending_me")

    me = st.selectbox("Who are you?", names, index=None, placeholder="Choose your name", key="me")

    if settings.any_open:
        with st.expander("Don't see your name?"):
            with st.form("add_name", clear_on_submit=True, border=False):
                new_name = st.text_input("Your name", max_chars=NAME_LIMIT)
                submitted = st.form_submit_button("Add me")
            if submitted:
                name = " ".join(new_name.split())
                if not name:
                    st.warning("Type your name first.")
                else:
                    try:
                        if append_unique("roster", "Name", name, [name]):
                            st.session_state["pending_me"] = name
                            flash(f"Added {name}")
                            st.rerun()
                        st.warning(f"{name} is already on the list.")
                    except gspread.exceptions.GSpreadException as err:
                        show_save_error(err)
    return me


# --------------------------------------------------------------------------
# App
# --------------------------------------------------------------------------
def load_or_stop():
    try:
        data = load_all()
        return data, parse_settings(data["config"])
    except SetupError as err:
        st.error(str(err))
    except gspread.exceptions.GSpreadException as err:
        log.exception("Load failed")
        st.error(f"Google Sheets isn't responding ({describe(err)}). Wait a minute and refresh.")
    st.stop()


def main():
    st.set_page_config(page_title="Morale Event Planner", page_icon="🎄", layout="wide")
    apply_style()
    data, settings = load_or_stop()

    if "flash" in st.session_state:
        st.toast(st.session_state.pop("flash"))

    names = roster_names(data["roster"])
    days = settings.days()
    answers = availability_by_person(data["availability"], names, days)
    votes = votes_by_person(data["votes"], names)
    event_titles = idea_titles(data["events"])
    dining_titles = idea_titles(data["dining"])

    st.markdown('<div class="holiday-ribbon"></div>', unsafe_allow_html=True)
    st.markdown(f'<h1>{esc(settings.event_title)}<span class="holiday-mark">❄</span></h1>',
                unsafe_allow_html=True)
    st.caption(deadline_text(settings))
    me = who_are_you(settings, names)

    availability, events, dining, voting, results = st.tabs(
        ["Availability", "Events", "Dining", "Vote", "Results"]
    )
    with availability:
        availability_tab(settings, me, names, answers, days)
    with events:
        ideas_tab(settings, me, "events", "Event", data["events"])
    with dining:
        ideas_tab(settings, me, "dining", "Dining", data["dining"])
    with voting:
        votes_tab(settings, me, votes, event_titles, dining_titles)
    with results:
        results_tab(names, answers, votes, days, event_titles, dining_titles)


main()
