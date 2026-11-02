"""Team morale event planner.

Everything that changes per event lives in the Google Sheet, not in this file:
  config tab  - title, date range, deadline, lock behaviour
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
    ("first_date", "2026-12-01", "First date people can choose (YYYY-MM-DD)"),
    ("last_date", "2026-12-31", "Last date people can choose (YYYY-MM-DD)"),
    ("weekdays_only", "yes", "yes = Monday to Friday only, no = include weekends"),
    ("skip_dates", "", "Dates to leave out, comma separated (YYYY-MM-DD)"),
    ("deadline", "2026-11-26", "Last day for responses (YYYY-MM-DD)"),
    ("timezone", "America/Los_Angeles", "Deadline is 11:59 PM in this time zone"),
    ("lock_after_deadline", "yes", "yes = read-only after the deadline, no = just show the date"),
]

YES, NO = "Yes", "No"
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
    return True


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
    deadline: datetime
    lock_after_deadline: bool

    @property
    def past_deadline(self):
        return datetime.now(self.deadline.tzinfo) > self.deadline

    @property
    def is_open(self):
        return not (self.past_deadline and self.lock_after_deadline)

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
    deadline_day = parse_date("deadline", raw["deadline"])
    return Settings(
        event_title=raw["event_title"] or "Morale Event",
        first_date=first_date,
        last_date=last_date,
        weekdays_only=is_yes(raw["weekdays_only"]),
        skip_dates=skip,
        deadline=datetime.combine(deadline_day, time(23, 59, 59), tzinfo=zone),
        lock_after_deadline=is_yes(raw["lock_after_deadline"]),
    )


def day_key(day):
    """The one and only format used for date columns in the sheet."""
    return day.isoformat()


def day_label(day):
    return f"{day.strftime('%a %b')} {day.day}"


def deadline_text(settings):
    moment = settings.deadline
    when = f"{moment.strftime('%a, %b')} {moment.day} at 11:59 PM {moment.strftime('%Z')}"
    if settings.past_deadline:
        return f"Responses closed {when}." if settings.lock_after_deadline else f"The deadline was {when}."
    left = (moment.date() - datetime.now(moment.tzinfo).date()).days
    countdown = "today" if left == 0 else "1 day left" if left == 1 else f"{left} days left"
    return f"Respond by {when} ({countdown})."


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


def availability_by_person(frame, names, days):
    """{name: {day_key: Yes/No}} for people on the roster. First row per name wins."""
    answers = {}
    if frame.empty:
        return answers
    canonical = {n.casefold(): n for n in names}
    keys = [day_key(d) for d in days]
    for _, row in frame.iterrows():
        name = canonical.get(row["Name"].strip().casefold())
        if name and name not in answers:
            answers[name] = {k: row[k].strip() for k in keys if k in row and row[k].strip() in (YES, NO)}
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


def date_summary(answers, names, days):
    rows = []
    for day in days:
        key = day_key(day)
        yes = [n for n in names if answers.get(n, {}).get(key) == YES]
        no = [n for n in names if answers.get(n, {}).get(key) == NO]
        rows.append({
            "Date": day_label(day),
            "Can make it": len(yes),
            "Can't make it": len(no),
            "No answer": len(names) - len(yes) - len(no),
            "Who can't": ", ".join(no),
        })
    return pd.DataFrame(sorted(rows, key=lambda r: (-r["Can make it"], r["Can't make it"])),
                        columns=["Date", "Can make it", "Can't make it", "No answer", "Who can't"])


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
        .block-container { max-width: 1100px; padding-top: 2.5rem; }
        .stTabs [data-baseweb="tab"] { font-weight: 600; }
        .day-top { display: flex; justify-content: space-between; align-items: baseline; gap: .4rem; }
        .day-name { font-weight: 600; }
        .day-count { font-size: .75rem; color: #5f6368; white-space: nowrap; }
        .chips { display: flex; flex-wrap: wrap; gap: 3px; padding-bottom: .9rem; }
        [data-testid="stColumn"] [data-testid="stVerticalBlock"]:has(.day-top) { height: 100%; }
        .chip { font-size: .7rem; padding: 0 5px; border-radius: 4px; line-height: 1.5; }
        .chip-yes { background: #e6f4ea; color: #1a6b34; }
        .chip-no { background: #fce8e6; color: #a8231a; }
        .idea { border: 1px solid #e3e3df; border-radius: 8px; padding: .55rem .8rem; margin-bottom: .5rem; }
        .idea-head { display: flex; justify-content: space-between; gap: .75rem; }
        .idea-title { font-weight: 600; }
        .idea-price { color: #5f6368; font-size: .85rem; white-space: nowrap; }
        .idea-by { color: #80868b; font-size: .75rem; }
        .idea-desc { color: #444; font-size: .85rem; margin-top: .15rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def calendar_grid(settings, days, render_day):
    """Month grids with one bordered cell per selectable day."""
    selectable = set(days)
    columns = 5 if settings.weekdays_only else 7
    for year, month in sorted({(d.year, d.month) for d in days}):
        st.markdown(f"#### {calendar.month_name[month]} {year}")
        for week in calendar.Calendar(firstweekday=0).monthdatescalendar(year, month):
            week = week[:columns]
            if not selectable.intersection(week):
                continue
            for column, day in zip(st.columns(columns, gap="small"), week):
                if day in selectable and day.month == month:
                    with column, st.container(border=True):
                        render_day(day)


def day_html(day, answers, names):
    key = day_key(day)
    yes = [n for n in names if answers.get(n, {}).get(key) == YES]
    no = [n for n in names if answers.get(n, {}).get(key) == NO]
    chips = "".join(f'<span class="chip chip-yes">{esc(n)}</span>' for n in yes)
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

    if not settings.is_open:
        st.info("Responses are closed. This is what everyone said.")
        calendar_grid(settings, days, read_only_day)
        return
    if not me:
        st.info("Choose your name above to fill this in.")
        calendar_grid(settings, days, read_only_day)
        return

    st.caption("Choose Yes or No for each date, then press Save at the bottom. "
               "Press your choice again to clear it. Names under each date show saved answers.")
    mine = answers.get(me, {})

    with st.form(f"availability_{me}", border=False):
        choices = {}

        def editable_day(day):
            key = day_key(day)
            top, chips = day_html(day, answers, names)
            st.markdown(top, unsafe_allow_html=True)
            choices[key] = st.segmented_control(
                day_label(day), [YES, NO], default=mine.get(key), key=f"day_{me}_{key}",
                label_visibility="collapsed", width="stretch",
            )
            st.markdown(chips, unsafe_allow_html=True)

        calendar_grid(settings, days, editable_day)
        saved = st.form_submit_button("Save my availability", type="primary")

    if saved:
        try:
            upsert("availability", "Name", me, {"Name": me, **{k: v or "" for k, v in choices.items()}})
            flash("Availability saved")
            st.rerun()
        except gspread.exceptions.GSpreadException as err:
            show_save_error(err)


def idea_card(row):
    title = esc(row["Title"].strip())
    link = row["Link"].strip()
    if link.lower().startswith(("http://", "https://")):
        title = f'<a href="{html.escape(link, quote=True)}" target="_blank" rel="noopener">{title}</a>'
    price = row["Est. Price/Person"].strip().lstrip("$").strip()
    price_html = f'<span class="idea-price">{esc("$" + price)} per person</span>' if price else ""
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
        if not settings.is_open:
            st.info("Responses are closed.")
        elif not me:
            st.info("Choose your name above to suggest an idea.")
        else:
            st.markdown(f"**Suggest {'an' if label == 'Event' else 'a'} {label.lower()} idea**")
            with st.form(f"idea_{sheet_tab}", clear_on_submit=True, border=False):
                title = st.text_input("Title", max_chars=TITLE_LIMIT)
                description = st.text_area("Description (optional)", height=80)
                link = st.text_input("Link (optional)", placeholder="https://")
                price = st.text_input("Estimated price per person (optional)", placeholder="25")
                submitted = st.form_submit_button("Add idea", type="primary")
            if submitted:
                add_idea(sheet_tab, me, title, description, link, price)


def add_idea(sheet_tab, me, title, description, link, price):
    title, link = title.strip(), link.strip()
    if not title:
        st.warning("Add a title first.")
    elif link and not link.lower().startswith(("http://", "https://")):
        st.warning("Links need to start with http:// or https://.")
    else:
        row = [me, title, description.strip(), link, price.strip()]
        try:
            if append_unique(sheet_tab, "Title", title, row):
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
    st.caption("Rank up to 3 of each. You can change your vote until the deadline.")
    if not settings.is_open:
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

    if saved:
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

    if settings.is_open:
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
    st.set_page_config(page_title="Morale Event Planner", layout="wide")
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

    st.title(settings.event_title)
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
