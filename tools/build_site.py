#!/usr/bin/env python3
"""Build the NEUR 411 course website from the syllabus and the lecture manifest.

Sources (edit these, then rebuild):
  source/JL_NEUR411.docx   the syllabus; every policy, the calendar and the readings come from it
  source/lectures.json     one Google Drive file (or local file) per lecture deck
  source/site.json         site-wide settings (search-engine visibility, LMS name)

Output: docs/ (served by GitHub Pages). Run:  python3 tools/build_site.py
Requires: python-docx. Optional: LibreOffice (soffice) to refresh the syllabus PDF.
"""
import datetime as dt
import html
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "source"
OUT = ROOT / "docs"
SYLLABUS = SRC / "JL_NEUR411.docx"
YEAR = 2026

# Syllabus sections that get their own page instead of appearing on the syllabus page.
OWN_PAGE = {"Course Calendar", "Readings Calendar"}
# Syllabus sections shown on the Labs & Assignments page (and also kept in the full syllabus).
ASSIGNMENT_SECTIONS = ["Laboratory Component", "Lecture Project (20% of Lecture Grade)", "Late Work"]

esc = html.escape


# ---------------------------------------------------------------- docx → HTML

URL_RE = re.compile(r"(https?://[^\s<]+[^\s<.,;:)\]])")
EMAIL_RE = re.compile(r"(?<![\w.@/])([\w.+-]+@[\w-]+(?:\.[\w-]+)+)")


def linkify(text_html):
    """Turn bare URLs and e-mail addresses (already HTML-escaped) into links."""
    text_html = URL_RE.sub(lambda m: f'<a href="{m.group(1)}">{m.group(1)}</a>', text_html)
    return EMAIL_RE.sub(lambda m: f'<a href="mailto:{m.group(1)}">{m.group(1)}</a>', text_html)


def runs_html(par):
    """Inline HTML for a paragraph: bold/italic runs and hyperlinks, URLs linkified."""
    pieces = []  # (html, bold, italic, is_link)
    for item in par.iter_inner_content():
        if isinstance(item, Hyperlink):
            text = item.text
            href = item.address or ""
            if not href and item.fragment:
                href = "#" + item.fragment
            if href.startswith("mailto:") or href.startswith("http"):
                pieces.append((f'<a href="{esc(href)}">{esc(text)}</a>', False, False, True))
            else:
                pieces.append((esc(text), False, False, False))
        else:
            if not item.text:
                continue
            pieces.append((esc(item.text), bool(item.bold), bool(item.italic), False))
    # merge neighbouring runs with the same formatting so tags are not fragmented
    merged = []
    for h, b, i, link in pieces:
        if merged and not link and not merged[-1][3] and merged[-1][1:3] == (b, i):
            merged[-1] = (merged[-1][0] + h, b, i, False)
        else:
            merged.append((h, b, i, link))
    out = []
    for h, b, i, link in merged:
        if not link:
            h = linkify(h)
        if not h.strip():
            out.append(h)
            continue
        if i:
            h = f"<em>{h}</em>"
        if b:
            h = f"<strong>{h}</strong>"
        out.append(h)
    s = "".join(out)
    s = re.sub(r"</strong>(\s*)<strong>", r"\1", s)
    return s.strip()


def is_list(par):
    ppr = par._p.pPr
    return ppr is not None and ppr.numPr is not None


def fully_bold(par):
    runs = [r for r in par.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def cell_html(cell):
    parts = [runs_html(p) for p in cell.paragraphs]
    return "<br>".join(p for p in parts if p)


def table_html(table, caption=None, hide_caption=False):
    rows = table.rows
    head = "".join(f'<th scope="col">{cell_html(c)}</th>' for c in rows[0].cells)
    body = []
    for r in rows[1:]:
        cells = [cell_html(c) for c in r.cells]
        if not any(cells):
            continue
        tds = [f'<th scope="row">{cells[0]}</th>'] + [f"<td>{c}</td>" for c in cells[1:]]
        body.append("<tr>" + "".join(tds) + "</tr>")
    cls = ' class="visually-hidden"' if hide_caption else ""
    cap = f"<caption{cls}>{esc(caption)}</caption>" if caption else ""
    return wrap_table(f"<table>{cap}<thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table>", caption)


_region = [0]


def wrap_table(table, label=None):
    """A keyboard-focusable scroll region so wide tables stay usable on small screens."""
    _region[0] += 1
    lab = f' aria-label="{esc(label)} (table)"' if label else ' aria-label="Table"'
    return f'<div class="table-wrap" role="region" tabindex="0"{lab}>{table}</div>'


def iter_body(doc):
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def read_sections(doc):
    """Split the syllabus at Heading 2 into [(title, [blocks])]; title None = front matter."""
    sections = [(None, [])]
    for block in iter_body(doc):
        if isinstance(block, Paragraph) and block.style.name == "Heading 2":
            title = block.text.strip()
            if title:
                sections.append((title, []))
            continue
        sections[-1][1].append(block)
    return sections


def blocks_html(blocks, base_level=3, table_captions=None, label=None):
    """Render a section's paragraphs, bullet lists, sub-headings and tables."""
    out, in_list = [], False
    captions = iter(table_captions or [])
    last_heading = None
    used = set()
    for b in blocks:
        if isinstance(b, Table):
            if in_list:
                out.append("</ul>")
                in_list = False
            cap = next(captions, None) or last_heading
            hide = cap is None
            cap = cap or label
            if cap in used:
                cap = None
            used.add(cap)
            out.append(table_html(b, cap, hide_caption=hide))
            continue
        text = b.text.strip()
        if not text:
            continue
        if is_list(b):
            if not in_list:
                out.append("<ul>")
                in_list = True
            out.append(f"<li>{runs_html(b)}</li>")
            continue
        if in_list:
            out.append("</ul>")
            in_list = False
        if b.style.name == "Heading 3" or (fully_bold(b) and len(text) < 90):
            last_heading = text
            level = base_level
            out.append(f'<h{level} id="{slug(text)}">{esc(text)}</h{level}>')
        else:
            out.append(f"<p>{runs_html(b)}</p>")
    if in_list:
        out.append("</ul>")
    return "\n".join(out)


def slug(text):
    s = re.sub(r"\(.*?\)", "", text.lower())
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "section"


# ---------------------------------------------------------------- calendar and readings

MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


def parse_date(label):
    m = re.match(r"(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+([A-Z][a-z]{2})\s+(\d{1,2})", label)
    if not m:
        return None
    return dt.date(YEAR, MONTHS[m.group(2)], int(m.group(3)))


def row_kind(session):
    s = session.lower()
    if s.startswith("lecture"):
        return "lecture"
    if s.startswith("exam"):
        return "exam"
    if s.startswith("no class"):
        return "noclass"
    if s.startswith("special"):
        return "special"
    if s.startswith("lab") or s.startswith("make"):
        return "lab"
    return "due"


def read_calendar(table):
    rows = []
    for r in table.rows[1:]:
        cells = r.cells
        date_label = cells[0].text.strip()
        session = " ".join(cells[1].text.split())
        rows.append({
            "date_label": date_label,
            "date": parse_date(date_label),
            "session": session,
            "kind": row_kind(session),
            "topic_html": cell_html(cells[2]),
            "topic": " ".join(cells[2].paragraphs[0].text.split()),
            "notes_html": cell_html(cells[3]),
        })
    return rows


def read_readings(table):
    out = {}
    for r in table.rows[1:]:
        m = re.search(r"(\d+)", r.cells[1].text)
        if not m:
            continue
        refs = []
        for p in r.cells[3].paragraphs:
            t = p.text.strip()
            if t:
                refs.append(t)
        out[int(m.group(1))] = refs
    return out


def reading_html(ref):
    """'Authors (year). Title. Journal ... https://doi.org/x' → structured citation."""
    m = re.search(r"\s*(https?://\S+)\s*$", ref)
    url = None
    if m:
        url = m.group(1)
        ref = ref[: m.start()]
    body = esc(ref.strip())
    am = re.match(r"(.+?\(\d{4}[a-z]?\)\.)\s*(.*)", body)
    if am:
        body = f'<span class="authors">{am.group(1)}</span> {am.group(2)}'
    link = ""
    if url:
        label = url.replace("https://doi.org/", "doi:") if "doi.org/" in url else url
        link = f' <a class="doi" href="{esc(url)}">{esc(label)}</a>'
    return f"<li>{body}{link}</li>"


# ---------------------------------------------------------------- page shell

NAV = [
    ("index.html", "Home"),
    ("syllabus.html", "Syllabus"),
    ("schedule.html", "Schedule"),
    ("lectures.html", "Lectures"),
    ("readings.html", "Readings"),
    ("assignments.html", "Labs &amp; Assignments"),
]


def page(filename, title, body, settings, description=""):
    nav = []
    for href, label in NAV:
        cur = ' aria-current="page"' if href == filename else ""
        nav.append(f'<li><a href="{href}"{cur}>{label}</a></li>')
    robots = '<meta name="robots" content="noindex, nofollow">\n' if not settings["public"] else ""
    full_title = f"{title} | NEUR 411 Neuroethology" if filename != "index.html" else "NEUR 411 Neuroethology, Fall 2026"
    built = dt.date.today().strftime("%B %-d, %Y")
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
{robots}<title>{full_title}</title>
<meta name="description" content="{esc(description or 'Course website for NEUR 411 Neuroethology, Fall 2026.')}">
<link rel="stylesheet" href="assets/site.css">
</head>
<body>
<a class="skip-link" href="#main">Skip to main content</a>
<header class="site-header">
  <div class="wrap">
    <p class="site-title"><a href="index.html">NEUR 411 <span class="sep" aria-hidden="true">·</span> Neuroethology</a></p>
    <p class="site-term">Fall 2026 <span aria-hidden="true">·</span> Jordan Lopez, M.D., Ph.D.</p>
  </div>
  <nav class="site-nav" aria-label="Main">
    <ul class="wrap">{''.join(nav)}</ul>
  </nav>
</header>
<main id="main" tabindex="-1">
<div class="wrap">
{body}
</div>
</main>
<footer class="site-footer">
  <div class="wrap">
    <p>NEUR 411 Neuroethology, Fall 2026. Instructor: Jordan Lopez, M.D., Ph.D.,
    <a href="mailto:jlopez11@usc.edu">jlopez11@usc.edu</a>. Grades and submissions are handled in {settings['lms']}.</p>
    <p><a href="accessibility.html">Accessibility</a> <span aria-hidden="true">·</span> Last updated {built}</p>
  </div>
</footer>
</body>
</html>
"""


def fmt_size(n):
    return f"{n / 1_000_000:.1f} MB"


def section_html(sections, title, level=2, sid=None):
    for t, blocks in sections:
        if t == title:
            return (f'<section aria-labelledby="{sid or slug(t)}"><h{level} id="{sid or slug(t)}">{esc(t)}</h{level}>\n'
                    f"{blocks_html(blocks, base_level=level + 1, label=t)}</section>")
    raise KeyError(title)


# ---------------------------------------------------------------- pages

def build_home(sections, cal, lectures, settings):
    by = dict(sections)
    desc = next(b for b in by["Course Description"] if isinstance(b, Paragraph) and b.text.strip())
    instructor = blocks_html(by["Instructor"])
    times = by["Course Times & Locations"]
    events = []
    for r in cal:
        if r["date"] is None or r["kind"] == "noclass":
            continue
        label = r["topic"]
        if r["kind"] == "lecture":
            label = f'{r["session"]}: {label}'
        elif r["kind"] == "exam":
            label = f'{r["session"]}: {label}'
        elif r["kind"] == "special":
            label = f"Field experience: {label}"
        events.append({"d": r["date"].isoformat(), "label": label, "kind": r["kind"],
                       "when": r["date_label"]})
    events_json = json.dumps(events)
    body = f"""
<h1>NEUR 411 Neuroethology</h1>
<p class="lede">{runs_html(desc)}</p>

<div class="grid-2">
<section aria-labelledby="instructor">
<h2 id="instructor">Instructor</h2>
{instructor}
<p>Teaching assistants and the exam proctors are listed in the <a href="syllabus.html#teaching-assistants">syllabus</a>. Contact your TA first for lab, report and grading questions.</p>
</section>
<section aria-labelledby="meetings">
<h2 id="meetings">Course times and locations</h2>
{blocks_html(times, label="Course times and locations")}
</section>
</div>

<section aria-labelledby="materials">
<h2 id="materials">Course materials</h2>
<ul class="link-list">
<li><a href="lectures.html">Lecture slides</a>: all 40 lectures, grouped by exam.</li>
<li><a href="schedule.html">Course calendar</a>: lectures, exams, labs, field experience and due dates.</li>
<li><a href="readings.html">Recommended readings</a>: the primary papers behind each lecture, with DOI links.</li>
<li><a href="assignments.html">Labs and assignments</a>: lab schedule, computational assignments, lecture project and reflection essay.</li>
<li><a href="syllabus.html">Syllabus</a>: grading, exams, policies; also as a <a href="files/NEUR411_Syllabus_Fall2026.pdf">PDF</a> or <a href="files/NEUR411_Syllabus_Fall2026.docx">Word document</a>.</li>
</ul>
<p>Grades and assignment submissions are handled in {settings['lms']}. This site holds the course materials.</p>
</section>

<section aria-labelledby="upcoming" class="upcoming">
<h2 id="upcoming">Coming up</h2>
<ol id="upcoming-list" class="event-list"></ol>
<p id="upcoming-fallback">See the <a href="schedule.html">course calendar</a> for every lecture, exam and due date.</p>
</section>
<script>
(function () {{
  var events = {events_json};
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var list = document.getElementById('upcoming-list');
  var shown = 0;
  for (var i = 0; i < events.length && shown < 6; i++) {{
    var e = events[i], d = new Date(e.d + 'T00:00:00');
    if (d < today) continue;
    var li = document.createElement('li');
    if (e.kind === 'exam') li.className = 'is-exam';
    var t = document.createElement('time'); t.dateTime = e.d; t.textContent = e.when;
    li.appendChild(t); li.appendChild(document.createTextNode(' ' + e.label));
    list.appendChild(li); shown++;
  }}
  if (!shown) list.hidden = true;
}})();
</script>
"""
    return page("index.html", "Home", body, settings)


def build_syllabus(sections, settings):
    parts, toc = [], []
    for title, blocks in sections:
        if title is None:
            continue
        sid = slug(title)
        if title in OWN_PAGE:
            target = "schedule.html" if title == "Course Calendar" else "readings.html"
            toc.append(f'<li><a href="{target}">{esc(title)}</a> (separate page)</li>')
            continue
        toc.append(f'<li><a href="#{sid}">{esc(title)}</a></li>')
        parts.append(f'<section aria-labelledby="{sid}"><h2 id="{sid}">{esc(title)}</h2>\n{blocks_html(blocks, label=title)}</section>')
    body = f"""
<h1>Syllabus</h1>
<p class="lede">NEUR 411 Neuroethology, Fall 2026. Download the full syllabus as a
<a href="files/NEUR411_Syllabus_Fall2026.pdf">PDF</a> or a
<a href="files/NEUR411_Syllabus_Fall2026.docx">Word document</a>.</p>
<nav class="toc" aria-labelledby="toc-heading">
<h2 id="toc-heading">On this page</h2>
<ol>{''.join(toc)}</ol>
</nav>
{''.join(parts)}
"""
    return page("syllabus.html", "Syllabus", body, settings,
                "Syllabus for NEUR 411 Neuroethology: grading, exams, labs and course policies.")


def build_schedule(sections, cal, settings):
    by = dict(sections)
    intro = blocks_html([b for b in by["Course Calendar"] if isinstance(b, Paragraph)])
    intro = intro.replace("listed in the Lab Schedule above",
                          'listed on the <a href="assignments.html#lab-schedule-and-due-dates">Labs &amp; Assignments</a> page')
    rows = []
    for r in cal:
        session = esc(r["session"])
        topic = r["topic_html"]
        if r["kind"] == "lecture":
            n = int(re.search(r"\d+", r["session"]).group())
            topic = f'<a href="lectures.html#lecture-{n}">{topic}</a>'
        when = esc(r["date_label"])
        if r["date"]:
            when = f'<time datetime="{r["date"].isoformat()}">{when}</time>'
        notes = f'<td role="cell" class="notes">{r["notes_html"]}</td>' if r["notes_html"] else '<td role="cell" class="notes empty"></td>'
        rows.append(f'<tr role="row" class="row-{r["kind"]}"><th role="rowheader" scope="row">{when}</th>'
                    f'<td role="cell" class="session">{session}</td><td role="cell">{topic}</td>{notes}</tr>')
    table = ('<table class="schedule" role="table"><caption>Course calendar, Fall 2026. Rows marked Exam are exams; '
             'rows that begin with a week number are lab meetings for that week.</caption>'
             '<thead><tr role="row"><th role="columnheader" scope="col">Date (2026)</th><th role="columnheader" scope="col">Session</th>'
             '<th role="columnheader" scope="col">Topic</th><th role="columnheader" scope="col">Assessment / notes</th></tr></thead>'
             f'<tbody>{"".join(rows)}</tbody></table>')
    body = f"""
<h1>Course calendar</h1>
{intro}
<p>Lecture titles link to the slides. Exams: Exam 1 covers Lectures 1 to 10, Exam 2 Lectures 11 to 20,
Exam 3 Lectures 21 to 30 and Exam 4 Lectures 31 to 40.</p>
<div class="table-wrap schedule-wrap">{table}</div>
"""
    return page("schedule.html", "Schedule", body, settings,
                "Course calendar for NEUR 411 Neuroethology, Fall 2026.")


def deck_links(entry, n):
    hidden = f'<span class="visually-hidden"> for Lecture {n}</span>'
    if entry.get("local"):
        href = f'files/lectures/{entry["local"]}'
        return (f'<a class="btn" href="{esc(href)}" download>Download slides{hidden}</a> '
                f'<span class="meta">PowerPoint, {fmt_size(entry["size"])}</span>')
    fid = entry["drive_id"]
    view = f"https://drive.google.com/file/d/{fid}/view"
    dl = f"https://drive.google.com/uc?export=download&amp;id={fid}"
    return (f'<a class="btn" href="{view}">View slides{hidden}</a> '
            f'<a class="btn btn-secondary" href="{dl}">Download PowerPoint{hidden}</a> '
            f'<span class="meta">.pptx, {fmt_size(entry["size"])}</span>')


def lecture_rows(cal):
    out = []
    for r in cal:
        if r["kind"] == "lecture":
            out.append((int(re.search(r"\d+", r["session"]).group()), r))
    return out


def build_lectures(cal, lectures, readings, settings):
    manifest = {e["n"]: e for e in lectures["lectures"]}
    blocks = {1: [], 2: [], 3: [], 4: []}
    for n, r in lecture_rows(cal):
        blocks[(n - 1) // 10 + 1].append((n, r))
    exam_dates = {int(re.search(r"\d+", r["session"]).group()): r["date_label"]
                  for r in cal if r["kind"] == "exam"}
    out = []
    for exam, items in blocks.items():
        lo, hi = (exam - 1) * 10 + 1, exam * 10
        lis = []
        for n, r in items:
            entry = manifest.get(n)
            links = deck_links(entry, n) if entry else '<span class="meta">Slides not yet posted.</span>'
            nread = len(readings.get(n, []))
            rlink = (f' <a href="readings.html#lecture-{n}">Recommended readings'
                     f'<span class="visually-hidden"> for Lecture {n}</span></a> <span class="meta">({nread})</span>'
                     if nread else "")
            when = f'<time datetime="{r["date"].isoformat()}">{esc(r["date_label"])}</time>' if r["date"] else ""
            lis.append(f"""<li class="lecture" id="lecture-{n}">
<h3><span class="lnum">Lecture {n}</span> {r['topic_html']}</h3>
<p class="meta">{when}</p>
<p class="actions">{links}{rlink}</p>
</li>""")
        out.append(f"""<section aria-labelledby="exam-{exam}">
<h2 id="exam-{exam}">Lectures {lo} to {hi} <span class="meta">(Exam {exam}, {esc(exam_dates.get(exam, ''))})</span></h2>
<ol class="lecture-list">{''.join(lis)}</ol>
</section>""")
    body = f"""
<h1>Lecture slides</h1>
<p class="lede">Slides for every lecture. Each deck opens in Google Drive, where it can be read in the browser
without PowerPoint, or downloaded as a PowerPoint file with speaker notes.</p>
<p>Sign in with your USC Google account if Drive asks for access. Lecture slides are the primary source of
instruction, and exams are based on lecture content, including material presented in class that is not on
the slides. Updated versions of each deck will replace these as they are posted.</p>
<nav class="toc" aria-label="Lecture groups"><ul class="inline">
<li><a href="#exam-1">Lectures 1 to 10</a></li><li><a href="#exam-2">11 to 20</a></li>
<li><a href="#exam-3">21 to 30</a></li><li><a href="#exam-4">31 to 40</a></li></ul></nav>
{''.join(out)}
"""
    return page("lectures.html", "Lectures", body, settings,
                "Lecture slides for all 40 lectures of NEUR 411 Neuroethology.")


def build_readings(sections, cal, readings, settings):
    by = dict(sections)
    intro = blocks_html([b for b in by["Readings Calendar"] if isinstance(b, Paragraph)])
    out = []
    for n, r in lecture_rows(cal):
        refs = readings.get(n, [])
        if not refs:
            continue
        when = esc(r["date_label"])
        out.append(f"""<section class="reading-group" aria-labelledby="lecture-{n}">
<h2 id="lecture-{n}"><span class="lnum">Lecture {n}</span> {r['topic_html']}</h2>
<p class="meta">{when} <span aria-hidden="true">·</span> <a href="lectures.html#lecture-{n}">Slides<span class="visually-hidden"> for Lecture {n}</span></a></p>
<ul class="refs">{''.join(reading_html(x) for x in refs)}</ul>
</section>""")
    jump = "".join(f'<li><a href="#lecture-{n}">{n}</a></li>' for n, _ in lecture_rows(cal) if readings.get(n))
    body = f"""
<h1>Recommended readings</h1>
{intro}
<p>Links go to each paper's DOI. Many journals require signing in through the USC Libraries; off campus, use the library proxy.</p>
<nav class="toc" aria-labelledby="jump"><h2 id="jump" class="visually-hidden">Jump to a lecture</h2>
<p class="meta" aria-hidden="true">Jump to lecture:</p><ul class="inline numbers">{jump}</ul></nav>
{''.join(out)}
"""
    return page("readings.html", "Readings", body, settings,
                "Recommended primary readings for each NEUR 411 lecture.")


def build_assignments(sections, settings):
    parts = [section_html(sections, t) for t in ASSIGNMENT_SECTIONS]
    body = f"""
<h1>Labs and assignments</h1>
<p class="lede">The lab schedule, report due dates, computational assignments, lecture project and the
Captivity and Conservation Reflection. Submit work in {settings['lms']} or Gradescope as directed; grading
weights are in the <a href="syllabus.html#grading">syllabus</a>.</p>
{''.join(parts)}
"""
    return page("assignments.html", "Labs & Assignments", body, settings,
                "Lab schedule, assignments and project milestones for NEUR 411.")


def build_accessibility(settings):
    body = f"""
<h1>Accessibility</h1>
<p class="lede">This site is built to be usable with a keyboard, a screen reader, browser zoom and high-contrast
or dark display settings.</p>
<h2>How the site is built</h2>
<ul>
<li>Every page has the same navigation, a skip link to the main content and a logical heading structure.</li>
<li>Text and links meet WCAG 2.2 AA contrast in both light and dark modes; links are underlined, and the keyboard focus is always visible.</li>
<li>Tables have captions and header cells. Wide tables scroll inside their own region on small screens.</li>
<li>Pages work without JavaScript. The only script on the site lists upcoming dates on the home page.</li>
<li>The syllabus is available as an HTML page, a PDF and a Word document.</li>
</ul>
<h2>Lecture slides</h2>
<p>Slides open in Google Drive's viewer, which works with screen readers and zoom, and can be downloaded as
PowerPoint files that include the speaker notes. Figures in the slides come from published research articles;
each carries its source and citation.</p>
<h2>Requesting another format or reporting a problem</h2>
<p>If any material is hard to use, e-mail <a href="mailto:jlopez11@usc.edu">jlopez11@usc.edu</a> with the page
or lecture number and the format you need. For approved accommodations, contact the Office of Student
Accessibility Services and see the <a href="syllabus.html#accommodations">accommodations policy</a>.</p>
"""
    return page("accessibility.html", "Accessibility", body, settings)


# ---------------------------------------------------------------- main

def refresh_syllabus_files():
    files = OUT / "files"
    files.mkdir(parents=True, exist_ok=True)
    docx_out = files / "NEUR411_Syllabus_Fall2026.docx"
    shutil.copyfile(SYLLABUS, docx_out)
    pdf_out = files / "NEUR411_Syllabus_Fall2026.pdf"
    stale = not pdf_out.exists() or pdf_out.stat().st_mtime < SYLLABUS.stat().st_mtime
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if stale and soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(files), str(docx_out)],
                       check=True, capture_output=True)
    elif stale:
        print("warning: syllabus PDF is out of date and LibreOffice is not installed", file=sys.stderr)


def main():
    settings = json.loads((SRC / "site.json").read_text())
    lectures = json.loads((SRC / "lectures.json").read_text())
    doc = Document(SYLLABUS)
    sections = read_sections(doc)
    by = dict(sections)
    cal = read_calendar(next(b for b in by["Course Calendar"] if isinstance(b, Table)))
    readings = read_readings(next(b for b in by["Readings Calendar"] if isinstance(b, Table)))

    OUT.mkdir(exist_ok=True)
    pages = {
        "index.html": build_home(sections, cal, lectures, settings),
        "syllabus.html": build_syllabus(sections, settings),
        "schedule.html": build_schedule(sections, cal, settings),
        "lectures.html": build_lectures(cal, lectures, readings, settings),
        "readings.html": build_readings(sections, cal, readings, settings),
        "assignments.html": build_assignments(sections, settings),
        "accessibility.html": build_accessibility(settings),
    }
    for name, content in pages.items():
        (OUT / name).write_text(content)
    (OUT / ".nojekyll").write_text("")
    robots = "User-agent: *\nAllow: /\n" if settings["public"] else "User-agent: *\nDisallow: /\n"
    (OUT / "robots.txt").write_text(robots)
    refresh_syllabus_files()

    n_lec = len(lecture_rows(cal))
    n_ref = sum(len(v) for v in readings.values())
    missing = sorted(set(range(1, n_lec + 1)) - {e["n"] for e in lectures["lectures"]})
    print(f"built {len(pages)} pages: {n_lec} lectures, {n_ref} readings"
          + (f"; no slides for lectures {missing}" if missing else ""))


if __name__ == "__main__":
    main()
