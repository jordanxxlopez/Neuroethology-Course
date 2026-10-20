#!/usr/bin/env python3
"""Make the Spring 2027 syllabus from the Fall 2026 syllabus.

    python3 tools/make_spring_2027_syllabus.py source/JL_NEUR411.docx source/spring-2027/JL_NEUR411_Spring2027.docx

Everything is copied from the fall Word file. Only these change: the term, the field-experience
dates (Spring Recess), exam coverage (three special topics), the course calendar, the readings
calendar dates, the lab topics (Version 2), and the due dates.
"""
import copy
import datetime as dt
import re
import sys

from docx import Document

FALL, OUT = sys.argv[1], sys.argv[2]
START, END = dt.date(2027, 1, 11), dt.date(2027, 4, 30)
HOLIDAYS = {dt.date(2027, 1, 18): "Martin Luther King Jr. Day", dt.date(2027, 2, 15): "Presidents' Day"}
RECESS = (dt.date(2027, 3, 14), dt.date(2027, 3, 21))

SPECIAL = {  # exam number: (special topic number, title, readings); given in the lecture slot on that exam day
    1: (1, "Jellyfish I: neuroanatomy of Cassiopea, distributed nerve nets, rhopalia, and pacemaker control of rhythmic "
           "bell contractions", [
        "Amplatz K, Zieger E, Abed-Navandi D, Weissenbacher A, Wanninger A (2024). Neuromuscular development in the "
        "emerging scyphozoan model system, Cassiopea xamachana: implications for the evolution of cnidarian nervous "
        "systems. Frontiers in Neuroscience 17: 1324980. https://doi.org/10.3389/fnins.2023.1324980",
        "Nakanishi N, Hartenstein V, Jacobs DK (2009). Development of the rhopalial nervous system in Aurelia sp.1 "
        "(Cnidaria, Scyphozoa). Development Genes and Evolution 219(6): 301-317. https://doi.org/10.1007/s00427-009-0291-y",
        "Muffett KM, Mammone M, Puckett J, Martino G, Aditi, Miglietta MP (2025). Regeneration and Musculature in Halved "
        "Cassiopea xamachana Ephyrae. Integrative Organismal Biology 7(1): obaf030. https://doi.org/10.1093/iob/obaf030",
        "Anuszczyk SR, Yoder N, Costello JH, Dabiri JO, Gemmell BJ, Rutledge KM, Colin SP (2025). Increasing the "
        "Reliability and Versatility of Jellyfish Biohybrid Vehicles via Species Selection and Rhopalia Removal. "
        "Biomimetics 10(12): 810. https://doi.org/10.3390/biomimetics10120810"]),
    2: (2, "Axolotl: lateral-line sensing, electroreception, and sensory processing of courtship signals", [
        "Wang Y, Pan Y, Liu X, Zhang X, Fan C, Wang J (2026). Characterization of the electrosensory receptor in axolotl "
        "and the role of voltage-gated potassium channels in electrosensitivity. Comparative Biochemistry and Physiology "
        "Part A 319:112051. https://doi.org/10.1016/j.cbpa.2026.112051",
        "Münz H, Claas B, Fritzsch B (1982). Electrophysiological evidence of electroreception in the axolotl Siredon "
        "mexicanum. Neuroscience Letters 28:107–111. https://doi.org/10.1016/0304-3940(82)90216-6",
        "Bartels M, Münz H, Claas B (1990). Representation of lateral line and electrosensory systems in the midbrain of "
        "the axolotl, Ambystoma mexicanum. Journal of Comparative Physiology A 167:347–356. "
        "https://doi.org/10.1007/BF00192570",
        "Maex M, Van Bocxlaer I, Mortier A, Proost P, Bossuyt F (2016). Courtship Pheromone Use in a Model Urodele, the "
        "Mexican Axolotl (Ambystoma mexicanum). Scientific Reports 6:20184. https://doi.org/10.1038/srep20184"]),
    3: (3, "Songbirds I: Artificial song memories, optogenetic tutoring, and neural mechanisms of vocal imitation", [
        "Zhao W, Garcia-Oscos F, Dinh D, Roberts TF (2019). Inception of memories that guide vocal learning in the "
        "songbird. Science 366:83–89. https://doi.org/10.1126/science.aaw4226",
        "Roberts TF, Gobes SMH, Murugan M, Ölveczky BP, Mooney R (2012). Motor circuits are required to encode a sensory "
        "model for imitative learning. Nature Neuroscience 15:1454–1459. https://doi.org/10.1038/nn.3206",
        "Tanaka M, Sun F, Li Y, Mooney R (2018). A mesocortical dopamine circuit enables the cultural transmission of "
        "vocal behaviour. Nature 563:117–120. https://doi.org/10.1038/s41586-018-0636-7",
        "Roberts TF, Tschida KA, Klein ME, Mooney R (2010). Rapid spine stabilization and synaptic enhancement at the "
        "onset of behavioural learning. Nature 463:948–952. https://doi.org/10.1038/nature08759"]),
}
EXAMS = {1: (dt.date(2027, 2, 8), "Covers Lectures 1 to 10", "12:00 PM, NEU 100A or NEU 100B"),
         2: (dt.date(2027, 3, 8), "Covers Lectures 11 to 20", "12:00 PM, NEU 100A or NEU 100B"),
         3: (dt.date(2027, 4, 12), "Covers Lectures 21 to 30", "12:00 PM, NEU 100A or NEU 100B"),
         4: (dt.date(2027, 5, 10), "Covers Lectures 31 to 40",
             "12:00 PM, NEU 100A or NEU 100B (final examination period)")}
LABS = [  # lab, week, Version 2 topic, report due
    ("Lab 1", 1, "Ethograms and Tinbergen's four questions: quantitative observation of zebrafish behavior from recorded video", dt.date(2027, 1, 24)),
    ("Lab 2", 2, "Cockroach cercal system: directional wind tuning of giant interneurons from published recordings", dt.date(2027, 1, 31)),
    ("Lab 3", 4, "Stomatogastric pyloric rhythm: burst timing, phase constancy, and neuromodulation from published recordings", dt.date(2027, 2, 14)),
    ("Lab 4", 5, "Toad prey-catching: worm versus anti-worm stimulus tuning from recorded behavior", dt.date(2027, 2, 21)),
    ("Lab 5", 6, "Jumping spider visual tracking: saccades and pursuit from video", dt.date(2027, 2, 28)),
    ("Lab 6", 8, "Moth hearing and bat evasion: auditory receptor responses to ultrasound and evasive flight", dt.date(2027, 3, 14)),
    ("Lab 7", 9, "Weakly electric fish: electric organ discharge recordings and active electrolocation", dt.date(2027, 3, 29)),
    ("Lab 8", 12, "Rapid prey capture: star-nosed mole handling times and archerfish shot geometry from high-speed video", dt.date(2027, 4, 11)),
    ("Lab 9", 13, "Orientation and navigation: circular statistics on monarch butterfly and desert ant headings", dt.date(2027, 4, 18)),
    ("Lab 10", 14, "Songbird song structure: syllable segmentation and song stereotypy in zebra finch recordings", dt.date(2027, 4, 25)),
]
MAKEUP_WEEK = 16
COMP = [dt.date(2027, 2, 4), dt.date(2027, 3, 11), dt.date(2027, 4, 15)]
REFLECTION = dt.date(2027, 4, 24)
PROPOSAL, PROGRESS, FINAL = dt.date(2027, 3, 29), dt.date(2027, 4, 14), dt.date(2027, 5, 7)
TRIP = [dt.date(2027, 3, 18), dt.date(2027, 3, 19), dt.date(2027, 3, 20), dt.date(2027, 3, 21)]


def short(d):
    return d.strftime("%a %b ") + str(d.day)


def week_monday(n):
    return START + dt.timedelta(weeks=n - 1)


# ---------------------------------------------------------------- docx text helpers

def replace(par, old, new):
    """Replace text inside a paragraph, keeping the formatting of the run where the match starts."""
    runs = par.runs
    full = "".join(r.text for r in runs)
    i = full.find(old)
    if i < 0:
        raise ValueError(f"not found: {old!r}")
    j, pos, done = i + len(old), 0, False
    for r in runs:
        a, b = pos, pos + len(r.text)
        pos = b
        if b <= i or a >= j:
            continue
        lo, hi = max(i, a) - a, min(j, b) - a
        r.text = r.text[:lo] + ("" if done else new) + r.text[hi:]
        done = True


def set_par(par, text):
    for h in par._p.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink"):
        h.getparent().remove(h)
    runs = par.runs
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


NOTE_SIZE = None  # font size of the calendar's notes column, read from the fall file


def set_cell(cell, lines, size=None):
    """Put one line per paragraph, reusing the cell's paragraphs (and their formatting) where they exist."""
    pars = cell.paragraphs
    while len(pars) < len(lines):
        new = copy.deepcopy(pars[-1]._p)
        pars[-1]._p.addnext(new)
        pars = cell.paragraphs
    for p, line in zip(pars, lines):
        set_par(p, line)
        if size:
            for r in p.runs:
                r.font.size = size
    for p in pars[len(lines):]:
        p._p.getparent().remove(p._p)


def find_par(doc, start):
    return next(p for p in doc.paragraphs if p.text.strip().startswith(start))


# ---------------------------------------------------------------- build the spring schedule

def class_days():
    days, d = [], START
    while d <= END:
        if d.weekday() in (0, 2, 4) and d not in HOLIDAYS and not (RECESS[0] <= d <= RECESS[1]):
            days.append(d)
        d += dt.timedelta(days=1)
    return days


def main():
    doc = Document(FALL)
    cal, readings, labs_t, grading, milestones = doc.tables[7], doc.tables[8], doc.tables[6], doc.tables[2], doc.tables[5]

    # fall lecture titles, notes and readings rows, keyed by lecture number
    fall_rows = {}
    for r in cal.rows[1:]:
        m = re.match(r"Lecture (\d+)", " ".join(r.cells[1].text.split()))
        if m:
            fall_rows[int(m.group(1))] = r
    read_rows = {}
    for r in readings.rows[1:]:
        m = re.search(r"(\d+)", r.cells[1].text)
        if m:
            read_rows[int(m.group(1))] = r
    template = {}
    for r in cal.rows[1:]:
        s = " ".join(r.cells[1].text.split())
        template.setdefault(s.split(" ")[0] if s else "", r)
    t_week = cal.rows[1]
    t_due = next(r for r in cal.rows if r.cells[2].text.strip() == "Lab 1 report due")
    t_noclass = next(r for r in cal.rows if r.cells[1].text.strip() == "No class")
    t_exam = next(r for r in cal.rows if r.cells[1].text.strip() == "Exam 1")
    t_comp = next(r for r in cal.rows if r.cells[1].text.strip().startswith("Comp"))
    t_refl = next(r for r in cal.rows if "Reflection" in r.cells[1].text)
    t_proj = next(r for r in cal.rows if r.cells[1].text.strip() == "Project")
    t_make = next(r for r in cal.rows if "ake-up" in r.cells[1].text)
    field_rows = [r for r in cal.rows if r.cells[1].text.strip().startswith("Special Topic")]

    global NOTE_SIZE
    NOTE_SIZE = next(r.font.size for r in t_due.cells[3].paragraphs[0].runs if r.text.strip())
    # order of sessions: 40 lectures with the three special topics after their lectures
    days = class_days()
    when, n, pending = {}, 0, None
    for d in days:
        if pending and d.weekday() == 0:
            when[("ST", pending)] = d
            pending = None
            continue
        n += 1
        when[("L", n)] = d
        if n in (10, 20, 30):
            pending = n // 10
    assert n == 40 and len(when) == 43, (n, len(when))
    for ex, (d, _, _) in EXAMS.items():
        assert ex == 4 or when[("ST", ex)] == d, (ex, d)

    notes_extra = {PROPOSAL: "Project proposal due 11:59 PM", PROGRESS: "Project progress check due 11:59 PM"}
    events = []  # (date, order, row)

    def add(date, order, src, cells):
        row = copy.deepcopy(src._tr)
        events.append((date, order, row, cells))

    for key, d in when.items():
        kind, n = key
        extra = [notes_extra[d]] if d in notes_extra else []
        if kind == "L":
            src = fall_rows[n]
            notes = [p.text for p in src.cells[3].paragraphs if p.text.strip()]
            notes = [x for x in notes if "Project" not in x] + extra
            add(d, 1, src, [[short(d)], [f"Lecture {n}"], None, notes or [""]])
        else:
            st, title, _ = SPECIAL[n]
            add(d, 1, fall_rows[2], [[short(d)], [f"Special Topic {st}"], [title],
                                     ["Not on the exam"] + extra])
    for n, (d, cover, note) in EXAMS.items():
        add(d, 2, t_exam, [[short(d)], [f"Exam {n}"], [cover], [note]])
    for i, (lab, wk, topic, due) in enumerate(LABS):
        first = topic[0].lower() + topic[1:]
        notes = ["Must have approval by IACUC before entering lab"] if i == 0 else [""]
        add(week_monday(wk), 0, t_week, [[f"Week {wk}"], [lab], [f"{lab} meets: {first}"], notes])
        add(due, 3, t_due, [[short(due)], [lab], [f"{lab} report due"], ["Report due"]])
    add(week_monday(MAKEUP_WEEK), 0, t_make, [[f"Week {MAKEUP_WEEK}"], None, None, None])
    for i, d in enumerate(COMP, 1):
        topic = [x for x in cal.rows if x.cells[1].text.strip() == f"Comp. {i}"][0]
        add(d, 1, topic, [[short(d)], None, None, None])
    add(REFLECTION, 1, t_refl, [[short(REFLECTION)], None, None, None])
    add(FINAL, 1, t_proj, [[short(FINAL)], None, None, None])
    for d, name in HOLIDAYS.items():
        add(d, 1, t_noclass, [[short(d)], ["No class"], [name], [""]])
    add(dt.date(2027, 3, 15), 1, t_noclass, [["Mon Mar 15"], ["No class"],
                                            ["Spring Recess (no classes March 15–19)"], [""]])
    for d, src in zip(TRIP, field_rows):
        add(d, 2, src, [[short(d)], None, None, None])

    # rebuild the calendar table
    for r in list(cal.rows)[1:]:
        r._tr.getparent().remove(r._tr)
    from docx.table import _Row
    for d, order, tr, cells in sorted(events, key=lambda e: (e[0], e[1])):
        cal._tbl.append(tr)
        row = _Row(tr, cal)
        for i, (c, lines) in enumerate(zip(row.cells, cells)):
            if lines is not None:
                set_cell(c, lines, NOTE_SIZE if i == 3 else None)
    set_par(cal.rows[0].cells[0].paragraphs[0], "Date (2027)")

    # readings calendar: same rows with spring dates, plus the special topics
    for r in list(readings.rows)[1:]:
        r._tr.getparent().remove(r._tr)
    for key, d in when.items():
        kind, n = key
        if kind == "L":
            tr = copy.deepcopy(read_rows[n]._tr)
            readings._tbl.append(tr)
            row = _Row(tr, readings)
            set_cell(row.cells[0], [short(d)])
        else:
            st, title, refs = SPECIAL[n]
            tr = copy.deepcopy(read_rows[1]._tr)
            readings._tbl.append(tr)
            row = _Row(tr, readings)
            set_cell(row.cells[0], [short(d)])
            set_cell(row.cells[1], [f"Special Topic {st}"])
            set_cell(row.cells[2], [title])
            set_cell(row.cells[3], refs)
    set_par(readings.rows[0].cells[0].paragraphs[0], "Date (2027)")

    # lab schedule (Version 2 topics), computational assignments, reflection
    for row, (lab, wk, topic, due) in zip(labs_t.rows[1:11], LABS):
        set_cell(row.cells[1], [f"Week {wk}"])
        set_cell(row.cells[2], [topic])
        set_cell(row.cells[3], [f"Report due {short(due)}"])
    set_cell(labs_t.rows[11].cells[1], [f"Week {MAKEUP_WEEK}"])
    for row, d in zip(labs_t.rows[12:15], COMP):
        set_cell(row.cells[1], [short(d)])
    set_cell(labs_t.rows[15].cells[1], [short(REFLECTION)])

    # grading table and project milestones
    for row, d in zip(milestones.rows[1:4], (PROPOSAL, PROGRESS, FINAL)):
        set_cell(row.cells[1], [f"{short(d)}, 11:59 PM"])

    # term, field experience, exams, reflection text
    replace(find_par(doc, "Fall 2026"), "Fall 2026", "Spring 2027")
    replace(find_par(doc, "Required Field Experience"), "Fall Recess", "Spring Recess")
    p = find_par(doc, "The Fall Recess Field Experience")
    replace(p, "The Fall Recess Field Experience", "The Spring Recess Field Experience")
    replace(p, "Thursday, October 8 through Sunday, October 11, 2026", "Thursday, March 18 through Sunday, March 21, 2027")
    for old, new in (("Thursday, October 8:", "Thursday, March 18:"), ("Friday, October 9:", "Friday, March 19:"),
                     ("Saturday, October 10:", "Saturday, March 20:"), ("Sunday, October 11:", "Sunday, March 21:")):
        replace(find_par(doc, old), old, new)
    p = find_par(doc, "There are four exams")
    replace(p, "and they start on time.", "and they start on time. Exam 4 is given during the final examination "
               "period. On the days of Exams 1, 2, and 3, the lecture is a special topic; special topics are not "
               "covered on any exam.")
    p = find_par(doc, "Captivity and Conservation Reflection (15%)")
    replace(p, "Fall Recess Field Experience", "Spring Recess Field Experience")
    replace(p, "The essay is due Sat Nov 21.", f"The essay is due {short(REFLECTION)}.")

    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            for par in part.paragraphs:
                if "Fall 2026" in par.text:
                    replace(par, "Fall 2026", "Spring 2027")
    doc.core_properties.title = "NEUR 411 Neuroethology, Spring 2027"
    doc.save(OUT)
    print(f"saved {OUT}: {len(events)} calendar rows, {len(readings.rows) - 1} readings rows")


if __name__ == "__main__":
    main()
