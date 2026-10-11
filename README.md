# NEUR 411 Neuroethology — Fall 2026 course website

A plain, accessible course website for students: syllabus, course calendar, lecture slides, recommended
readings, and labs and assignments. Grading and submissions stay in Brightspace.

The site is generated from two sources, so the website never drifts from the syllabus:

| File | What it controls |
|---|---|
| `source/JL_NEUR411.docx` | The syllabus. Every policy, the course calendar, lab schedule and readings list on the site are read from it. |
| `source/lectures.json` | One entry per lecture deck: which file is the current version, and its copy on the site. |
| `source/site.json` | `public` (search-engine visibility) and the name of the grading platform. |

The built site is in `docs/` (plain HTML and one stylesheet, no framework).

## Update the site

```bash
pip install python-docx
python3 tools/build_site.py
```

- **Syllabus changed:** replace `source/JL_NEUR411.docx` (keep the file name) and rebuild. The PDF and Word
  downloads are refreshed automatically when LibreOffice (`soffice`) is installed.
- **Lecture slides** link to the decks in Google Drive (*NEUROETHOLGY-UPDATED LECTURES*, shared with USC as
  Viewer), with View and Download buttons. To post a new version (e.g. the scaffolded deck), use Drive's
  *Manage versions → Upload new version* on the existing file so the link stays the same; or put the new
  file's Drive ID, name and size in `source/lectures.json` and rebuild.
  Optional: `python3 tools/add_decks.py <zip or folder>` hosts decks on the site itself instead of Drive.
- **Class and university resources** are edited in `source/class_resources.html` and
  `source/university_resources.html`.

## Publish (GitHub Pages)

Repository **Settings → Pages → Build and deployment → Deploy from a branch**, pick the branch and the
`/docs` folder. The site appears at `https://jordanxxlopez.github.io/neuroethology-course/`.

## Keeping it private for now

- Leave GitHub Pages **off** until the course goes live. To look at the site, download the repository
  (Code → Download ZIP) and open `docs/index.html`.
- Going live: Settings → Pages → Deploy from a branch → this branch, `/docs`. Search engines are asked not to
  list the site while `"public": false` in `source/site.json`.
- The slides stay behind USC sign-in through the Drive folder's sharing setting.
- Before going public, check reuse rights for the published figures and videos in the decks.

## Accessibility

Semantic HTML with landmarks, a skip link and one `h1` per page; WCAG 2.2 AA contrast, always on a white background (dark mode is ignored)
visible focus and underlined links; captioned tables with header cells; a stacked calendar layout on
phones; no reliance on color alone (exam rows are labeled “Exam”); works without JavaScript (the only script
fills the home page's “Coming up” list). Every page passes axe-core's WCAG 2.2 AA and best-practice rules.
