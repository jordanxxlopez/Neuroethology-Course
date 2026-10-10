# NEUR 411 Neuroethology — Fall 2026 course website

A plain, accessible course website for students: syllabus, course calendar, lecture slides, recommended
readings, and labs and assignments. Grading and submissions stay in Brightspace.

The site is generated from two sources, so the website never drifts from the syllabus:

| File | What it controls |
|---|---|
| `source/JL_NEUR411.docx` | The syllabus. Every policy, the course calendar, lab schedule and readings list on the site are read from it. |
| `source/lectures.json` | One entry per lecture deck: its Google Drive file ID, file name and size. |
| `source/site.json` | `public` (search-engine visibility) and the name of the grading platform. |

The built site is in `docs/` (plain HTML and one stylesheet, no framework).

## Update the site

```bash
pip install python-docx
python3 tools/build_site.py
```

- **Syllabus changed:** replace `source/JL_NEUR411.docx` (keep the file name) and rebuild. The PDF and Word
  downloads are refreshed automatically when LibreOffice (`soffice`) is installed.
- **New version of a lecture deck** (e.g. the scaffolded version): upload it to Drive, then put its file ID
  (the long string in the Drive link `drive.google.com/file/d/<ID>/view`), file name and size in bytes into
  that lecture's entry in `source/lectures.json`, and rebuild. To keep the same link, use Drive's
  *Manage versions → Upload new version* on the existing file instead; then nothing here changes.
- **Host a deck on the site itself instead of Drive:** copy it to `docs/files/lectures/` and add
  `"local": "<file name>"` to its entry. (GitHub rejects single files over 100 MB.)

## Publish (GitHub Pages)

Repository **Settings → Pages → Build and deployment → Deploy from a branch**, pick the branch and the
`/docs` folder. The site appears at `https://jordanxxlopez.github.io/neuroethology-course/`.

## Keeping it private for now

- The lecture links point to the files in Google Drive, so **Drive sharing is what controls access to the
  slides**. Share the *NEUROETHOLGY-UPDATED LECTURES* folder (or the individual files) with
  **“USC — anyone in this group with the link” as Viewer**, so only signed-in USC accounts can open them.
  Students do not see the folder itself, only the file linked for each lecture.
- `source/site.json` has `"public": false`, which adds `noindex` tags and a `robots.txt` asking search
  engines not to list the site. Set it to `true` when the course goes public.
- GitHub Pages sites are reachable by anyone who has the URL, even from a private repository (private-repo
  Pages needs a paid GitHub plan; access-controlled Pages needs GitHub Enterprise Cloud). The pages
  themselves hold only syllabus information.
- Before going public, check reuse rights for the published figures and videos in the decks.

## Accessibility

Semantic HTML with landmarks, a skip link and one `h1` per page; WCAG 2.2 AA contrast in light and dark
modes; visible focus and underlined links; captioned tables with header cells; a stacked calendar layout on
phones; no reliance on color alone (exam rows are labeled “Exam”); works without JavaScript (the only script
fills the home page's “Coming up” list). Every page passes axe-core's WCAG 2.2 AA and best-practice rules.
