---
name: atscv-word
description: Builds the client's CV as a Word (.docx) and PDF file in the exact ATS template design (Calibri, centred name header, black line under every heading, tab-aligned dates, • bullets). Use whenever an approved CV must be turned into Word/PDF files — e.g. the user says "تمام", "اعتمد", "طلعه", "ok go", "حطه في الوورد", or asks for the .docx/.pdf of a CV.
---

# ATS CV → Word & PDF

The design lives in `assets/ATS_Template.docx`. `scripts/build_cv.py` clones
that file's own paragraphs and only swaps in the text, so the output always
looks exactly like the template. **Never build the CV with docx-js, a new
python-docx document, HTML, or any other method — and never change fonts,
sizes, colours or spacing.** Your only job is to put the approved text into
a JSON file and run the script.

## Steps

1. Take the **approved** CV text from the chat, word for word. Do not reword,
   shorten, reorder or "improve" anything at this stage.
2. Write it to `cv.json` using the format below. Section order in the JSON =
   order in the CV. Leave out any section the client has no content for.
3. Run:
   ```bash
   pip install python-docx --break-system-packages -q 2>/dev/null
   python /path/to/atscv-word/scripts/build_cv.py cv.json --out "/mnt/user-data/outputs/473.9091 1090.docx" --pdf
   ```
   (use the real skill path, the real order number and the client's phone
   number without country code, as two groups of 4 digits). The script also
   creates the PDF next to the .docx and prints the page count.
4. Check the page count: 1 page if the client has under 2 years of
   experience, max 2 pages otherwise. If it overflows, tell the user which
   content to trim (do not trim on your own) — never shrink fonts or margins.
5. Render page 1 of the PDF to an image (`pdftoppm -r 60 -png -f 1 -l 1`)
   and look at it: no empty headings, nothing cut off, dates on the right.
6. Send both files (.docx and .pdf). No recap afterwards.

## cv.json format

```json
{
  "name": "Ahmed Salim Al Harthy",
  "title": "HSE Officer",
  "contact": {
    "phone": "+968 9417 7338",
    "email": "ahmed@gmail.com",
    "address": "Muscat, Sultanate of Oman",
    "linkedin": "linkedin.com/in/ahmed"
  },
  "sections": [
    {"type": "summary", "heading": "PROFILE SUMMARY", "text": "..."},

    {"type": "skills", "heading": "KEY SKILLS", "groups": [
      {"label": "Technical Skills:", "items": ["Risk Assessment", "Permit to Work"]},
      {"label": "Soft Skills:", "items": ["Communication", "Teamwork"]}
    ]},

    {"type": "entries", "heading": "PROFESSIONAL EXPERIENCE", "items": [
      {"title": "HSE Officer | PDO – Fahud", "date": "March 2022 – Present",
       "overview": "One-line role overview.",
       "bullets": ["Bullet one.", "Bullet two."]}
    ]},

    {"type": "entries", "heading": "ACADEMIC PROJECT", "items": [
      {"title": "Project Title", "lines": ["What was done...", "Outcome..."]}
    ]},

    {"type": "education", "heading": "EDUCATION", "items": [
      {"degree": "BACHELOR OF ENGINEERING IN CHEMICAL ENGINEERING",
       "detail": "CGPA: 3.12/4.00", "date": "2016 – 2020",
       "institution": "SULTAN QABOOS UNIVERSITY, MUSCAT, OMAN"}
    ]},

    {"type": "items", "heading": "COURSES & CERTIFICATIONS", "items": [
      {"name": "NEBOSH IGC", "details": "NEBOSH, Muscat | 2022"},
      {"name": "Certificate in Sewerage Networks", "details": "| 2024"}
    ]},

    {"type": "items", "heading": "AWARDS", "items": [
      {"name": "Best Employee Award", "details": "PDO | 2023"}
    ]},

    {"type": "entries", "heading": "ACTIVITIES / VOLUNTEER WORK", "items": [
      {"title": "Volunteer | Oman Red Crescent – Muscat", "date": "2019 – 2020",
       "bullets": ["Supported first aid awareness campaigns."]}
    ]},

    {"type": "languages", "heading": "LANGUAGES", "items": [
      {"language": "Arabic", "level": "Native"},
      {"language": "English", "level": "Fluent"}
    ]}
  ]
}
```

### Section types
| type | use for | fields per item |
|---|---|---|
| `summary` | PROFILE SUMMARY | `text` |
| `skills` | KEY SKILLS | `groups[]`: `label`, `items[]` |
| `entries` | PROFESSIONAL EXPERIENCE, ACADEMIC PROJECT, ACTIVITIES / VOLUNTEER WORK, any extra section with dates/duties | `title` (`Job Title \| Company – City`), `date` (optional), `overview` (optional), `lines[]` (optional plain lines), `bullets[]` (optional) |
| `education` | EDUCATION | `degree`, `detail` (optional, e.g. GPA), `date`, `institution` |
| `items` | COURSES & CERTIFICATIONS, AWARDS, simple extra sections (MEMBERSHIPS, PUBLICATIONS…) | `name` (bold), `details`. Details starting with `\|` print as `**Name** \| 2024`; otherwise as `**Name,** Org, City \| 2024`. Or a plain string where `**text**` is bold |
| `languages` | LANGUAGES | `language`, `level`; add `"layout": "lines"` on the section for one language per line (default: all on one line with " \| ") |

### Rules
- `heading` is printed in CAPS exactly as given — use the project's heading
  names; an extra section just needs its own heading and a suitable type.
- Dates go in `date` only — the script puts them at the right margin with a
  tab. Never put dates inside `title`.
- Use `–` (en dash) in dates and between company and city, as in the
  approved text.
- Omit `linkedin` if the client has none; omit `detail` if there is no GPA;
  omit `institution` only if it is truly unknown.
- Write text exactly as approved (CAPS or Title Case as in the draft) — the
  script only forces CAPS on the name, target role and section headings.
- A section with no items is skipped automatically, but leave it out anyway.
