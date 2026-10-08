#!/usr/bin/env python3
"""Fill the ATS CV Word template with a client's approved CV content.

The design is NOT rebuilt here: every paragraph is cloned from the real
template (assets/ATS_Template.docx), so fonts, sizes, colours, the line under
each heading, bullet style and tab-aligned dates are exactly the template's.

Usage:
    python build_cv.py cv.json --out "473.9091 1090.docx" [--pdf]
"""
import argparse
import copy
import json
import os
import re
import shutil
import subprocess
import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "assets", "ATS_Template.docx")

GREY = "595959"
LINK_BLUE = "1155CC"
BODY = 21  # half-points = 10.5pt

# ---------------------------------------------------------------- xml helpers
def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def rpr(bold=False, italic=False, underline=False, color="000000", size=BODY,
        style=None):
    r = el("w:rPr")
    if style:
        r.append(el("w:rStyle", **{"w:val": style}))
    r.append(el("w:rFonts", **{"w:ascii": "Calibri", "w:hAnsi": "Calibri",
                               "w:cs": "Calibri"}))
    if bold:
        r.append(el("w:b"))
        r.append(el("w:bCs"))
    if italic:
        r.append(el("w:i"))
        r.append(el("w:iCs"))
    if color:
        r.append(el("w:color", **{"w:val": color}))
    r.append(el("w:sz", **{"w:val": size}))
    r.append(el("w:szCs", **{"w:val": size}))
    if underline:
        r.append(el("w:u", **{"w:val": "single"}))
    return r


def run(text, **fmt):
    r = el("w:r")
    r.append(rpr(**fmt))
    t = el("w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def tab_run(**fmt):
    r = el("w:r")
    r.append(rpr(**fmt))
    r.append(el("w:tab"))
    return r


def set_spacing(p, **vals):
    ppr = p.find(qn("w:pPr"))
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = el("w:spacing")
        # spacing must come after tabs/numPr/pBdr and before jc/rPr
        anchor = None
        for tag in ("w:ind", "w:jc", "w:rPr"):
            anchor = ppr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(sp)
        else:
            ppr.append(sp)
    for k, v in vals.items():
        sp.set(qn("w:" + k), str(v))


def bold_split(text, **fmt):
    """'**Bold** rest' -> runs; lets section text carry bold parts."""
    out = []
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if part:
            out.append(run(part, **dict(fmt, bold=(i % 2 == 1) or fmt.get("bold", False))))
    return out


# ---------------------------------------------------------------- builder
LINE = 276      # 1.15 line spacing
BLACK = "000000"


def ppr_insert(p, elem, before=("w:shd", "w:tabs", "w:spacing", "w:ind",
                                "w:jc", "w:rPr")):
    """Insert a pPr child respecting the schema order."""
    ppr = p.find(qn("w:pPr"))
    old = ppr.find(elem.tag)
    if old is not None:
        ppr.remove(old)
    for tag in before:
        anchor = ppr.find(qn(tag))
        if anchor is not None:
            anchor.addprevious(elem)
            return
    ppr.append(elem)


def set_jc(p, val):
    ppr_insert(p, el("w:jc", **{"w:val": val}), before=("w:rPr",))


def keep_next(p):
    ppr = p.find(qn("w:pPr"))
    if ppr.find(qn("w:keepNext")) is None:
        ppr.insert(0, el("w:keepNext"))


class CVBuilder:
    def __init__(self):
        self.doc = Document(TEMPLATE)
        body = self.doc.element.body
        paras = body.findall(qn("w:p"))
        texts = ["".join(t.text or "" for t in p.iter(qn("w:t"))) for p in paras]

        def find(pred, start=0):
            for i in range(start, len(paras)):
                if pred(texts[i], paras[i]):
                    return i
            raise SystemExit("Template changed: prototype paragraph not found")

        i_name = find(lambda t, p: "TYPE THE NAME" in t)
        i_contact = find(lambda t, p: t.startswith("Phone:"))
        i_head = find(lambda t, p: t.strip() == "PROFILE SUMMARY")
        i_summary = find(lambda t, p: t.startswith("Add a summary"))
        i_label = find(lambda t, p: t.startswith("Title of the Skills"))
        i_skills = find(lambda t, p: t.startswith("Skills name"))
        i_job = find(lambda t, p: t.startswith("Position Title"))
        i_overview = find(lambda t, p: t.startswith("Write details"))
        i_bullet = find(lambda t, p: p.find(".//" + qn("w:numPr")) is not None)
        i_edu1 = find(lambda t, p: t.startswith("DEGREE TYPE"))
        i_edu2 = find(lambda t, p: t.startswith("Name of the University"))
        i_item = find(lambda t, p: t.startswith("Certificate or Training"))
        i_lang = find(lambda t, p: t.startswith("Arabic:"))

        self.proto = {k: copy.deepcopy(paras[i]) for k, i in dict(
            name=i_name, contact=i_contact, heading=i_head, summary=i_summary,
            label=i_label, skills=i_skills, job=i_job, overview=i_overview,
            bullet=i_bullet, edu1=i_edu1, edu2=i_edu2, item=i_item,
            lang=i_lang).items()}
        self._bullet_style(self.proto["bullet"])

        # Clear the template body, keep only the section properties.
        self.sect = body.find(qn("w:sectPr"))
        for child in list(body):
            if child is not self.sect:
                body.remove(child)
        mar = self.sect.find(qn("w:pgMar"))
        if mar is not None:
            mar.set(qn("w:bottom"), "720")  # 0.5" (template had 0.14")

    def _bullet_style(self, bullet_p):
        """Small round Symbol bullet, 0.19" in, text at 0.49"."""
        num_id = bullet_p.find(".//" + qn("w:numId")).get(qn("w:val"))
        numbering = self.doc.part.numbering_part.element
        num = [n for n in numbering.findall(qn("w:num"))
               if n.get(qn("w:numId")) == num_id][0]
        abs_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
        absn = [a for a in numbering.findall(qn("w:abstractNum"))
                if a.get(qn("w:abstractNumId")) == abs_id][0]
        lvl = [l for l in absn.findall(qn("w:lvl"))
               if l.get(qn("w:ilvl")) == "0"][0]
        lvl.find(qn("w:lvlText")).set(qn("w:val"), "")
        ind = lvl.find(qn("w:pPr")).find(qn("w:ind"))
        ind.set(qn("w:left"), "700")
        ind.set(qn("w:hanging"), "420")
        r = lvl.find(qn("w:rPr"))
        for child in list(r):
            r.remove(child)
        r.append(el("w:rFonts", **{"w:ascii": "Symbol", "w:hAnsi": "Symbol",
                                   "w:hint": "default"}))
        r.append(el("w:color", **{"w:val": BLACK}))
        r.append(el("w:sz", **{"w:val": 20}))
        r.append(el("w:szCs", **{"w:val": 20}))

    # -- paragraph factory
    def para(self, kind, runs=(), **spacing):
        p = copy.deepcopy(self.proto[kind])
        for child in list(p):
            if child.tag != qn("w:pPr"):
                p.remove(child)
        for r in runs:
            p.append(r)
        if spacing:
            set_spacing(p, **spacing)
        self.sect.addprevious(p)
        return p

    def link(self, text, url):
        rid = self.doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
        h = el("w:hyperlink")
        h.set(qn("r:id"), rid)
        h.set(qn("w:history"), "1")
        h.append(run(text, color=LINK_BLUE, underline=True, style="Hyperlink"))
        return h

    # -- header
    def header(self, name, title, contact):
        runs = [run(name.upper(), bold=True, color=BLACK, size=48)]
        if title:
            runs += [run("|", color=BLACK, size=52), run(" ", size=40),
                     run(title.upper(), bold=True, color=GREY, size=24)]
        self.para("name", runs, before=100, after=0)

        parts = []
        phone = contact.get("phone")
        if phone:
            digits = re.sub(r"[^\d+]", "", phone)
            parts.append(("Phone:", self.link(phone, "tel:" + digits)))
        email = contact.get("email")
        if email:
            parts.append(("Email:", self.link(email, "mailto:" + email)))
        address = contact.get("address")
        if address:
            parts.append(("Address:", run(address, color=GREY)))
        li = contact.get("linkedin")
        if li:
            url = li if li.startswith("http") else "https://" + li.lstrip("/")
            shown = re.sub(r"^https?://(www\.)?", "", li)
            parts.append(("LinkedIn:", self.link(shown, url)))
        runs = []
        for i, (label, value) in enumerate(parts):
            if i:
                runs.append(run(" | ", color=GREY))
            # non-breaking space keeps a label on the same line as its value
            runs += [run(label, color=GREY), run("\u00a0", color=GREY), value]
        self.para("contact", runs, after=0, line=LINE)

    # -- sections
    def heading(self, text):
        p = self.para("heading", [run(text.upper(), bold=True, color=BLACK, size=26)],
                      before=280, after=180)
        bdr = el("w:pBdr")
        bdr.append(el("w:bottom", **{"w:val": "single", "w:sz": 8,
                                     "w:space": 4, "w:color": BLACK}))
        ppr_insert(p, bdr)
        keep_next(p)

    def summary(self, s):
        self.para("summary", bold_split(s["text"]), before=0, after=0, line=LINE)

    def skills(self, s):
        for gi, g in enumerate(s["groups"]):
            label = g.get("label")
            if label:
                p = self.para("label", [run(label, bold=True)],
                              before=240 if gi else 0, after=0, line=LINE)
                keep_next(p)
            items = [i.strip() for i in g["items"]]
            p = self.para("skills", [run(" | ".join(items))],
                          before=0 if label else (240 if gi else 0), after=0, line=LINE)
            set_jc(p, "both")

    def entries(self, s):
        for i, e in enumerate(s["items"]):
            runs = bold_split(e["title"], bold=True)
            if e.get("date"):
                runs += [tab_run(bold=True), run(e["date"], bold=True)]
            lines = ([e["overview"]] if e.get("overview") else []) + e.get("lines", [])
            has_more = bool(lines or e.get("bullets"))
            p = self.para("job", runs, before=240 if i else 0,
                          after=160 if has_more else 0, line=LINE)
            keep_next(p)
            for j, line in enumerate(lines):
                p = self.para("overview", bold_split(line), before=0, line=LINE,
                              after=80 if e.get("bullets") or j < len(lines) - 1 else 0)
                set_jc(p, "both")
            for b in e.get("bullets", []):
                self.para("bullet", [run(b)], before=0, after=90, line=LINE)

    def education(self, s):
        for i, e in enumerate(s["items"]):
            runs = [run(e["degree"], bold=True)]
            if e.get("detail"):
                runs.append(run(" | " + e["detail"], bold=True))
            if e.get("date"):
                runs += [tab_run(bold=True), run(e["date"], bold=True)]
            p = self.para("edu1", runs, before=200 if i else 0, after=0)
            if e.get("institution"):
                keep_next(p)
                self.para("edu2", [run(e["institution"])], before=0, after=0)

    def items(self, s):
        for e in s["items"]:
            if isinstance(e, str):
                runs = bold_split(e)
            else:
                name = e["name"].strip().rstrip(",")
                details = (e.get("details") or "").strip()
                if not details:
                    runs = [run(name, bold=True)]
                elif details.startswith("|"):   # "Name | 2024"
                    runs = [run(name, bold=True), run(" " + details)]
                else:                           # "Name, Org, City | 2024"
                    runs = [run(name + ",", bold=True), run(" " + details)]
            self.para("item", runs, before=0, after=60, line=LINE)

    def languages(self, s):
        langs = [(l["language"].strip().rstrip(":") + ":", " " + l["level"].strip())
                 for l in s["items"]]
        if s.get("layout") == "lines":          # one language per line
            for lang, level in langs:
                self.para("lang", [run(lang, bold=True), run(level)],
                          before=0, after=0, line=LINE)
            return
        runs = []
        for i, (lang, level) in enumerate(langs):
            if i:
                runs.append(run(" | "))
            runs += [run(lang, bold=True), run(level)]
        self.para("lang", runs, before=0, after=0, line=LINE)

    def build(self, data):
        self.header(data["name"], data.get("title", ""), data.get("contact", {}))
        render = dict(summary=self.summary, skills=self.skills,
                      entries=self.entries, education=self.education,
                      items=self.items, languages=self.languages)
        for s in data["sections"]:
            if not any(s.get(k) for k in ("text", "groups", "items")):
                print("SKIPPED empty section: %s" % s.get("heading"), file=sys.stderr)
                continue
            if s["type"] not in render:
                raise SystemExit("Unknown section type: %s" % s["type"])
            self.heading(s["heading"])
            render[s["type"]](s)
        cp = self.doc.core_properties
        cp.title = "%s CV" % data["name"]
        cp.author = cp.last_modified_by = ""
        cp.subject = cp.keywords = cp.comments = ""


# ---------------------------------------------------------------- pdf
def to_pdf(docx_path):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        print("WARNING: LibreOffice not found - PDF not created", file=sys.stderr)
        return None
    out_dir = os.path.dirname(os.path.abspath(docx_path))
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", out_dir,
                    docx_path], check=True, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=180)
    pdf = os.path.splitext(docx_path)[0] + ".pdf"
    with open(pdf, "rb") as f:
        pages = len(re.findall(rb"/Type\s*/Page[^s]", f.read()))
    print("PDF: %s (%d page%s)" % (pdf, pages, "" if pages == 1 else "s"))
    return pdf


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", help="CV content as JSON")
    ap.add_argument("--out", required=True, help="output .docx path")
    ap.add_argument("--pdf", action="store_true", help="also export PDF")
    a = ap.parse_args()
    with open(a.data, encoding="utf-8") as f:
        data = json.load(f)
    b = CVBuilder()
    b.build(data)
    b.doc.save(a.out)
    print("DOCX: %s" % a.out)
    if a.pdf:
        to_pdf(a.out)


if __name__ == "__main__":
    main()
