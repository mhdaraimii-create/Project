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
LABEL_GREY = "808080"
LINK_BLUE = "1155CC"
BODY = 21  # half-points = 10.5pt

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


# ---------------------------------------------------------------- xml helpers
def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def rpr(bold=False, italic=False, underline=False, color=None, size=BODY,
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

        has_line = lambda t, p: p.find(".//" + qn("w:drawing")) is not None
        i_name = find(lambda t, p: "TYPE THE NAME" in t)
        i_contact = find(lambda t, p: t.startswith("Phone:"))
        i_head = find(lambda t, p: t.strip() == "PROFILE SUMMARY")
        i_line = find(has_line)
        i_summary = find(lambda t, p: t.startswith("Add a summary"))
        i_label = find(lambda t, p: t.startswith("Title of the Skills"))
        i_skills = find(lambda t, p: t.startswith("Skills name"))
        i_job1 = find(lambda t, p: t.startswith("Position Title"))
        i_overview = find(lambda t, p: t.startswith("Write details"))
        i_bullet = find(lambda t, p: p.find(".//" + qn("w:numPr")) is not None)
        i_job2 = find(lambda t, p: t.startswith("Position Title"), i_job1 + 1)
        i_edu1 = find(lambda t, p: t.startswith("DEGREE TYPE"))
        i_edu2 = find(lambda t, p: t.startswith("Name of the University"))
        i_item = find(lambda t, p: t.startswith("Certificate or Training"))
        i_lang = find(lambda t, p: t.startswith("Arabic:"))

        self.proto = {k: copy.deepcopy(paras[i]) for k, i in dict(
            name=i_name, contact=i_contact, heading=i_head, line=i_line,
            summary=i_summary, label=i_label, skills=i_skills,
            job_first=i_job1, job_next=i_job2, overview=i_overview,
            bullet=i_bullet, edu1=i_edu1, edu2=i_edu2, item=i_item,
            lang=i_lang).items()}

        # Drop the legacy VML fallback of the heading line (duplicate ids).
        for fb in self.proto["line"].iter("{%s}Fallback" % MC_NS):
            fb.getparent().remove(fb)

        # Clear the template body, keep only the section properties.
        self.sect = body.find(qn("w:sectPr"))
        for child in list(body):
            if child is not self.sect:
                body.remove(child)
        mar = self.sect.find(qn("w:pgMar"))
        if mar is not None:
            mar.set(qn("w:bottom"), "720")  # 0.5" (template had 0.14")
        self.body = body
        self.shape_id = 5000

    # -- paragraph factory
    def para(self, kind, runs=()):
        p = copy.deepcopy(self.proto[kind])
        for child in list(p):
            if child.tag != qn("w:pPr"):
                p.remove(child)
        for r in runs:
            p.append(r)
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
        runs = [run(name.upper(), bold=True, size=52)]
        if title:
            runs += [run("|", size=56), run(" ", bold=True, size=40),
                     run(title.upper(), bold=True, color=GREY, size=24)]
        self.para("name", runs)

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
            runs += [run(label, bold=True, color=GREY), run("\u00a0", color=GREY), value]
        self.para("contact", runs)

    # -- sections
    def heading(self, text):
        self.para("heading", [run(text.upper(), bold=True, size=26)])
        p = copy.deepcopy(self.proto["line"])
        for dp in p.iter("{%s}docPr" % WP_NS):
            self.shape_id += 1
            dp.set("id", str(self.shape_id))
        self.sect.addprevious(p)

    def summary(self, s):
        self.para("summary", bold_split(s["text"]))

    def skills(self, s):
        for gi, g in enumerate(s["groups"]):
            label = g.get("label")
            if label:
                p = self.para("label", [run(label, bold=True, italic=True,
                                            underline=True, color=LABEL_GREY)])
                if gi:
                    set_spacing(p, before=160)
            # non-breaking spaces stop a multi-word skill splitting over lines
            items = [i.strip().replace(" ", "\u00a0") for i in g["items"]]
            p = self.para("skills", [run(" | ".join(items))])
            set_spacing(p, before=60 if label else 120, after=0)

    def entries(self, s):
        for i, e in enumerate(s["items"]):
            runs = bold_split(e["title"], bold=True)
            if e.get("date"):
                runs += [tab_run(bold=True), run(e["date"], bold=True)]
            p = self.para("job_first" if i == 0 else "job_next", runs)
            set_spacing(p, after=0)
            lines = ([e["overview"]] if e.get("overview") else []) + e.get("lines", [])
            for j, line in enumerate(lines):
                p = self.para("overview", bold_split(line))
                set_spacing(p, after=60 if e.get("bullets") or j < len(lines) - 1 else 0)
            for b in e.get("bullets", []):
                self.para("bullet", [run(b, color="000000")])

    def education(self, s):
        for i, e in enumerate(s["items"]):
            runs = [run(e["degree"], bold=True)]
            if e.get("detail"):
                runs.append(run(" | " + e["detail"], bold=True))
            if e.get("date"):
                runs += [tab_run(bold=True), run(e["date"], bold=True)]
            p = self.para("edu1", runs)
            if i:
                set_spacing(p, before=160)
            p = self.para("edu2", [run(e["institution"])])
            set_spacing(p, after=0)

    def items(self, s):
        for e in s["items"]:
            if isinstance(e, str):
                runs = bold_split(e)
            else:
                name = e["name"].rstrip(",")
                runs = [run(name + ",", bold=True)]
                if e.get("details"):
                    runs.append(run(" " + e["details"]))
            self.para("item", runs)

    def languages(self, s):
        runs = []
        for i, l in enumerate(s["items"]):
            if i:
                runs.append(run(" | "))
            runs += [run(l["language"].rstrip(":") + ":", bold=True),
                     run(" " + l["level"])]
        self.para("lang", runs)

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
