"""DMS report compiler.

Takes a copy of the DMS template (.docx or .docm) and inserts the chapters of a
SCADE report into it, at the paragraph containing the marker {{SCADE_CHAPTERS}}.

  * The template is copied part by part, so its front pages, headers, footers,
    macros and styles are untouched.
  * From the SCADE report only the chapters are taken: everything from the first
    level-1 heading to the end. Tables, lists, images, hyperlinks, bookmarks and
    cross-references are carried over.
  * Chapters written by hand in the template are kept. {{placeholders}} in the
    template (body, headers, footers) are filled per report, and blocks between
    {{#if name}} and {{/if}} paragraphs are kept or dropped per report.

Usage:
  python dms_compiler.py --template T.docx --batch reports.xlsx --out outdir
  python dms_compiler.py --template T.docx --scade S.docx --output O.docx \
                         --set system_name="Fuel Pump" --set revision=A
"""
import argparse
import copy
import datetime
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
REL_NUMBERING = R + "/numbering"
CT_NUMBERING = ("application/vnd.openxmlformats-officedocument."
                "wordprocessingml.numbering+xml")

DEFAULT_CONFIG = {
    # paragraph in the template that is replaced by the SCADE chapters
    "marker": "SCADE_CHAPTERS",
    # optional regex: first chapter to take (default: first level-1 heading)
    "start_heading": None,
    # optional regex: level-1 heading at which to stop (it is not included)
    "end_heading": None,
    # SCADE style name -> template style name, for styles with different names
    "style_map": {},
    # "auto": remove chapter numbers typed by hand in SCADE headings when the
    # template numbers its headings automatically. true / false to force.
    "strip_typed_numbers": "auto",
    # keep section breaks of the SCADE report (e.g. landscape pages)
    "keep_section_breaks": False,
}

PLACEHOLDER = re.compile(r"\{\{\s*([A-Za-z_]\w*)\s*\}\}")
IF_OPEN = re.compile(r"^\{\{\s*#if\s+(!?)\s*([A-Za-z_]\w*)\s*\}\}$")
IF_CLOSE = re.compile(r"^\{\{\s*/if\s*\}\}$")
TYPED_NUMBER = re.compile(r"^\s*\d+(\.\d+)*\.?\s+(?=\S)")
TRUE_WORDS = {"true", "yes", "y", "1", "x", "si", "sì"}


def w(tag):
    return f"{{{W}}}{tag}"


def wval(el, default=None):
    return el.get(w("val"), default) if el is not None else default


# ============================================================== package
class Package:
    """A .docx/.docm zip: raw parts, with XML parts parsed on demand."""

    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(path) as z:
            self.order = z.namelist()
            self.parts = {n: z.read(n) for n in self.order}
        self._xml = {}
        ct = self.xml("[Content_Types].xml")
        main = [o.get("PartName") for o in ct.findall(f"{{{CT}}}Override")
                if o.get("ContentType", "").endswith("main+xml")]
        if not main:
            raise ValueError(f"{path}: not a Word document")
        self.doc_name = main[0].lstrip("/")
        self.doc_dir = posixpath.dirname(self.doc_name)

    def xml(self, name):
        if name not in self._xml:
            self._xml[name] = etree.fromstring(self.parts[name])
        return self._xml[name]

    def set_xml(self, name, root):
        self._xml[name] = root
        if name not in self.parts:
            self.parts[name] = b""
            self.order.append(name)

    def add_raw(self, name, data):
        self.parts[name] = data
        self.order.append(name)

    @staticmethod
    def rels_name(part):
        d, b = posixpath.split(part)
        return posixpath.join(d, "_rels", b + ".rels")

    def rels(self, part, create=False):
        name = self.rels_name(part)
        if name not in self.parts:
            if not create:
                return None
            self.set_xml(name, etree.Element(f"{{{PKG_REL}}}Relationships",
                                             nsmap={None: PKG_REL}))
        return self.xml(name)

    def related(self, rel_type):
        """Name of the part related to the main document by rel_type."""
        rels = self.rels(self.doc_name)
        for r in rels if rels is not None else []:
            if r.get("Type") == rel_type and r.get("TargetMode") != "External":
                return posixpath.normpath(
                    posixpath.join(self.doc_dir, r.get("Target")))
        return None

    def content_type(self, name):
        ct = self.xml("[Content_Types].xml")
        for o in ct.findall(f"{{{CT}}}Override"):
            if o.get("PartName") == "/" + name:
                return o.get("ContentType"), True
        ext = name.rsplit(".", 1)[-1].lower()
        for d in ct.findall(f"{{{CT}}}Default"):
            if d.get("Extension", "").lower() == ext:
                return d.get("ContentType"), False
        return None, False

    def register_type(self, name, ctype, override):
        if not ctype:
            return
        ct = self.xml("[Content_Types].xml")
        if override:
            etree.SubElement(ct, f"{{{CT}}}Override", PartName="/" + name,
                             ContentType=ctype)
            return
        ext = name.rsplit(".", 1)[-1]
        if self.content_type(name)[0] is None:
            ct.insert(0, etree.Element(f"{{{CT}}}Default", Extension=ext,
                                       ContentType=ctype))

    def unique_name(self, name):
        if name not in self.parts:
            return name
        stem, dot, ext = name.rpartition(".")
        i = 1
        while f"{stem}_s{i}{dot}{ext}" in self.parts:
            i += 1
        return f"{stem}_s{i}{dot}{ext}"

    def save(self, path):
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for name in self.order:
                if name in self._xml:
                    data = etree.tostring(self._xml[name], xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                else:
                    data = self.parts[name]
                z.writestr(name, data)


# ============================================================== styles
class Styles:
    def __init__(self, pkg):
        self.pkg = pkg
        name = pkg.related(R + "/styles")
        self.root = pkg.xml(name)
        self.by_id = {}
        self.by_name = {}
        for s in self.root.findall(w("style")):
            self._index(s)

    def _index(self, s):
        sid = s.get(w("styleId"))
        self.by_id[sid] = s
        nm = wval(s.find(w("name")), sid)
        self.by_name[(s.get(w("type")), nm.lower())] = sid

    def name(self, sid):
        s = self.by_id.get(sid)
        return wval(s.find(w("name")), sid) if s is not None else sid

    def chain(self, sid):
        seen = set()
        while sid in self.by_id and sid not in seen:
            seen.add(sid)
            s = self.by_id[sid]
            yield s
            sid = wval(s.find(w("basedOn")))

    def heading_level(self, sid):
        """1-9 if the style is a heading (by name or outline level), else None."""
        for s in self.chain(sid):
            name = wval(s.find(w("name")), "").lower()
            m = re.fullmatch(r"heading (\d)", name)
            if m:
                return int(m.group(1))
            if name.startswith("toc"):
                return None  # "TOC Heading" is based on Heading 1 but is no chapter
            lvl = s.find(f"{w('pPr')}/{w('outlineLvl')}")
            if lvl is not None and wval(lvl, "9").isdigit():
                return int(wval(lvl)) + 1 if int(wval(lvl)) < 9 else None
        return None

    def num_id(self, sid):
        for s in self.chain(sid):
            n = s.find(f"{w('pPr')}/{w('numPr')}/{w('numId')}")
            if n is not None:
                return wval(n)
        return None

    def default_paragraph_style(self):
        for s in self.root.findall(w("style")):
            if s.get(w("type")) == "paragraph" and s.get(w("default")) in ("1", "true"):
                return s.get(w("styleId"))
        return None


# ============================================================== text helpers
def own_pieces(p):
    """Text pieces of paragraph p (not of paragraphs nested in it).
    Returns [(element, text)], where tabs count as a '\\t' piece."""
    out = []
    for el in p.iter(w("t"), w("tab")):
        a = el.getparent()
        while a is not None and a.tag != w("p"):
            a = a.getparent()
        if a is not p:
            continue
        if el.tag == w("tab"):
            if el.getparent().tag == w("r"):
                out.append((el, "\t"))
        else:
            out.append((el, el.text or ""))
    return out


def para_text(p):
    return "".join(t for _, t in own_pieces(p))


def set_text(t, text):
    t.text = text
    t.set(XML_SPACE, "preserve")


def fill_placeholders(root, values, missing, skip=None):
    """Replace {{name}} in every paragraph, even when split across runs."""
    for p in root.iter(w("p")):
        if p is skip:
            continue
        ts = [(e, t) for e, t in own_pieces(p) if e.tag == w("t")]
        text = "".join(t for _, t in ts)
        if "{{" not in text:
            continue
        offs, pos = [], 0
        for _, t in ts:
            offs.append(pos)
            pos += len(t)

        def locate(i):
            k = max(j for j, o in enumerate(offs) if o <= i)
            return k, i - offs[k]

        for m in reversed(list(PLACEHOLDER.finditer(text))):
            key = m.group(1)
            if key not in values:
                missing.add(key)
                continue
            (i, si), (j, ej) = locate(m.start()), locate(m.end() - 1)
            ej += 1
            ti, tj = ts[i][0], ts[j][0]
            if i == j:
                set_text(ti, ti.text[:si] + values[key] + ti.text[ej:])
            else:
                set_text(tj, tj.text[ej:])
                for k in range(i + 1, j):
                    set_text(ts[k][0], "")
                set_text(ti, ti.text[:si] + values[key])


def apply_conditionals(body, values, log):
    """Body-level {{#if name}} ... {{/if}} blocks ({{#if !name}} negates)."""
    keep, inside = True, False
    for el in list(body):
        text = para_text(el).strip() if el.tag == w("p") else ""
        m = IF_OPEN.match(text)
        if m:
            truth = str(values.get(m.group(2), "")).strip().lower() in TRUE_WORDS
            keep = truth != bool(m.group(1))
            inside = True
            log.append(f"  condition '{m.group(1)}{m.group(2)}': "
                       f"{'kept' if keep else 'removed'}")
            body.remove(el)
        elif IF_CLOSE.match(text):
            keep, inside = True, False
            body.remove(el)
        elif inside and not keep:
            body.remove(el)
    if inside:
        log.append("  WARNING: {{#if}} without closing {{/if}}")


def count_content(elements):
    c = {"h1": 0, "h2": 0, "h3": 0, "tables": 0, "links": 0, "images": 0}
    for top in elements:
        for el in top.iter():
            if not isinstance(el.tag, str):
                continue
            if el.tag == w("tbl"):
                c["tables"] += 1
            elif el.tag == w("hyperlink"):
                c["links"] += 1
            elif el.tag.endswith("}blip") or el.tag.endswith("}imagedata"):
                c["images"] += 1
    return c


# ============================================================== transplant
class Transplant:
    """Moves a slice of the SCADE document body into the template package."""

    def __init__(self, tpl, src, cfg, log):
        self.tpl, self.src, self.cfg, self.log = tpl, src, cfg, log
        self.tstyles, self.sstyles = Styles(tpl), Styles(src)
        self.style_memo = {}
        self.num_memo = {}
        self.part_memo = {}
        self.rel_memo = {}
        self.rel_counter = 0
        self.tpl_heading_auto = any(
            self.tstyles.num_id(self.tpl_heading(l)) not in (None, "0")
            for l in (1, 2, 3) if self.tpl_heading(l))

    # ---- slice selection
    def tpl_heading(self, level):
        return self.tstyles.by_name.get(("paragraph", f"heading {level}"))

    def src_level(self, p):
        if p.tag != w("p"):
            return None
        sid = wval(p.find(f"{w('pPr')}/{w('pStyle')}"))
        return self.sstyles.heading_level(sid) if sid else None

    def select(self):
        body = self.src.xml(self.src.doc_name).find(w("body"))
        items = [e for e in body if e.tag != w("sectPr")]
        start_rx = self.cfg["start_heading"] and re.compile(self.cfg["start_heading"])
        end_rx = self.cfg["end_heading"] and re.compile(self.cfg["end_heading"])
        start = end = None
        for i, el in enumerate(items):
            if self.src_level(el) != 1:
                continue
            text = para_text(el)
            if start is None:
                if not start_rx or start_rx.search(text):
                    start = i
            elif end_rx and end_rx.search(text):
                end = i
                break
        if start is None:
            raise ValueError("no level-1 heading found in the SCADE report "
                             "(check heading styles or 'start_heading')")
        chosen = items[start:end]
        self.source_counts = self.count(chosen)
        return [copy.deepcopy(e) for e in chosen]

    def count(self, elements):
        c = count_content(elements)
        for top in elements:
            lvl = self.src_level(top)
            if lvl and lvl <= 3:
                c[f"h{lvl}"] += 1
        return c

    # ---- styles
    def map_style(self, sid, stype):
        key = (sid, stype)
        if key in self.style_memo:
            return self.style_memo[key]
        s = self.sstyles.by_id.get(sid)
        if s is None:
            return sid
        name = self.sstyles.name(sid)
        target = None
        for k, v in self.cfg["style_map"].items():
            if k.lower() == name.lower():
                target = self.tstyles.by_name.get((stype, v.lower()))
                if target is None:
                    self.log.append(f"  WARNING: style_map target '{v}' "
                                    "not found in template")
        if target is None and stype == "paragraph":
            lvl = self.sstyles.heading_level(sid)
            if lvl:
                target = self.tpl_heading(lvl)
        if target is None:
            target = self.tstyles.by_name.get((stype, name.lower()))
        if target is not None:
            self.style_memo[key] = target
            return target
        # not in the template: import the SCADE style definition
        new_id = sid
        while new_id in self.tstyles.by_id:
            new_id = "Scade" + new_id
        self.style_memo[key] = new_id
        new = copy.deepcopy(s)
        new.set(w("styleId"), new_id)
        new.attrib.pop(w("default"), None)
        for tag in ("basedOn", "next", "link"):
            ref = new.find(w(tag))
            if ref is not None:
                rtype = stype
                if tag == "link":
                    rtype = "character" if stype == "paragraph" else "paragraph"
                ref.set(w("val"), self.map_style(wval(ref), rtype))
        for n in new.iter(w("numId")):
            n.set(w("val"), self.map_num(wval(n)))
        self.tstyles.root.append(new)
        self.tstyles._index(new)
        self.log.append(f"  imported style '{name}' from SCADE report")
        return new_id

    # ---- numbering
    def tpl_numbering(self):
        name = self.tpl.related(REL_NUMBERING)
        if name is None:
            name = posixpath.join(self.tpl.doc_dir, "numbering.xml")
            self.tpl.set_xml(name, etree.Element(w("numbering"), nsmap={"w": W}))
            self.tpl.register_type(name, CT_NUMBERING, True)
            self.add_rel(REL_NUMBERING, "numbering.xml")
        return self.tpl.xml(name)

    def map_num(self, num_id):
        if num_id in (None, "0"):
            return num_id
        if num_id in self.num_memo:
            return self.num_memo[num_id]
        sname = self.src.related(REL_NUMBERING)
        snum = self.src.xml(sname) if sname else None
        num = None
        if snum is not None:
            num = next((n for n in snum.findall(w("num"))
                        if n.get(w("numId")) == num_id), None)
        if num is None:
            return num_id
        abs_id = wval(num.find(w("abstractNumId")))
        absn = next((a for a in snum.findall(w("abstractNum"))
                     if a.get(w("abstractNumId")) == abs_id), None)
        tnum = self.tpl_numbering()
        ids = lambda tag, attr: [int(e.get(w(attr))) for e in tnum.findall(w(tag))]
        new_abs = str(max(ids("abstractNum", "abstractNumId"), default=0) + 1)
        new_num = str(max(ids("num", "numId"), default=0) + 1)
        self.num_memo[num_id] = new_num
        if absn is not None:
            a = copy.deepcopy(absn)
            a.set(w("abstractNumId"), new_abs)
            for tag in ("pStyle", "styleLink", "numStyleLink"):
                for e in list(a.iter(w(tag))):
                    e.getparent().remove(e)
            first = tnum.find(w("num"))
            if first is not None:
                first.addprevious(a)
            else:
                tnum.append(a)
        n = copy.deepcopy(num)
        n.set(w("numId"), new_num)
        n.find(w("abstractNumId")).set(w("val"), new_abs)
        cleanup = tnum.find(w("numIdMacAtCleanup"))
        if cleanup is not None:
            cleanup.addprevious(n)
        else:
            tnum.append(n)
        return new_num

    # ---- relationships and parts
    def add_rel(self, rtype, target, external=False, owner=None):
        rels = self.tpl.rels(owner or self.tpl.doc_name, create=True)
        used = {r.get("Id") for r in rels}
        while True:
            self.rel_counter += 1
            rid = f"rIdS{self.rel_counter}"
            if rid not in used:
                break
        attrs = {"Id": rid, "Type": rtype, "Target": target}
        if external:
            attrs["TargetMode"] = "External"
        etree.SubElement(rels, f"{{{PKG_REL}}}Relationship", **attrs)
        return rid

    def copy_part(self, sname):
        """Copy a part (image, chart, ...) and, recursively, what it refers to."""
        if sname in self.part_memo:
            return self.part_memo[sname]
        tname = self.tpl.unique_name(sname)
        self.part_memo[sname] = tname
        self.tpl.add_raw(tname, self.src.parts[sname])
        ctype, override = self.src.content_type(sname)
        self.tpl.register_type(tname, ctype, override)
        srels = self.src.rels(sname)
        if srels is not None:
            new = copy.deepcopy(srels)
            for r in new:
                if r.get("TargetMode") == "External":
                    continue
                sub = posixpath.normpath(
                    posixpath.join(posixpath.dirname(sname), r.get("Target")))
                if sub in self.src.parts:
                    r.set("Target", posixpath.relpath(
                        self.copy_part(sub), posixpath.dirname(tname)))
            self.tpl.set_xml(self.tpl.rels_name(tname), new)
        return tname

    def map_rel(self, rid):
        if rid in self.rel_memo:
            return self.rel_memo[rid]
        srels = self.src.rels(self.src.doc_name)
        rel = next((r for r in srels if r.get("Id") == rid), None)
        if rel is None:
            return None
        if rel.get("TargetMode") == "External":
            new = self.add_rel(rel.get("Type"), rel.get("Target"), external=True)
        else:
            sname = posixpath.normpath(
                posixpath.join(self.src.doc_dir, rel.get("Target")))
            if sname not in self.src.parts:
                return None
            tname = self.copy_part(sname)
            new = self.add_rel(rel.get("Type"),
                               posixpath.relpath(tname, self.tpl.doc_dir))
        self.rel_memo[rid] = new
        return new

    # ---- the per-element fix-ups
    def strip_typed_number(self, p):
        pieces = own_pieces(p)
        m = TYPED_NUMBER.match("".join(t for _, t in pieces))
        if not m:
            return False
        left = m.end()
        for el, text in pieces:
            if left <= 0:
                break
            if el.tag == w("tab"):
                el.getparent().remove(el)
                left -= 1
            else:
                cut = min(left, len(text))
                set_text(el, text[cut:])
                left -= cut
        return True

    def process(self, elements):
        cfg, log = self.cfg, self.log
        tdoc = self.tpl.xml(self.tpl.doc_name)
        holder = etree.Element(w("body"))
        for e in elements:
            holder.append(e)

        # section breaks
        for sect in list(holder.iter(w("sectPr"))):
            if cfg["keep_section_breaks"]:
                for ref in list(sect):
                    if ref.tag in (w("headerReference"), w("footerReference")):
                        sect.remove(ref)
                continue
            ppr = sect.getparent()
            p = ppr.getparent()
            ppr.remove(sect)
            if p.tag == w("p") and not para_text(p) and len(p) == 1 and len(ppr) == 0:
                p.getparent().remove(p)

        # notes and comments are not carried over
        dropped = 0
        for tag in ("footnoteReference", "endnoteReference", "commentReference",
                    "commentRangeStart", "commentRangeEnd"):
            for e in list(holder.iter(w(tag))):
                e.getparent().remove(e)
                dropped += tag.endswith("Reference")
        if dropped:
            log.append(f"  WARNING: {dropped} footnote/endnote/comment "
                       "reference(s) dropped (not supported yet)")

        # headings: template numbering, typed numbers
        strip = cfg["strip_typed_numbers"]
        stripped = 0
        for p in holder.iter(w("p")):
            ppr = p.find(w("pPr"))
            sid = wval(ppr.find(w("pStyle"))) if ppr is not None else None
            if not sid or not self.sstyles.heading_level(sid):
                continue
            direct = ppr.find(w("numPr"))
            src_auto = (wval(direct.find(w("numId"))) if direct is not None
                        else self.sstyles.num_id(sid)) not in (None, "0")
            if direct is not None and self.tpl_heading_auto:
                ppr.remove(direct)
            do_strip = strip is True or (
                strip == "auto" and self.tpl_heading_auto and not src_auto)
            if do_strip and self.strip_typed_number(p):
                stripped += 1
        if stripped:
            log.append(f"  removed typed chapter numbers from {stripped} headings")

        # styles
        for tag, stype in (("pStyle", "paragraph"), ("rStyle", "character"),
                           ("tblStyle", "table")):
            for e in holder.iter(w(tag)):
                e.set(w("val"), self.map_style(wval(e), stype))

        # numbering (lists)
        for n in holder.iter(w("numId")):
            n.set(w("val"), self.map_num(wval(n)))

        # relationships: hyperlinks, images, embedded objects
        lost = 0
        for el in holder.iter():
            if not isinstance(el.tag, str):
                continue
            for attr, val in list(el.attrib.items()):
                if attr.startswith(f"{{{R}}}"):
                    new = self.map_rel(val)
                    if new is None:
                        lost += 1
                        del el.attrib[attr]
                    else:
                        el.set(attr, new)
        if lost:
            log.append(f"  WARNING: {lost} broken relationship(s) in SCADE report")

        # bookmarks: unique ids and names, references kept in sync
        t_names = {b.get(w("name")) for b in tdoc.iter(w("bookmarkStart"))}
        next_id = 1 + max([int(b.get(w("id"), "0"))
                           for b in tdoc.iter(w("bookmarkStart"))], default=0)
        id_map, renames = {}, {}
        for b in holder.iter(w("bookmarkStart")):
            id_map[b.get(w("id"))] = str(next_id)
            b.set(w("id"), str(next_id))
            next_id += 1
            name = b.get(w("name"))
            if name in t_names:
                new = name
                while new in t_names:
                    new += "_s"
                renames[name] = new
                b.set(w("name"), new)
            t_names.add(b.get(w("name")))
        for b in list(holder.iter(w("bookmarkEnd"))):
            if b.get(w("id")) in id_map:
                b.set(w("id"), id_map[b.get(w("id"))])
            else:
                b.getparent().remove(b)
        if renames:
            rx = re.compile(r"(?<![\w])(" + "|".join(map(re.escape, renames))
                            + r")(?![\w])")
            for h in holder.iter(w("hyperlink")):
                a = h.get(w("anchor"))
                if a in renames:
                    h.set(w("anchor"), renames[a])
            for it in holder.iter(w("instrText")):
                if it.text:
                    it.text = rx.sub(lambda m: renames[m.group(1)], it.text)
            for f in holder.iter(w("fldSimple")):
                f.set(w("instr"), rx.sub(lambda m: renames[m.group(1)],
                                         f.get(w("instr"), "")))

        # drawing ids must be unique in the document
        tag = f"{{{WP}}}docPr"
        next_dp = 1 + max([int(d.get("id", "0")) for d in tdoc.iter(tag)], default=0)
        for d in holder.iter(tag):
            d.set("id", str(next_dp))
            next_dp += 1

        return list(holder)


# ============================================================== compile
def merged_root(root, other):
    """Rebuild root so it also declares the namespace prefixes of other."""
    nsmap = dict(root.nsmap)
    for prefix, uri in other.nsmap.items():
        if prefix not in nsmap and uri not in nsmap.values():
            nsmap[prefix] = uri
    new = etree.Element(root.tag, nsmap=nsmap)
    for k, v in root.attrib.items():
        new.set(k, v)
    for child in list(root):
        new.append(child)
    return new


def find_marker(body, marker):
    rx = re.compile(r"^\{\{\s*" + re.escape(marker) + r"\s*\}\}$")
    for p in body.iter(w("p")):
        if rx.match(para_text(p).strip()):
            return p
    return None


def set_update_fields(pkg):
    """Ask Word to refresh TOC/fields when the file is opened."""
    name = pkg.related(R + "/settings")
    if not name:
        return
    settings = pkg.xml(name)
    if settings.find(w("updateFields")) is None:
        el = etree.Element(w("updateFields"))
        el.set(w("val"), "true")
        settings.insert(0, el)


def compile_report(template, scade, output, values, cfg, update_on_open=True):
    """Build one report. Returns (ok, log lines)."""
    log = []
    tpl, src = Package(template), Package(scade)
    values = {k: format_value(v) for k, v in values.items()}

    root = merged_root(tpl.xml(tpl.doc_name), src.xml(src.doc_name))
    tpl.set_xml(tpl.doc_name, root)
    body = root.find(w("body"))

    apply_conditionals(body, values, log)
    marker = find_marker(body, cfg["marker"])
    if marker is None:
        log.append(f"  WARNING: marker {{{{{cfg['marker']}}}}} not found; "
                   "chapters appended at the end of the template")

    # placeholders: body, headers, footers, notes
    missing = set()
    fill_placeholders(root, values, missing, skip=marker)
    for name in list(tpl.parts):
        if re.fullmatch(r"word/(header|footer|footnotes|endnotes)\d*\.xml", name):
            fill_placeholders(tpl.xml(name), values, missing)
    tpl_counts = count_content([body])

    tr = Transplant(tpl, src, cfg, log)
    new = tr.process(tr.select())
    if marker is not None:
        for el in new:
            marker.addprevious(el)
        marker.getparent().remove(marker)
    else:
        last = body.find(w("sectPr"))
        for el in new:
            last.addprevious(el) if last is not None else body.append(el)

    if update_on_open:
        set_update_fields(tpl)
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    tpl.save(output)

    # ---- checks on the saved file
    ok = True
    sc = tr.source_counts
    log.append("  chapters taken: {h1} (subchapters: {h2} + {h3}), tables: "
               "{tables}, links: {links}, images: {images}".format(**sc))
    out = Package(output)
    obody = out.xml(out.doc_name).find(w("body"))
    oc = count_content([obody])
    for key in ("tables", "links", "images"):
        if oc[key] != tpl_counts[key] + sc[key]:
            ok = False
            log.append(f"  ERROR: {key}: expected {tpl_counts[key] + sc[key]}, "
                       f"found {oc[key]} in output")
    rel_ids = {r.get("Id"): r for r in out.rels(out.doc_name)}
    for el in obody.iter():
        if not isinstance(el.tag, str):
            continue
        for attr, val in el.attrib.items():
            if attr.startswith(f"{{{R}}}"):
                rel = rel_ids.get(val)
                target = rel is not None and (
                    rel.get("TargetMode") == "External" or posixpath.normpath(
                        posixpath.join(out.doc_dir, rel.get("Target"))) in out.parts)
                if not target:
                    ok = False
                    log.append(f"  ERROR: dangling relationship {val}")
    if missing:
        ok = False
        log.append("  ERROR: no value for placeholder(s): "
                   + ", ".join(sorted(missing)))
    return ok, log


def format_value(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "yes" if v else "no"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.datetime) and v.time() == datetime.time(0):
        return v.date().isoformat()
    return str(v)


# ============================================================== Word refresh
def refresh_with_word(paths, pdf=False):
    """Open each file in Word (macros disabled) and update TOC, lists and fields."""
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3  # never run macros during automation
        for path in paths:
            doc = word.Documents.Open(str(Path(path).resolve()),
                                      ConfirmConversions=False,
                                      AddToRecentFiles=False)
            doc.Fields.Update()
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            for i in range(1, doc.TablesOfFigures.Count + 1):
                doc.TablesOfFigures(i).Update()
            for s in range(1, doc.Sections.Count + 1):
                for k in (1, 2, 3):
                    doc.Sections(s).Headers(k).Range.Fields.Update()
                    doc.Sections(s).Footers(k).Range.Fields.Update()
            doc.Save()
            if pdf:
                doc.ExportAsFixedFormat(
                    str(Path(path).resolve().with_suffix(".pdf")), 17)
            doc.Close(False)
    finally:
        word.Quit()


# ============================================================== CLI
def read_batch(path):
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    if "scade_report" not in header:
        raise ValueError("batch sheet needs a 'scade_report' column")
    jobs = []
    for row in rows[1:]:
        rec = {h: v for h, v in zip(header, row) if h}
        if rec.get("scade_report"):
            jobs.append(rec)
    return jobs


def main(argv=None, config=None):
    ap =argparse.ArgumentParser(description="Compile DMS reports from SCADE reports")
    ap.add_argument("--template", required=True, help="DMS template (.docx/.docm)")
    ap.add_argument("--batch", help="Excel sheet: one row per report")
    ap.add_argument("--scade", help="single SCADE report (instead of --batch)")
    ap.add_argument("--output", help="output file for --scade")
    ap.add_argument("--set", action="append", default=[], metavar="NAME=VALUE",
                    help="placeholder value for --scade (repeatable)")
    ap.add_argument("--out", default="output", help="output folder for --batch")
    ap.add_argument("--config", help="JSON file overriding the default settings")
    ap.add_argument("--no-word", action="store_true",
                    help="do not use Word to refresh the table of contents")
    ap.add_argument("--pdf", action="store_true", help="also export a PDF (needs Word)")
    args = ap.parse_args(argv)

    cfg = dict(DEFAULT_CONFIG)
    cfg.update(config or {})
    if args.config:
        cfg.update(json.loads(Path(args.config).read_text(encoding="utf-8")))
    template = Path(args.template)

    jobs = []
    if args.batch:
        base = Path(args.batch).parent
        for rec in read_batch(args.batch):
            scade = base / str(rec.pop("scade_report"))
            name = rec.pop("output", None) or f"{scade.stem}_DMS{template.suffix}"
            jobs.append((scade, Path(args.out) / str(name), rec))
    elif args.scade:
        vals = dict(s.split("=", 1) for s in args.set)
        out = Path(args.output or f"{Path(args.scade).stem}_DMS{template.suffix}")
        jobs.append((Path(args.scade), out, vals))
    else:
        ap.error("give --batch or --scade")

    done, failed = [], 0
    for scade, out, vals in jobs:
        print(f"{scade.name} -> {out}")
        try:
            ok, log = compile_report(template, scade, out, vals, cfg,
                                     update_on_open=args.no_word)
        except Exception as exc:  # keep going with the other reports
            ok, log = False, [f"  ERROR: {exc}"]
        else:
            done.append(out)
        print("\n".join(log))
        print("  OK" if ok else "  FAILED")
        failed += not ok

    if done and not args.no_word:
        try:
            refresh_with_word(done, pdf=args.pdf)
            print("Table of contents and fields refreshed with Word.")
        except Exception as exc:
            print(f"WARNING: Word refresh failed ({exc}). Open the reports and "
                  "press Ctrl+A, F9 to update the table of contents.")
    print(f"{len(jobs) - failed} of {len(jobs)} reports compiled without errors.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
