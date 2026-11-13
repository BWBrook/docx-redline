# /// script
# requires-python = ">=3.10"
# dependencies = ["lxml"]
# ///
"""Add word-level tracked changes and comments to a .docx under a chosen author.

Existing comments and revisions are left untouched. Edits and comment anchors
must fall on text that is not already inside someone's tracked insertion or
deletion; the script refuses rather than guessing.

    uv run redline.py IN.docx --dump                 # visible text per paragraph
    uv run redline.py IN.docx OPS.json OUT.docx --author "J. Citizen" --initials JC

OPS.json is a list of operations, applied in order. Each locates text in the
current view (insertions shown, deletions hidden) by before + target + after,
which must occur exactly once in the document:

    {"op": "rep", "b": "context before ", "old": "text to delete", "new": "text to insert", "a": " context after"}
    {"op": "com", "b": "context before ", "t": "anchored text", "a": " context after", "c": "comment text"}

"old" or "new" may be empty for a pure insertion or deletion. A comment placed
before a "rep" on the same text ends up spanning the deletion and insertion.
"""

import argparse
import copy
import json
import random
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"


def w(tag):
    return f"{{{W}}}{tag}"


TRACKED = {w("ins"), w("del"), w("moveFrom"), w("moveTo")}
HIDDEN = {w("del"), w("moveFrom")}


def hex_id():
    return f"{random.randint(1, 0x7FFFFFFE):08X}"


class Redliner:
    def __init__(self, parts, author, initials):
        self.parts = parts
        self.author, self.initials = author, initials
        self.local_ts = datetime.now().strftime("%Y-%m-%dT%H:%M:00Z")
        self.utc_ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:00Z")
        self.doc = self._parse("word/document.xml")
        self.trees = {"word/document.xml": self.doc}
        ids = [int(v) for e in self.doc.iter() if (v := e.get(w("id"))) and v.lstrip("-").isdigit()]
        self.next_rev = max(ids + [0]) + 1000

    def _parse(self, name):
        return etree.fromstring(self.parts[name])

    def _tree(self, name):
        if name not in self.trees:
            self.trees[name] = self._parse(name) if name in self.parts else None
        return self.trees[name]

    def rev_attrs(self):
        self.next_rev += 1
        return {w("id"): str(self.next_rev), w("author"): self.author, w("date"): self.local_ts}

    def segments(self, p):
        segs = []
        for r in p.iter(w("r")):
            anc = list(r.iterancestors())
            if any(a.tag in HIDDEN for a in anc) or next(r.iterancestors(w("p"))) is not p:
                continue
            editable = not any(a.tag in TRACKED for a in anc)
            for child in r:
                if child.tag == w("t"):
                    segs.append((r, child, child.text or "", editable))
                elif child.tag == w("tab"):
                    segs.append((r, child, "\t", False))
        return segs

    def text(self, p):
        return "".join(s[2] for s in self.segments(p))

    def locate(self, before, target, after):
        needle = before + target + after
        hits = []
        for p in self.doc.iter(w("p")):
            t = self.text(p)
            i = t.find(needle)
            while i != -1:
                hits.append((p, i + len(before)))
                i = t.find(needle, i + 1)
        if len(hits) != 1:
            raise ValueError(f"{len(hits)} matches for {needle!r}")
        return hits[0]

    def _normalise(self, p):
        for r in list(p.iter(w("r"))):
            content = [c for c in r if c.tag != w("rPr")]
            if len(content) > 1 and any(c.tag == w("t") for c in content):
                rpr = r.find(w("rPr"))
                for c in content:
                    nr = etree.Element(w("r"))
                    if rpr is not None:
                        nr.append(copy.deepcopy(rpr))
                    nr.append(c)
                    r.addprevious(nr)
                r.getparent().remove(r)

    def _split(self, p, offset):
        pos = 0
        for r, t, s, _ in self.segments(p):
            if pos < offset < pos + len(s):
                k = offset - pos
                r2 = copy.deepcopy(r)
                t.text, t2 = s[:k], r2.find(w("t"))
                t2.text = s[k:]
                t.set(XML_SPACE, "preserve")
                t2.set(XML_SPACE, "preserve")
                r.addnext(r2)
                return
            pos += len(s)

    def _runs(self, p, s, e):
        out, pos = [], 0
        for r, _, txt, editable in self.segments(p):
            a, b = pos, pos + len(txt)
            if s <= a and b <= e and b > a:
                if not editable:
                    raise ValueError(f"range touches existing tracked text {txt!r}")
                out.append(r)
            elif a < e and b > s:
                raise ValueError("range does not fall on run boundaries")
            pos = b
        return out

    def _prepare(self, before, target, after):
        p, s = self.locate(before, target, after)
        e = s + len(target)
        self._normalise(p)
        self._split(p, s)
        self._split(p, e)
        return p, s, e

    def replace(self, before, old, new, after=""):
        p, s, e = self._prepare(before, old, after)
        if old:
            runs = self._runs(p, s, e)
            for r in runs:
                d = etree.Element(w("del"), self.rev_attrs())
                r.addprevious(d)
                d.append(r)
                for t in r.findall(w("t")):
                    t.tag = w("delText")
                    t.set(XML_SPACE, "preserve")
            style_run, after_el = runs[-1], runs[-1].getparent()
        else:
            style_run = after_el = None
            pos = 0
            for r, _, txt, _ in self.segments(p):
                pos += len(txt)
                if pos == s and txt:
                    style_run = after_el = r
            if style_run is not None:
                tracked = [a for a in style_run.iterancestors() if a.tag in TRACKED]
                if tracked:
                    after_el = tracked[-1]
            else:
                style_run = self.segments(p)[0][0]
        if not new:
            return
        ins = etree.Element(w("ins"), self.rev_attrs())
        nr = etree.SubElement(ins, w("r"))
        rpr = style_run.find(w("rPr"))
        if rpr is not None:
            rpr = copy.deepcopy(rpr)
            for rs in rpr.findall(w("rStyle")):
                if rs.get(w("val")) == "CommentReference":
                    rpr.remove(rs)
            nr.append(rpr)
        t = etree.SubElement(nr, w("t"))
        t.text = new
        t.set(XML_SPACE, "preserve")
        if after_el is None:
            style_run.addprevious(ins)
        else:
            after_el.addnext(ins)

    def comment(self, before, target, after, text):
        p, s, e = self._prepare(before, target, after)
        runs = self._runs(p, s, e)
        cid = self._add_comment_part(text)
        runs[0].addprevious(etree.Element(w("commentRangeStart"), {w("id"): cid}))
        end = etree.Element(w("commentRangeEnd"), {w("id"): cid})
        runs[-1].addnext(end)
        ref = etree.Element(w("r"))
        etree.SubElement(etree.SubElement(ref, w("rPr")), w("rStyle"), {w("val"): "CommentReference"})
        etree.SubElement(ref, w("commentReference"), {w("id"): cid})
        end.addnext(ref)

    def _add_comment_part(self, text):
        comments = self._tree("word/comments.xml")
        if comments is None:
            comments = self._create_comments_part()
        ids = [int(c.get(w("id"))) for c in comments.iter(w("comment"))]
        cid = str(max(ids + [-1]) + 1)
        para_id, durable = hex_id(), hex_id()
        c = etree.SubElement(comments, w("comment"), {
            w("id"): cid, w("author"): self.author, w("date"): self.local_ts, w("initials"): self.initials})
        cp = etree.SubElement(c, w("p"), {f"{{{W14}}}paraId": para_id, f"{{{W14}}}textId": "77777777"})
        ppr = etree.SubElement(cp, w("pPr"))
        etree.SubElement(ppr, w("pStyle"), {w("val"): "CommentText"})
        ar = etree.SubElement(cp, w("r"))
        etree.SubElement(etree.SubElement(ar, w("rPr")), w("rStyle"), {w("val"): "CommentReference"})
        etree.SubElement(ar, w("annotationRef"))
        tr = etree.SubElement(cp, w("r"))
        t = etree.SubElement(tr, w("t"))
        t.text = text
        t.set(XML_SPACE, "preserve")
        ext = self._tree("word/commentsExtended.xml")
        if ext is not None:
            etree.SubElement(ext, f"{{{W15}}}commentEx", {f"{{{W15}}}paraId": para_id, f"{{{W15}}}done": "0"})
        cids = self._tree("word/commentsIds.xml")
        if cids is not None:
            etree.SubElement(cids, f"{{{W16CID}}}commentId", {
                f"{{{W16CID}}}paraId": para_id, f"{{{W16CID}}}durableId": durable})
        cex = self._tree("word/commentsExtensible.xml")
        if cex is not None:
            etree.SubElement(cex, f"{{{W16CEX}}}commentExtensible", {
                f"{{{W16CEX}}}durableId": durable, f"{{{W16CEX}}}dateUtc": self.utc_ts})
        return cid

    def _create_comments_part(self):
        root = etree.Element(w("comments"), nsmap={"w": W, "w14": W14})
        self.trees["word/comments.xml"] = root
        rels = self._tree("word/_rels/document.xml.rels")
        used = {r.get("Id") for r in rels}
        n = 1
        while f"rId{n}" in used:
            n += 1
        etree.SubElement(rels, f"{{{REL_NS}}}Relationship", {"Id": f"rId{n}", "Type": COMMENTS_REL, "Target": "comments.xml"})
        ct = self._tree("[Content_Types].xml")
        etree.SubElement(ct, f"{{{CT_NS}}}Override", {"PartName": "/word/comments.xml", "ContentType": COMMENTS_CT})
        return root

    def serialise(self):
        for name, tree in self.trees.items():
            if tree is not None:
                self.parts[name] = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        return self.parts


def read_docx(path):
    with zipfile.ZipFile(path) as z:
        infos = [i for i in z.infolist() if not i.is_dir()]
        return infos, {i.filename: z.read(i) for i in infos}


def write_docx(path, infos, parts):
    names = [i.filename for i in infos] + [n for n in parts if n not in {i.filename for i in infos}]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, parts[n])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("ops", nargs="?")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--dump", action="store_true", help="print visible text of each paragraph")
    ap.add_argument("--author", default="J. Citizen", help="name shown on changes and comments (default: J. Citizen)")
    ap.add_argument("--initials", default="JC", help="initials shown on comments (default: JC)")
    args = ap.parse_args()

    infos, parts = read_docx(args.input)
    rl = Redliner(parts, args.author, args.initials)
    if args.dump:
        for i, p in enumerate(rl.doc.iter(w("p"))):
            if (t := rl.text(p)).strip():
                print(i, repr(t))
        return
    if not (args.ops and args.output):
        ap.error("OPS and OUTPUT are required unless --dump")
    if Path(args.output).resolve() == Path(args.input).resolve():
        ap.error("refusing to overwrite the input file")

    failures = 0
    for op in json.loads(Path(args.ops).read_text()):
        try:
            if op["op"] == "rep":
                rl.replace(op.get("b", ""), op["old"], op["new"], op.get("a", ""))
            elif op["op"] == "com":
                rl.comment(op.get("b", ""), op["t"], op.get("a", ""), op["c"])
            else:
                raise ValueError(f"unknown op {op['op']!r}")
        except Exception as ex:
            failures += 1
            print(f"FAIL {ex}: {op}", file=sys.stderr)
    if failures:
        sys.exit(f"{failures} operation(s) failed; nothing written")
    write_docx(args.output, infos, rl.serialise())
    print(f"wrote {args.output}")


if __name__ == "__main__":
    main()
