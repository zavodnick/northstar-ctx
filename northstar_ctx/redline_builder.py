#!/usr/bin/env python3
"""NorthStar redline builder (ctx 3.4.0). Claude reads extract output, never the XML.
  extract IN.docx | accept IN.docx OUT.docx [--keep "A,B"] [--ids 3,7] | apply IN.docx PLAN.json --out OUT.docx
  [--report R.md] | check OUT.docx PLAN.json (also runs at the end of apply)
Plan: {"edits": [...], "questions": true}. Ops: replace (find, with), insert_after (find, text), delete (find),
insert_paragraph_after (text), delete_paragraph, comment (anchor, text, "question": true), reply (parent, text),
resolve (parent). find and anchor match one paragraph's rendered text; interior whitespace is normalized, leading
and trailing spaces are kept. Author and date are constants."""
from __future__ import annotations
import argparse, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from datetime import datetime, timezone
from pathlib import Path
import defusedxml.minidom
from xml.dom import minidom

DOCX_SKILL = Path(os.environ.get("DOCX_SKILL", "/mnt/skills/public/docx/scripts"))
if not (DOCX_SKILL / "comment.py").exists():
    sys.exit(f"HALT: docx skill scripts not found at {DOCX_SKILL} (set DOCX_SKILL).")
sys.path.insert(0, str(DOCX_SKILL))
from office.helpers import rezip, safe_extract  # noqa: E402
import comment as docx_comment  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
AUTHOR = os.environ.get("CTX_AUTHOR", "").strip()
INITIALS = os.environ.get("CTX_INITIALS", "").strip() or "".join(w[0] for w in AUTHOR.split())[:3].upper()
if not AUTHOR:
    sys.exit("HALT: set CTX_AUTHOR (the tracked-change author name) before running.")
WS = " \t\r\n "
TEXT_OPS = {"replace", "insert_after", "delete", "insert_paragraph_after", "delete_paragraph"}
OPS = TEXT_OPS | {"comment", "reply", "resolve"}
KEYS = {"id", "op", "find", "with", "text", "anchor", "parent", "question"}
MARKERS = {"w:commentRangeStart", "w:commentRangeEnd", "w:bookmarkStart", "w:bookmarkEnd"}
PR_CHANGES = ("w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange", "w:tcPrChange",
              "w:trPrChange", "w:tblGridChange")


def now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def halt(msg, code=2):
    print(f"\nHALT: {msg}", file=sys.stderr)
    sys.exit(code)


def norm(s):
    """Collapse interior whitespace runs for matching only; edges are kept (never written back)."""
    return re.sub(r"\s+", " ", s.replace(" ", " "))


def kids(node, *tags):
    return [c for c in node.childNodes if c.nodeType == 1 and (not tags or c.tagName in tags)]


def ancestor(node, tag, stop="w:p"):
    p = node.parentNode
    while p is not None and p.nodeType == 1 and p.tagName != stop:
        if p.tagName == tag:
            return p
        p = p.parentNode
    return None


def unwrap(e):
    par = e.parentNode
    for c in list(e.childNodes):
        par.insertBefore(c, e)
    par.removeChild(e)


def t_text(t):
    s = "".join(c.data for c in t.childNodes if c.nodeType == c.TEXT_NODE)
    return s if t.getAttributeNS(XML_NS, "space") == "preserve" else s.strip(" \t\r\n")


def el_text(e, deleted=False):
    return "".join(t_text(t) for t in e.getElementsByTagName("w:delText" if deleted else "w:t"))


# ---------------------------------------------------------------- document model
class Para:
    def __init__(self, pid, node, in_table):
        self.id, self.node, self.in_table = pid, node, in_table

    def runs(self):
        out = []
        for c in kids(self.node):
            if c.tagName == "w:r":
                out.append(c)
            elif c.tagName in ("w:ins", "w:del", "w:hyperlink", "w:smartTag"):
                out.extend(c.getElementsByTagName("w:r"))
        return out

    def text(self):
        parts = []
        for r in self.runs():
            if ancestor(r, "w:del") is not None:
                continue
            for t in kids(r):
                if t.tagName == "w:t":
                    parts.append(t_text(t))
                elif t.tagName == "w:tab":
                    parts.append("\t")
                elif t.tagName in ("w:br", "w:cr"):
                    parts.append("\n")
        return "".join(parts)

    def changes(self):
        out = []
        for tag in ("w:ins", "w:del"):
            for el in self.node.getElementsByTagName(tag):
                mark = el.parentNode.tagName == "w:rPr"
                txt = "(paragraph mark)" if mark else (el_text(el) or el_text(el, True))
                out.append((tag[2:].upper(), el.getAttribute("w:id"), el.getAttribute("w:author") or "?", txt))
        return out


def load_doc(unpacked):
    dom = defusedxml.minidom.parse(str(unpacked / "word" / "document.xml"))
    body = dom.getElementsByTagName("w:body")[0]
    paras, counts, tbk = [], {"p": 0, "t": 0, "tb": 0}, {}

    def walk(node, in_table, tpos):
        for child in kids(node):
            tag = child.tagName
            if tag == "w:p":
                if tpos and tpos[0] == "TB":
                    tbk[tpos[1]] = tbk.get(tpos[1], 0) + 1
                    pid = f"TB{tpos[1]}" + (f"P{tbk[tpos[1]]}" if tbk[tpos[1]] > 1 else "")
                elif in_table and tpos:
                    t, r, c, k = tpos
                    pid = f"T{t}R{r}C{c}" + (f"P{k}" if k > 0 else "")
                    tpos = (t, r, c, k + 1)
                else:
                    counts["p"] += 1
                    pid = f"P{counts['p']:03d}"
                paras.append(Para(pid, child, in_table))
                walk(child, in_table, tpos)
            elif tag == "w:tbl":
                counts["t"] += 1
                for ri, row in enumerate(kids(child, "w:tr"), 1):
                    for ci, cell in enumerate(kids(row, "w:tc"), 1):
                        walk(cell, True, (counts["t"], ri, ci, 0))
            elif tag == "w:txbxContent":
                counts["tb"] += 1
                walk(child, False, ("TB", counts["tb"]))
            elif tag in ("mc:Fallback", "w:sectPr"):
                continue
            elif tag in ("w:sdt", "w:sdtContent", "w:customXml", "w:smartTag", "mc:AlternateContent", "mc:Choice",
                         "w:drawing", "w:pict", "wp:anchor", "wp:inline", "a:graphic", "a:graphicData", "wps:wsp",
                         "wps:txbx", "v:shape", "v:textbox", "w:r"):
                walk(child, in_table, tpos)

    walk(body, False, None)
    return dom, paras


def doc_para_count(unpacked):
    raw = (unpacked / "word" / "document.xml").read_text(encoding="utf-8", errors="replace")
    body = re.search(r"<w:body>(.*)</w:body>", raw, flags=re.S)
    if not body:
        return -1
    return len(re.findall(r"<w:p[\s>]", re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", body.group(1), flags=re.S)))


def unpack(docx, dest, merge=True):
    dest.mkdir()
    with zipfile.ZipFile(docx) as zf:
        safe_extract(zf, dest)
    if merge:
        subprocess.run([sys.executable, str(DOCX_SKILL / "merge_runs.py"), str(dest)], check=True, capture_output=True)


# ---------------------------------------------------------------- comment threads (satellite parts, minidom)
def comment_maps(unpacked):
    """id -> (author, text, last paraId); paraId -> (parent paraId, done)."""
    cpath, epath = unpacked / "word" / "comments.xml", unpacked / "word" / "commentsExtended.xml"
    cm, ext = {}, {}
    if cpath.exists():
        for c in defusedxml.minidom.parse(str(cpath)).getElementsByTagName("w:comment"):
            ps = c.getElementsByTagName("w:p")
            cm[int(c.getAttribute("w:id"))] = (c.getAttribute("w:author"), " / ".join(el_text(p) for p in ps),
                                               ps[-1].getAttribute("w14:paraId") if ps else "")
    if epath.exists():
        for x in defusedxml.minidom.parse(str(epath)).getElementsByTagName("w15:commentEx"):
            ext[x.getAttribute("w15:paraId")] = (x.getAttribute("w15:paraIdParent"), x.getAttribute("w15:done"))
    return cm, ext


def thread(unpacked, cid):
    cm, ext = comment_maps(unpacked)
    if cid not in cm:
        halt(f"comment {cid} does not exist (extract lists the comments).")
    root_pid = ext.get(cm[cid][2], ("", ""))[0] or cm[cid][2]
    root = next((i for i, v in cm.items() if v[2] == root_pid), cid)
    members = [root] + sorted(i for i, v in cm.items() if i != root and ext.get(v[2], ("", ""))[0] == root_pid)
    return root, root_pid, members


def ext_edit(unpacked, fn):
    epath = unpacked / "word" / "commentsExtended.xml"
    if not epath.exists():
        halt("the document has no commentsExtended.xml; threads cannot be edited.")
    edom = defusedxml.minidom.parse(str(epath))
    n = sum(1 for x in edom.getElementsByTagName("w15:commentEx") if fn(x))
    epath.write_bytes(edom.toxml(encoding="UTF-8"))
    return n


def ensure_people(unpacked, author):
    """Ensure people.xml exists and contains an entry for the given author. Word 365 requires this."""
    ppath = unpacked / "word" / "people.xml"
    tpl = Path(os.environ.get("DOCX_SKILL", "/mnt/skills/public/docx/scripts")) / "templates" / "people.xml"
    if not ppath.exists():
        if tpl.exists():
            import shutil
            shutil.copy(tpl, ppath)
        else:
            ppath.write_text('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                             '<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
                             '</w15:people>', encoding="utf-8")
        # add relationship and content type
        rels_path = unpacked / "word" / "_rels" / "document.xml.rels"
        if rels_path.exists():
            rdom = defusedxml.minidom.parse(str(rels_path))
            has_people = any(r.getAttribute("Target") == "people.xml" for r in rdom.getElementsByTagName("Relationship"))
            if not has_people:
                used = {int(r.getAttribute("Id")[3:]) for r in rdom.getElementsByTagName("Relationship")
                        if r.getAttribute("Id").startswith("rId") and r.getAttribute("Id")[3:].isdigit()}
                nid = (max(used) + 1) if used else 100
                rel = rdom.createElement("Relationship")
                rel.setAttribute("Id", f"rId{nid}")
                rel.setAttribute("Type", "http://schemas.microsoft.com/office/2011/relationships/people")
                rel.setAttribute("Target", "people.xml")
                rdom.documentElement.appendChild(rel)
                rels_path.write_bytes(rdom.toxml(encoding="UTF-8"))
        ct_path = unpacked / "[Content_Types].xml"
        if ct_path.exists():
            cdom = defusedxml.minidom.parse(str(ct_path))
            has_ct = any(o.getAttribute("PartName") == "/word/people.xml" for o in cdom.getElementsByTagName("Override"))
            if not has_ct:
                ov = cdom.createElement("Override")
                ov.setAttribute("PartName", "/word/people.xml")
                ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml")
                cdom.documentElement.appendChild(ov)
                ct_path.write_bytes(cdom.toxml(encoding="UTF-8"))
    # ensure the author has a person entry
    pdom = defusedxml.minidom.parse(str(ppath))
    for p in pdom.getElementsByTagName("w15:person"):
        if p.getAttribute("w15:author") == author:
            return  # already present
    ns_w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
    person = pdom.createElementNS(ns_w15, "w15:person")
    person.setAttribute("w15:author", author)
    pi = pdom.createElementNS(ns_w15, "w15:presenceInfo")
    pi.setAttribute("w15:providerId", "None")
    pi.setAttribute("w15:userId", author)
    person.appendChild(pi)
    pdom.documentElement.appendChild(person)
    ppath.write_bytes(pdom.toxml(encoding="UTF-8"))


def ref_run(dom, cid):
    r, rpr, st = dom.createElement("w:r"), dom.createElement("w:rPr"), dom.createElement("w:rStyle")
    st.setAttribute("w:val", "CommentReference")
    rpr.appendChild(st)
    r.appendChild(rpr)
    ref = dom.createElement("w:commentReference")
    ref.setAttribute("w:id", str(cid))
    r.appendChild(ref)
    return r


def marker(dom, tag, cid):
    m = dom.createElement(tag)
    m.setAttribute("w:id", str(cid))
    return m


def op_reply(dom, unpacked, parent, text):
    """Reply inside an existing thread, anchored after the thread's last existing reply."""
    root, root_pid, members = thread(unpacked, parent)
    cid, pid, _ = docx_comment.add_comment(unpacked, text, author=AUTHOR, initials=INITIALS)
    ensure_people(unpacked, AUTHOR)
    ext_edit(unpacked, lambda x: x.getAttribute("w15:paraId") == pid and (x.setAttribute("w15:paraIdParent", root_pid) or True))
    refs = [e for e in dom.getElementsByTagName("w:commentReference") if int(e.getAttribute("w:id")) in members]
    if not refs:
        halt(f"thread {root}: no commentReference run in the body.")
    last_run = refs[-1].parentNode
    new_run = ref_run(dom, cid)
    last_run.parentNode.insertBefore(new_run, last_run.nextSibling)
    starts = [e for e in dom.getElementsByTagName("w:commentRangeStart") if int(e.getAttribute("w:id")) in members]
    ends = [e for e in dom.getElementsByTagName("w:commentRangeEnd") if int(e.getAttribute("w:id")) in members]
    rs, re_ = marker(dom, "w:commentRangeStart", cid), marker(dom, "w:commentRangeEnd", cid)
    (starts[-1].parentNode.insertBefore(rs, starts[-1].nextSibling) if starts
     else last_run.parentNode.insertBefore(rs, last_run))
    (ends[-1].parentNode.insertBefore(re_, ends[-1].nextSibling) if ends
     else new_run.parentNode.insertBefore(re_, new_run))
    return cid, root, int(refs[-1].getAttribute("w:id"))


def op_resolve(unpacked, parent):
    root, root_pid, members = thread(unpacked, parent)
    n = ext_edit(unpacked, lambda x: (x.getAttribute("w15:paraId") == root_pid or x.getAttribute("w15:paraIdParent") == root_pid)
                 and (x.setAttribute("w15:done", "1") or True))
    return root, n


# ---------------------------------------------------------------- accept (round-2 step and rule 6)
def rev_ok(e, keep, ids):
    i = e.getAttribute("w:id")
    return e.getAttribute("w:author") not in keep and (ids is None or (i.isdigit() and int(i) in ids))


def accept_dom(dom, keep, ids=None):
    """Accept tracked changes by authors not in keep (and, when ids is given, only those ids). Comment and
    bookmark markers inside deleted runs survive; a deleted paragraph mark merges into the next paragraph."""
    n, merge = 0, []
    for tag in PR_CHANGES:
        for e in list(dom.getElementsByTagName(tag)):
            if rev_ok(e, keep, ids):
                e.parentNode.removeChild(e)
                n += 1
    for e in list(dom.getElementsByTagName("w:ins")):
        if e.parentNode.tagName == "w:rPr" and rev_ok(e, keep, ids):
            e.parentNode.removeChild(e)
            n += 1
    for e in list(dom.getElementsByTagName("w:del")):
        if e.parentNode.tagName == "w:rPr" and rev_ok(e, keep, ids):
            p = e.parentNode.parentNode.parentNode
            e.parentNode.removeChild(e)
            n += 1
            if p.tagName == "w:p":
                merge.append(p)
    for e in list(dom.getElementsByTagName("w:del")):
        if e.parentNode is None or e.parentNode.tagName == "w:rPr" or not rev_ok(e, keep, ids):
            continue
        par = e.parentNode
        for c in kids(e):
            if c.tagName in MARKERS or c.getElementsByTagName("w:commentReference"):
                if c.tagName == "w:r":
                    for t in kids(c):
                        if t.tagName not in ("w:rPr", "w:commentReference"):
                            c.removeChild(t)
                par.insertBefore(c, e)
        par.removeChild(e)
        n += 1
    for e in list(dom.getElementsByTagName("w:ins")):
        if e.parentNode is None or e.parentNode.tagName == "w:rPr" or not rev_ok(e, keep, ids):
            continue
        unwrap(e)
        n += 1
    for p in merge:
        q = p.nextSibling
        while q is not None and q.nodeType != 1:
            q = q.nextSibling
        if q is None or q.tagName != "w:p":
            continue
        qp = kids(q, "w:pPr")
        anchor = qp[0].nextSibling if qp else q.firstChild
        for c in kids(p):
            if c.tagName != "w:pPr":
                q.insertBefore(c, anchor)
        p.parentNode.removeChild(p)
    return n


def revisions(dom):
    out = {}
    for tag in ("w:ins", "w:del"):
        for e in dom.getElementsByTagName(tag):
            k = (tag[2:], e.getAttribute("w:author"))
            out[k] = out.get(k, 0) + 1
    return out


def cmd_accept(src, dst, keep, ids, quiet=False):
    with zipfile.ZipFile(src) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
    dom = defusedxml.minidom.parseString(data["word/document.xml"])
    before = revisions(dom)
    n = accept_dom(dom, keep, ids)
    data["word/document.xml"] = dom.toxml(encoding="UTF-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            zout.writestr(i, data[i.filename])
    if quiet:
        return
    print(f"accepted {n} tracked changes (kept authors: {', '.join(sorted(keep))}"
          f"{'; ids ' + ','.join(map(str, sorted(ids))) if ids else ''})")
    print(f"revisions before: {before}\nrevisions after:  {revisions(dom)}\nwrote {dst}")


# ---------------------------------------------------------------- extract
def cmd_extract(docx):
    with tempfile.TemporaryDirectory() as td:
        unpacked = Path(td) / "u"
        unpack(docx, unpacked)
        dom, paras = load_doc(unpacked)
        indep = doc_para_count(unpacked)
        cm, ext = comment_maps(unpacked)
    existing = []
    print(f"# {docx.name}\n# {len(paras)} addressable blocks\n")
    for p in paras:
        existing += [(p.id,) + c for c in p.changes()]
        print(f"[{p.id}] {p.text().replace(chr(10), ' / ').replace(chr(9), ' | ')}")
    print()
    if existing:
        print("# EXISTING TRACKED CHANGES (read before editing; accept takes the ids):")
        for pid, kind, rid, author, txt in existing:
            print(f"#   {pid} {kind} id={rid} by {author}: {txt[:120]!r}")
        print()
    if cm:
        print("# COMMENTS (reply and resolve take the id):")
        for i, (author, txt, pid) in sorted(cm.items()):
            par, done = ext.get(pid, ("", ""))
            root = next((j for j, v in cm.items() if v[2] == par), "?") if par else None
            kind = f"reply to C{root}" if par else "thread root" + (", resolved" if done == "1" else ", open")
            print(f"#   C{i} by {author} ({kind}): {txt[:120]!r}")
        print()
    status = "OK" if indep == len(paras) else "MISMATCH"
    print(f"# STRUCTURE: <w:p> in body={indep} walked={len(paras)} [{status}]")
    if status == "MISMATCH":
        halt("the walker did not visit every paragraph in the body. A container type was skipped. "
             "Do not proceed until this is understood.")


# ---------------------------------------------------------------- apply
class RevIds:
    def __init__(self, dom):
        used = {int(e.getAttribute("w:id")) for tag in ("w:ins", "w:del", "w:commentRangeStart", "w:bookmarkStart")
                for e in dom.getElementsByTagName(tag) if e.getAttribute("w:id").isdigit()}
        self.next = (max(used) + 1) if used else 1000

    def take(self):
        self.next += 1
        return str(self.next - 1)


def make_run(dom, text, rpr, deleted):
    r = dom.createElement("w:r")
    if rpr is not None:
        r.appendChild(rpr.cloneNode(True))
    t = dom.createElement("w:delText" if deleted else "w:t")
    t.setAttributeNS(XML_NS, "xml:space", "preserve")
    t.appendChild(dom.createTextNode(text))
    r.appendChild(t)
    return r


def rpr_of(run):
    k = kids(run, "w:rPr")
    return k[0] if k else None


def wrap(dom, tag, ids, *runs):
    w = dom.createElement(tag)
    w.setAttribute("w:id", ids.take())
    w.setAttribute("w:author", AUTHOR)
    w.setAttribute("w:date", now_iso())
    for r in runs:
        w.appendChild(r)
    return w


def norm_to_raw(raw, nidx):
    i = n = 0
    while i < len(raw) and n < nidx:
        if raw[i] in WS:
            while i < len(raw) and raw[i] in WS:
                i += 1
        else:
            i += 1
        n += 1
    return i


def locate(para, needle, label):
    hay = para.text()
    hn, nn = norm(hay), norm(needle)
    if not nn.strip():
        halt(f"{para.id}: empty '{label}' string.")
    count = hn.count(nn)
    if count != 1:
        halt(f"{para.id}: '{label}' {'not found' if count == 0 else f'matches {count} times; widen it so it is unique'}."
             f"\n  looking for: {needle!r}\n  paragraph reads: {hay!r}")
    nidx = hn.find(nn)
    return norm_to_raw(hay, nidx), norm_to_raw(hay, nidx + len(nn))


def run_text(r):
    return "".join(t_text(c) if c.tagName == "w:t" else ("\t" if c.tagName == "w:tab" else "\n")
                   for c in kids(r, "w:t", "w:tab", "w:br", "w:cr"))


def decompose_runs(dom, para):
    """Split any run with mixed w:t/w:tab/w:br/w:cr children into one run per child."""
    for r in list(para.runs()):
        if ancestor(r, "w:del") is not None:
            continue
        content = kids(r, "w:t", "w:tab", "w:br", "w:cr")
        if len(content) <= 1:
            continue
        rpr = rpr_of(r)
        parent = r.parentNode
        for c in content:
            new_r = dom.createElement("w:r")
            if rpr is not None:
                new_r.appendChild(rpr.cloneNode(True))
            new_r.appendChild(c)
            parent.insertBefore(new_r, r)
        parent.removeChild(r)


def split_at(dom, para, offsets):
    decompose_runs(dom, para)
    for target in sorted(set(offsets)):
        pos = 0
        for r in para.runs():
            if ancestor(r, "w:del") is not None:
                continue
            rtext = run_text(r)
            if pos < target < pos + len(rtext):
                ts = kids(r, "w:t", "w:tab", "w:br", "w:cr")
                if len(ts) != 1 or ts[0].tagName != "w:t":
                    halt(f"{para.id}: cannot split a run containing non-text elements at offset {target}.")
                full, cut = t_text(ts[0]), target - pos
                r.parentNode.insertBefore(make_run(dom, full[:cut], rpr_of(r), False), r)
                r.parentNode.insertBefore(make_run(dom, full[cut:], rpr_of(r), False), r)
                r.parentNode.removeChild(r)
                break
            pos += len(rtext)


def runs_in_span(para, start, end):
    out, pos = [], 0
    for r in para.runs():
        if ancestor(r, "w:del") is not None:
            continue
        rtext = run_text(r)
        if pos >= start and pos + len(rtext) <= end and rtext:
            out.append(r)
        pos += len(rtext)
    return out


def to_deleted(dom, run):
    r = run.cloneNode(True)
    for t in list(r.getElementsByTagName("w:t")):
        d = dom.createElement("w:delText")
        d.setAttributeNS(XML_NS, "xml:space", "preserve")
        for c in list(t.childNodes):
            d.appendChild(c)
        t.parentNode.replaceChild(d, t)
    return r


def guard(para, runs, accepted):
    """Rule 6 and the cross-author refusal: a find range inside another author's still-tracked insertion
    accepts that insertion first; a range crossing from one author's insertion into another's stops."""
    owners, plain = {}, False
    for r in runs:
        e = ancestor(r, "w:ins")
        if e is None:
            plain = True
        else:
            owners.setdefault(e.getAttribute("w:author"), []).append(e)
    names = list(owners)
    if len(names) > 1:
        halt(f"{para.id}: find range crosses from {names[0]}'s insertion into {names[1]}'s. "
             "Split the edit at the boundary.")
    if names and names[0] != AUTHOR:
        for e in dict.fromkeys(owners[names[0]]):
            rid = e.getAttribute("w:id")
            unwrap(e)
            accepted.append((rid, names[0]))
            print(f"# {para.id}: accepted insertion id={rid} by {names[0]} before editing inside it (rule 6)")
    elif names and plain:
        halt(f"{para.id}: find range crosses from your own insertion into base text. Split the edit.")


def span(dom, para, find, label, accepted):
    start, end = locate(para, find, label)
    split_at(dom, para, [start, end])
    runs = runs_in_span(para, start, end)
    if not runs:
        halt(f"{para.id}: internal: no runs in span after split.")
    if accepted is not None:
        guard(para, runs, accepted)
    return runs


def op_replace_or_delete(dom, para, ids, find, with_, accepted):
    runs = span(dom, para, find, "find", accepted)
    rpr, anchor = rpr_of(runs[0]), runs[-1]
    parent, nxt = anchor.parentNode, anchor.nextSibling
    deleted = [to_deleted(dom, r) for r in runs]
    for r in runs:
        r.parentNode.removeChild(r)
    del_wrap = wrap(dom, "w:del", ids, *deleted)
    parent.insertBefore(del_wrap, nxt)
    if with_ is not None:
        parent.insertBefore(wrap(dom, "w:ins", ids, make_run(dom, with_, rpr, False)), del_wrap.nextSibling)


def op_insert_after(dom, para, ids, find, text, accepted):
    anchor = span(dom, para, find, "find", accepted)[-1]
    anchor.parentNode.insertBefore(wrap(dom, "w:ins", ids, make_run(dom, text, rpr_of(anchor), False)), anchor.nextSibling)


def para_mark(dom, para_node, ids, tag):
    ppr = kids(para_node, "w:pPr")
    if not ppr:
        ppr = [dom.createElement("w:pPr")]
        para_node.insertBefore(ppr[0], para_node.firstChild)
    rpr = kids(ppr[0], "w:rPr")
    if not rpr:
        rpr = [dom.createElement("w:rPr")]
        ppr[0].appendChild(rpr[0])
    rpr[0].insertBefore(wrap(dom, tag, ids), rpr[0].firstChild)


def op_insert_paragraph_after(dom, para, ids, text):
    newp = dom.createElement("w:p")
    src = kids(para.node, "w:pPr")
    newp.appendChild(src[0].cloneNode(True) if src else dom.createElement("w:pPr"))
    para_mark(dom, newp, ids, "w:ins")
    rs = para.runs()
    newp.appendChild(wrap(dom, "w:ins", ids, make_run(dom, text, rpr_of(rs[0]) if rs else None, False)))
    para.node.parentNode.insertBefore(newp, para.node.nextSibling)


def op_delete_paragraph(dom, para, ids, accepted):
    for r in para.runs():
        e = ancestor(r, "w:ins")
        if e is not None and e.getAttribute("w:author") != AUTHOR and e.parentNode is not None:
            accepted.append((e.getAttribute("w:id"), e.getAttribute("w:author")))
            print(f"# {para.id}: accepted insertion id={e.getAttribute('w:id')} by {e.getAttribute('w:author')} (rule 6)")
            unwrap(e)
    for r in para.runs():
        if ancestor(r, "w:del") is None:
            r.parentNode.replaceChild(wrap(dom, "w:del", ids, to_deleted(dom, r)), r)
    para_mark(dom, para.node, ids, "w:del")


def op_comment(dom, para, anchor, text, unpacked):
    runs = span(dom, para, anchor, "anchor", None)
    cid, _, _ = docx_comment.add_comment(unpacked, text, author=AUTHOR, initials=INITIALS)
    ensure_people(unpacked, AUTHOR)
    rs, re_ = marker(dom, "w:commentRangeStart", cid), marker(dom, "w:commentRangeEnd", cid)
    runs[0].parentNode.insertBefore(rs, runs[0])
    runs[-1].parentNode.insertBefore(re_, runs[-1].nextSibling)
    re_.parentNode.insertBefore(ref_run(dom, cid), re_.nextSibling)
    return cid


def render_marked(before, after):
    import difflib
    a, b, out = before.split(), after.split(), []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag == "equal":
            out.extend(a[i1:i2])
        else:
            if tag != "insert":
                out.append("~~" + " ".join(a[i1:i2]) + "~~")
            if tag != "delete":
                out.append("**" + " ".join(b[j1:j2]) + "**")
    return " ".join(out)


def check_plan(plan):
    edits = plan.get("edits")
    if not isinstance(edits, list) or not edits:
        halt("plan.json has no 'edits' list.")
    need = {"replace": ("id", "find", "with"), "insert_after": ("id", "find", "text"), "delete": ("id", "find"),
            "insert_paragraph_after": ("id", "text"), "delete_paragraph": ("id",), "comment": ("id", "anchor", "text"),
            "reply": ("parent", "text"), "resolve": ("parent",)}
    for i, e in enumerate(edits):
        if e.get("op") not in OPS:
            halt(f"edit #{i}: unknown op {e.get('op')!r}. Allowed: {sorted(OPS)}")
        for k in need[e["op"]]:
            if k not in e:
                halt(f"edit #{i}: {e['op']} needs {need[e['op']]}.")
        for k in e:
            if k not in KEYS:
                halt(f"edit #{i}: unexpected key {k!r}. Full-paragraph rewrites are not an op; use find/replace.")
    return edits


def cmd_apply(docx, plan_path, out, report):
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    edits = check_plan(plan)
    n_text = sum(1 for e in edits if e["op"] in TEXT_OPS)
    n_comment = sum(1 for e in edits if e["op"] == "comment" and not (e.get("question") or plan.get("questions")))
    if n_comment > n_text:
        halt(f"BUILD GATE: {n_comment} comments against {n_text} tracked changes. Comments are for business "
             'decisions; redline the rest, or mark the comments as questions ("question": true on the edit, '
             'or "questions": true on the plan).')
    with tempfile.TemporaryDirectory() as td:
        unpacked = Path(td) / "u"
        unpack(docx, unpacked)
        dom, paras = load_doc(unpacked)
        by_id = {p.id: p for p in paras}
        before = {p.id: p.text() for p in paras}
        ids, touched, accepted, notes = RevIds(dom), [], [], []
        for i, e in enumerate(edits):
            op = e["op"]
            if op in ("reply", "resolve"):
                if op == "reply":
                    cid, root, after = op_reply(dom, unpacked, int(e["parent"]), e["text"])
                    notes.append(f"## reply C{cid} in thread C{root}, placed after C{after}\nREPLY: {e['text']}")
                else:
                    root, n = op_resolve(unpacked, int(e["parent"]))
                    notes.append(f"## resolved thread C{root} ({n} entries marked done)")
                continue
            p = by_id.get(e["id"])
            if p is None:
                halt(f"edit #{i}: paragraph {e['id']} does not exist. Run extract again.")
            if op == "replace":
                op_replace_or_delete(dom, p, ids, e["find"], e["with"], accepted)
            elif op == "delete":
                op_replace_or_delete(dom, p, ids, e["find"], None, accepted)
            elif op == "insert_after":
                op_insert_after(dom, p, ids, e["find"], e["text"], accepted)
            elif op == "insert_paragraph_after":
                op_insert_paragraph_after(dom, p, ids, e["text"])
                notes.append(f"## new paragraph after {p.id}\nINSERTED: {e['text']}")
            elif op == "delete_paragraph":
                op_delete_paragraph(dom, p, ids, accepted)
            else:
                cid = op_comment(dom, p, e["anchor"], e["text"], unpacked)
                notes.append(f"## comment C{cid} on {p.id}{' (question)' if e.get('question') else ''}\nCOMMENT: {e['text']}")
            if op != "insert_paragraph_after" and op != "comment":
                touched.append(p.id)
        (unpacked / "word" / "document.xml").write_bytes(dom.toxml(encoding="UTF-8"))
        out.parent.mkdir(parents=True, exist_ok=True)
        rezip(unpacked, out)
        original = docx
        if accepted:  # validate against the input with the same insertions accepted (rule 6 leaves them untracked)
            original = Path(td) / "validation_original.docx"
            cmd_accept(docx, original, set(), {int(r) for r, _ in accepted if r.isdigit()}, quiet=True)
        after = {p.id: p.text() for p in paras}  # same nodes; text() is the accepted view
        lines = [f"# VERIFICATION REPORT: {out.name}", ""]
        for pid in dict.fromkeys(touched):
            lines += [f"## {pid}", f"BEFORE: {before[pid]}", f"AFTER:  {after.get(pid, '')}",
                      f"MARKED: {render_marked(before[pid], after.get(pid, ''))}", ""]
        for n in notes:
            lines += [n, ""]
        if accepted:
            lines += ["## accepted before editing (rule 6): " + ", ".join(f"id={r} by {a}" for r, a in accepted), ""]
        report.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        res = subprocess.run([sys.executable, str(DOCX_SKILL / "office" / "validate.py"), str(out),
                              "--original", str(original), "--author", AUTHOR], capture_output=True, text=True)
        print("\n# VALIDATE:\n" + (res.stdout.strip() or "(no output)"))
        if res.returncode != 0:
            print(res.stderr, file=sys.stderr)
            halt("validate.py failed. The redline is not safe to deliver.")
    print(f"\n# OK: {out}\n# report: {report}\n")
    cmd_check(out, plan_path)


# ---------------------------------------------------------------- post-build check
def cmd_check(out, plan_path):
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    fails = []
    try:
        import docx
        docx.Document(str(out))
        print("python-docx: loads")
    except Exception as x:
        fails.append(f"python-docx cannot open the output: {x}")
    if shutil.which("soffice"):
        from office.soffice import run_soffice
        with tempfile.TemporaryDirectory() as td:
            r = run_soffice(["--headless", "--convert-to", "pdf", "--outdir", td, str(out)],
                            capture_output=True, text=True, timeout=180)
            if any(Path(td).glob("*.pdf")):
                print("LibreOffice: loads")
            else:
                fails.append("LibreOffice headless could not open the output: " + (r.stderr or r.stdout)[-300:])
    else:
        print("LibreOffice: NOT RUN (soffice not found on this machine)")
    with tempfile.TemporaryDirectory() as td:
        u = Path(td) / "c"
        unpack(out, u, merge=False)
        dom, paras = load_doc(u)
        cm, ext = comment_maps(u)
        counts = {t: len(dom.getElementsByTagName(t)) for t in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")}
        print(f"comments: {len(cm)} in comments.xml, {counts['w:commentRangeStart']} commentRangeStart, "
              f"{counts['w:commentRangeEnd']} commentRangeEnd, {counts['w:commentReference']} commentReference")
        if len(set(counts.values()) | {len(cm)}) != 1:
            fails.append("comment marker counts do not match comments.xml")
        rs_ids = {int(e.getAttribute("w:id")) for e in dom.getElementsByTagName("w:commentRangeStart")}
        mine = {"ins": [], "del": []}
        for tag in ("w:ins", "w:del"):
            for e in dom.getElementsByTagName(tag):
                if e.getAttribute("w:author") == AUTHOR:
                    mine[tag[2:]].append((el_text(e) if tag == "w:ins" else el_text(e, True), e.parentNode.tagName == "w:rPr"))

        def has(kind, text, mark=False):
            return any(norm(text) in norm(t) and (m or not mark) for t, m in mine[kind]) if text else any(m for _, m in mine[kind])

        for i, e in enumerate(plan.get("edits", [])):
            op, ok = e["op"], True
            if op in ("replace", "delete"):
                ok = has("del", e["find"]) and (op == "delete" or has("ins", e["with"]))
            elif op == "insert_after":
                ok = has("ins", e["text"])
            elif op == "insert_paragraph_after":
                ok = has("ins", e["text"]) and has("ins", "", True)
            elif op == "delete_paragraph":
                ok = has("del", "", True)
            elif op == "comment":
                ok = any(norm(e["text"]) in norm(v[1]) and k in rs_ids for k, v in cm.items())
            elif op == "reply":
                ok = any(norm(e["text"]) in norm(v[1]) and k in rs_ids and ext.get(v[2], ("", ""))[0] for k, v in cm.items())
            elif op == "resolve":
                root, root_pid, members = thread(u, int(e["parent"]))
                ok = all(ext.get(cm[m][2], ("", "0"))[1] == "1" for m in members)
            if not ok:
                fails.append(f"edit #{i} ({op}{', ' + e['id'] if 'id' in e else ''}) has no matching tracked change or comment")
    print("CHECK: PASS" if not fails else "CHECK: FAIL\n  " + "\n  ".join(fails))
    if fails:
        sys.exit(1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("extract").add_argument("docx", type=Path)
    ac = sub.add_parser("accept")
    ac.add_argument("docx", type=Path)
    ac.add_argument("out", type=Path)
    ac.add_argument("--keep", default="", help='authors whose changes stay live, e.g. "Author A,Author B"')
    ac.add_argument("--ids", default="", help="accept only these revision ids, e.g. 12,15")
    apl = sub.add_parser("apply")
    apl.add_argument("docx", type=Path)
    apl.add_argument("plan", type=Path)
    apl.add_argument("--out", type=Path, required=True)
    apl.add_argument("--report", type=Path, default=None)
    ck = sub.add_parser("check")
    ck.add_argument("docx", type=Path)
    ck.add_argument("plan", type=Path)
    a = ap.parse_args()
    if a.cmd == "extract":
        cmd_extract(a.docx)
    elif a.cmd == "accept":
        keep = {AUTHOR} | {s.strip() for s in a.keep.split(",") if s.strip()}
        ids = {int(s) for s in a.ids.split(",") if s.strip().isdigit()} or None
        cmd_accept(a.docx, a.out, keep, ids)
    elif a.cmd == "apply":
        cmd_apply(a.docx, a.plan, a.out, a.report or a.out.with_suffix(".report.md"))
    else:
        cmd_check(a.docx, a.plan)


if __name__ == "__main__":
    main()
