#!/usr/bin/env python3
"""NorthStar redline verification pair (ctx 3.4.0, rule 11).
Usage: python verify_redline.py OUT.docx ORIG.docx BASE.docx
ORIG is the counterparty document as received; BASE is ORIG after the accept step (the apply input).
REJECT: OUT with this turn's the author's changes rejected must read, paragraph for paragraph, as BASE.
REPORT: parts, revisions by author, original comments preserved, part counts, new comments, resolved and
open threads, reply order. Ends VERIFY: PASS or VERIFY: FAIL with the failing items.
"""
import os, re, sys, zipfile
from xml.dom import minidom

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
AUTHOR = os.environ.get("CTX_AUTHOR", "").strip()
if not AUTHOR:
    sys.exit("HALT: set CTX_AUTHOR (the tracked-change author name) before running.")


def text(n):
    return "".join(c.data for c in n.childNodes if c.nodeType == c.TEXT_NODE)


def rev_ids(d, author=None):
    return {e.getAttributeNS(W, "id") for t in ("ins", "del") for e in d.getElementsByTagNameNS(W, t)
            if author is None or e.getAttributeNS(W, "author") == author}


def unwrap(e):
    p = e.parentNode
    for c in list(e.childNodes):
        p.insertBefore(c, e)
    p.removeChild(e)


def sig(xml, reject=frozenset(), accept=frozenset()):
    """Per-paragraph (owner, text) segments: '=' kept, '+author' inserted, '-author' deleted. Revisions in
    reject are undone (a paragraph inserted under a rejected mark is dropped); revisions in accept are accepted."""
    d = minidom.parseString(xml)
    for e in list(d.getElementsByTagNameNS(W, "ins")):
        i, mark = e.getAttributeNS(W, "id"), e.parentNode.localName == "rPr"
        if i in reject:
            (e.parentNode.parentNode.parentNode.parentNode.removeChild(e.parentNode.parentNode.parentNode)
             if mark else e.parentNode.removeChild(e))
        elif i in accept:
            e.parentNode.removeChild(e) if mark else unwrap(e)
    for e in list(d.getElementsByTagNameNS(W, "del")):
        i, mark = e.getAttributeNS(W, "id"), e.parentNode.localName == "rPr"
        if mark and (i in reject or i in accept):
            e.parentNode.removeChild(e)
        elif i in reject:
            for t in list(e.getElementsByTagNameNS(W, "delText")):
                n = d.createElementNS(W, "w:t")
                for c in list(t.childNodes):
                    n.appendChild(c)
                t.parentNode.replaceChild(n, t)
            unwrap(e)
        elif i in accept:
            e.parentNode.removeChild(e)
    out = []
    for p in d.getElementsByTagNameNS(W, "p"):
        m = []
        for n in p.getElementsByTagNameNS("*", "*"):
            if n.localName not in ("t", "delText"):
                continue
            a, st = n.parentNode, "="
            while a is not None and a.nodeType == 1 and a.localName != "p":
                if a.localName in ("ins", "del"):
                    st = ("+" if a.localName == "ins" else "-") + a.getAttributeNS(W, "author")
                    break
                a = a.parentNode
            if m and m[-1][0] == st:
                m[-1] = (st, m[-1][1] + text(n))
            else:
                m.append((st, text(n)))
        out.append([s for s in m if s[1]])
    return out


def comments(z):
    r = {}
    if "word/comments.xml" not in z.namelist():
        return r
    for x in minidom.parseString(z.read("word/comments.xml")).getElementsByTagNameNS(W, "comment"):
        ps = x.getElementsByTagNameNS(W, "p")
        txt = " / ".join("".join(text(t) for t in p.getElementsByTagNameNS(W, "t")) for p in ps)
        r[int(x.getAttributeNS(W, "id"))] = (x.getAttributeNS(W, "author"), txt, ps[-1].getAttributeNS(W14, "paraId"))
    return r


def ext(z):
    if "word/commentsExtended.xml" not in z.namelist():
        return {}
    return {x.getAttributeNS(W15, "paraId"): (x.getAttributeNS(W15, "paraIdParent"), x.getAttributeNS(W15, "done"))
            for x in minidom.parseString(z.read("word/commentsExtended.xml")).getElementsByTagNameNS(W15, "commentEx")}


def main(out, orig, base):
    zo, zg, zb = zipfile.ZipFile(out), zipfile.ZipFile(orig), zipfile.ZipFile(base)
    fails = []
    print("=====", out)
    ob, bb = zo.read("word/document.xml"), zb.read("word/document.xml")
    do, db = minidom.parseString(ob), minidom.parseString(bb)
    new_ids = rev_ids(do, AUTHOR) - rev_ids(db, AUTHOR)
    gone = rev_ids(db) - rev_ids(do)  # base revisions the build accepted before editing inside them (rule 6)
    a, b = sig(ob, reject=new_ids), sig(bb, accept=gone)
    bad = [i for i, (x, y) in enumerate(zip(a, b)) if x != y]
    if len(a) != len(b):
        bad.append(f"paragraph count {len(a)} vs {len(b)}")
    print(f"REJECT: this turn's {len(new_ids)} {AUTHOR} revisions rejected ({len(gone)} base revisions accepted "
          f"under rule 6); paragraphs {len(a)} vs base {len(b)}; mismatched: {bad[:20]}")
    for i in [i for i in bad if isinstance(i, int)][:5]:
        print("  OUT ", a[i], "\n  BASE", b[i])
    if bad:
        fails.append("REJECT: output with the author's changes rejected does not read as the base")
    exp = {"word/document.xml", "word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml",
           "word/commentsExtensible.xml", "word/people.xml", "[Content_Types].xml", "word/_rels/document.xml.rels"}
    changed = [n for n in zg.namelist() if n not in exp and (n not in zo.namelist() or zg.read(n) != zo.read(n))]
    missing = [n for n in zg.namelist() if n not in zo.namelist()]
    print("1 other parts changed vs counterparty original:", changed or "none", "| missing:", missing or "none")
    if changed or missing:
        fails.append("1 package parts changed or missing")
    d = minidom.parseString(ob)
    revs = {}
    for t in ("ins", "del", "rPrChange", "pPrChange", "moveFrom", "moveTo"):
        for e in d.getElementsByTagNameNS(W, t):
            k = (t, e.getAttributeNS(W, "author"))
            revs[k] = revs.get(k, 0) + 1
    print("2 revisions in output by type/author:", revs)
    co, cg = comments(zo), comments(zg)
    lost = [i for i in cg if i not in co or co[i][:2] != cg[i][:2]]
    print("3 original comments:", len(cg), "| lost or altered:", lost or "none")
    if lost:
        fails.append("3 original comments lost or altered")
    eo, eg = ext(zo), ext(zg)
    pid2id = {v[2]: k for k, v in co.items()}
    n_ids = len(re.findall(rb"<w16cid:commentId ", zo.read("word/commentsIds.xml"))) if "word/commentsIds.xml" in zo.namelist() else 0
    n_ex = len(re.findall(rb"<w16cex:commentExtensible ", zo.read("word/commentsExtensible.xml"))) if "word/commentsExtensible.xml" in zo.namelist() else 0
    print("4 part counts comments/ext/ids/extensible:", len(co), len(eo), n_ids, n_ex)
    if co and len({len(co), len(eo), n_ids, n_ex}) != 1:
        fails.append("4 comment part counts differ")
    doc = ob.decode("utf-8")
    for i, (au, t, p) in sorted(co.items()):
        if i in cg:
            continue
        par = eo.get(p, ("", ""))[0]
        dashes = bool(re.search("[\u2013\u2014]|--", t))
        print(f"5 NEW C{i} by {au} | reply to C{pid2id.get(par, '-')} | ref {f'<w:commentReference w:id={chr(34)}{i}{chr(34)}/>' in doc} "
              f"| range {f'<w:commentRangeStart w:id={chr(34)}{i}{chr(34)}/>' in doc} | dashes {dashes} | {t[:80]}")
        if dashes:
            fails.append(f"5 comment C{i} contains a dash")
        if au != AUTHOR:
            fails.append(f"5 comment C{i} is not by {AUTHOR}")
    flipped = sorted({pid2id.get(eo[p][0] or p, "?") for p in eo if p in eg and eg[p][1] != eo[p][1]})
    print("6 threads whose resolved state changed (root ids):", flipped)
    print("7 threads still open (root ids):", sorted(pid2id[p] for p, v in eo.items() if not v[0] and v[1] != "1" and p in pid2id))
    for i, (au, t, p) in co.items():
        par = eo.get(p, ("", ""))[0]
        if i in cg or not par:
            continue
        mem = [pid2id[q] for q, v in eo.items() if (v[0] == par or q == par) and q in pid2id]
        pos = {m: doc.find(f'<w:commentReference w:id="{m}"/>') for m in mem}
        last = max(pos, key=pos.get) == i
        print(f"8 reply C{i} displays last in its thread: {last}")
        if not last:
            fails.append(f"8 reply C{i} is not last in its thread")
    print("VERIFY: PASS" if not fails else "VERIFY: FAIL\n  " + "\n  ".join(fails))
    sys.exit(1 if fails else 0)


def cli():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(2)
    main(*sys.argv[1:4])


if __name__ == "__main__":
    cli()
