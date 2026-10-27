"""DOCX to Markdown API with LibreOffice-based page markers.

Requires: FastAPI, uvicorn, Pandoc, LibreOffice, and Python UNO (uno).
Run with: uvicorn docx_to_md_api:app --host 0.0.0.0 --port 8000
Endpoint: POST /convert (multipart form field: file)
"""
import copy
import difflib
import os
import re
import shutil
import socket
import subprocess
import tempfile
import time
import zipfile
from xml.etree import ElementTree as ET

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("xml", XML_NS)

def q(tag):
    return "{%s}%s" % (W, tag)

MARK = "ZZPAGEMARK%dZZ"
MARK_RE = re.compile(r"ZZPAGEMARK(\d+)ZZ")

app = FastAPI(title="DOCX to Markdown with Page Numbers", version="1.1.0")
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", "/app/output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

def _free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class LibreOffice:
    """Headless soffice with a private profile and a unique port (safe to run in parallel)."""

    def __init__(self):
        self.port = _free_port()
        self.profile = tempfile.mkdtemp(prefix="lo_profile_")
        self.proc = None

    def __enter__(self):
        import uno  # noqa: F401  (python3-uno)
        exe = shutil.which("soffice") or shutil.which("libreoffice")
        if not exe:
            raise RuntimeError("LibreOffice (soffice) not found")
        self.proc = subprocess.Popen(
            [exe, "--headless", "--invisible", "--norestore", "--nologo", "--nodefault",
             "-env:UserInstallation=file://" + self.profile,
             "--accept=socket,host=127.0.0.1,port=%d;urp;" % self.port],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        import uno
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        last = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(
                    "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % self.port)
                self.desktop = ctx.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", ctx)
                return self
            except Exception as e:  # not up yet
                last = e
                time.sleep(0.5)
        raise RuntimeError("Could not connect to LibreOffice: %r" % last)

    def __exit__(self, *a):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        if self.proc:
            try:
                self.proc.wait(timeout=10)
            except Exception:
                self.proc.kill()
        shutil.rmtree(self.profile, ignore_errors=True)

    def load(self, path):
        import uno
        from com.sun.star.beans import PropertyValue

        def pv(n, v):
            p = PropertyValue()
            p.Name, p.Value = n, v
            return p
        # MacroExecutionMode 0 = never execute macros (safe for untrusted files)
        return self.desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0,
            (pv("Hidden", True), pv("ReadOnly", True), pv("MacroExecutionMode", 0)))


def _cell_sort_key(name):
    m = re.match(r"([A-Z]+)(\d+)$", name)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


def lo_paragraphs(doc):
    """Flat, document-order list of paragraph records from LibreOffice.

    Each record: dict(obj, text, in_table, row_id)  (row_id identifies a table row)
    """
    out = []
    counter = [0]

    def walk_text(text, table_ctx):
        enum = text.createEnumeration()
        while enum.hasMoreElements():
            el = enum.nextElement()
            if el.supportsService("com.sun.star.text.TextTable"):
                counter[0] += 1
                tid = counter[0]
                names = sorted(el.getCellNames(), key=_cell_sort_key)
                for n in names:
                    cell = el.getCellByName(n)
                    walk_text(cell, (tid, _cell_sort_key(n)[0]))
            elif el.supportsService("com.sun.star.text.Paragraph"):
                out.append(dict(obj=el, text=el.getString(), table=table_ctx))
    walk_text(doc.getText(), None)
    return out


def lo_layout(path):
    """Return (lo_paras, info) where for each LO paragraph info has start page and
    a list of (offset, page) for every page break that falls inside the paragraph."""
    with LibreOffice() as lo:
        doc = lo.load(path)
        try:
            ctrl = doc.getCurrentController()
            vc = ctrl.getViewCursor()
            paras = lo_paragraphs(doc)

            def page_at(par, offset):
                vc.gotoRange(par["obj"].getStart(), False)
                if offset:
                    vc.goRight(offset, False)
                return vc.getPage()

            for par in paras:
                n = len(par["text"])
                vc.gotoRange(par["obj"].getStart(), False)
                par["page"] = vc.getPage()
                vc.gotoRange(par["obj"].getEnd(), False)
                end_page = vc.getPage()
                par["splits"] = []
                if n > 1 and end_page > par["page"]:
                    for target in range(par["page"] + 1, end_page + 1):
                        lo_, hi = 1, n
                        while lo_ < hi:
                            mid = (lo_ + hi) // 2
                            if page_at(par, mid) >= target:
                                hi = mid
                            else:
                                lo_ = mid + 1
                        par["splits"].append((lo_, target))
                par.pop("obj")
            return paras
        finally:
            doc.close(True)


# --------------------------------------------------------------------------- #
# Step 2: align LO paragraphs with <w:p> in document.xml and inject markers
# --------------------------------------------------------------------------- #

def norm(s):
    return re.sub(r"\s+", " ", s or "").strip().lower()


def p_text(p):
    out = []
    for el in p.iter():
        if el.tag == q("t"):
            out.append(el.text or "")
        elif el.tag == q("tab"):
            out.append("\t")
    return "".join(out)


def xml_paragraphs(root):
    """Document-order <w:p> list with table context (tid, row_index)."""
    out = []
    counter = [0]

    def walk(node, ctx):
        for ch in list(node):
            if ch.tag == q("p"):
                out.append(dict(el=ch, text=p_text(ch), table=ctx, parent=node))
            elif ch.tag == q("tbl"):
                counter[0] += 1
                tid = counter[0]
                ri = 0
                for tr in ch.findall(q("tr")):
                    ri += 1
                    for tc in tr.findall(q("tc")):
                        walk(tc, (tid, ri))
            elif ch.tag in (q("sdt"),):
                c = ch.find(q("sdtContent"))
                if c is not None:
                    walk(c, ctx)
            elif ch.tag in (q("ins"), q("smartTag"), q("customXml")):
                walk(ch, ctx)
    body = root.find(q("body"))
    walk(body, None)
    return out


def align(lo_paras, x_paras):
    """Map xml paragraph index -> lo paragraph (or None) using text alignment."""
    a = [norm(p["text"]) for p in lo_paras]
    b = [norm(p["text"]) for p in x_paras]
    mapping = {}
    if len(a) == len(b):
        ok = sum(1 for i in range(len(a)) if a[i] == b[i])
        if ok >= 0.9 * max(1, len(a)):
            return {i: lo_paras[i] for i in range(len(b))}
    sm = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                mapping[j1 + k] = lo_paras[i1 + k]
        elif tag == "replace" and (i2 - i1) == (j2 - j1):
            for k in range(i2 - i1):
                mapping[j1 + k] = lo_paras[i1 + k]
    return mapping


def marker_run(n):
    r = ET.Element(q("r"))
    t = ET.SubElement(r, q("t"))
    t.text = MARK % n
    t.set("{%s}space" % XML_NS, "preserve")
    return r


def marker_paragraph(n):
    p = ET.Element(q("p"))
    p.append(marker_run(n))
    return p


def insert_marker_at_offset(p, offset, n):
    """Insert a marker run into <w:p> at `offset` characters (w:t/w:tab counted)."""
    pos = 0
    runs = [r for r in p.iter(q("r"))]
    parent_of = {c: par for par in p.iter() for c in par}
    for r in runs:
        text_len = 0
        for el in r:
            if el.tag == q("t"):
                text_len += len(el.text or "")
            elif el.tag in (q("tab"), q("cr"), q("br")):
                text_len += 1
        if text_len == 0:
            continue
        if pos + text_len <= offset:
            pos += text_len
            continue
        inner = offset - pos
        par = parent_of[r]
        idx = list(par).index(r)
        if inner <= 0:
            par.insert(idx, marker_run(n))
            return True
        # split this run at `inner` (only simple runs with a single w:t are split)
        ts = [e for e in r if e.tag == q("t")]
        if len(ts) == 1 and all(e.tag in (q("rPr"), q("t")) for e in r):
            full = ts[0].text or ""
            first, second = full[:inner], full[inner:]
            ts[0].text = first
            ts[0].set("{%s}space" % XML_NS, "preserve")
            r2 = copy.deepcopy(r)
            t2 = [e for e in r2 if e.tag == q("t")][0]
            t2.text = second
            t2.set("{%s}space" % XML_NS, "preserve")
            par.insert(idx + 1, marker_run(n))
            par.insert(idx + 2, r2)
        else:
            par.insert(idx + 1, marker_run(n))
        return True
    return False


def inject(root, lo_paras):
    """Insert page markers. Returns number of markers inserted."""
    x_paras = xml_paragraphs(root)
    mapping = align(lo_paras, x_paras)
    inserted = 0
    current_page = 0
    seen_table_rows = set()
    for i, xp in enumerate(x_paras):
        lp = mapping.get(i)
        if lp is None:
            continue
        page = lp["page"]
        el, parent = xp["el"], xp["parent"]
        if page > current_page:
            if xp["table"] is None:
                parent.insert(list(parent).index(el), marker_paragraph(page))
            else:
                # inside a table: put marker in front of this paragraph's first run
                el.insert(1 if el.find(q("pPr")) is not None else 0, marker_run(page))
            inserted += 1
            current_page = page
        for off, pg in lp["splits"]:
            if pg > current_page:
                if insert_marker_at_offset(el, off, pg):
                    inserted += 1
                    current_page = pg
    return inserted


def build_marked_docx(src, lo_paras, dst):
    with zipfile.ZipFile(src) as zin:
        xml = zin.read("word/document.xml")
        # parse keeping namespace prefixes stable
        for pre, uri in re.findall(rb'xmlns:(\w+)="([^"]+)"', xml):
            ET.register_namespace(pre.decode(), uri.decode())
        root = ET.fromstring(xml)
        n = inject(root, lo_paras)
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
    return n


# --------------------------------------------------------------------------- #
# Step 3/4: Pandoc + marker rewrite
# --------------------------------------------------------------------------- #

def run_pandoc(docx, fmt="gfm"):
    r = subprocess.run(["pandoc", "-f", "docx", "-t", fmt, "--wrap=none", docx],
                       capture_output=True, text=True)
    if r.returncode:
        raise RuntimeError("Pandoc failed: " + r.stderr)
    return r.stdout


def rewrite_markers(md):
    md = re.sub(r"^\s*(?:\\?\*\*)?ZZPAGEMARK(\d+)ZZ(?:\*\*)?\s*$",
                lambda m: "<!-- page %s -->" % m.group(1), md, flags=re.M)
    md = MARK_RE.sub(lambda m: "<!-- page %s -->" % m.group(1), md)
    return md


def split_pages(md):
    parts = re.split(r"<!-- page (\d+) -->", md)
    pages = []
    if parts[0].strip():
        pages.append((1, parts[0].strip()))
    for i in range(1, len(parts), 2):
        pages.append((int(parts[i]), parts[i + 1].strip()))
    return pages


def convert(docx_path, fmt="gfm"):
    """Returns (markdown_with_page_comments, [(page_no, markdown), ...])."""
    paras = lo_layout(docx_path)
    tmp = tempfile.mkdtemp(prefix="d2mp_")
    try:
        marked = os.path.join(tmp, "marked.docx")
        build_marked_docx(docx_path, paras, marked)
        md = rewrite_markers(run_pandoc(marked, fmt))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return md, split_pages(md)




@app.post("/convert")
async def convert_docx(file: UploadFile = File(...)):
    """Upload a .docx and receive Markdown with page comments and page-wise Markdown."""
    filename = file.filename or "document.docx"
    if not filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Please upload a .docx file.")

    temp_dir = tempfile.mkdtemp(prefix="docx_api_")
    input_path = os.path.join(temp_dir, "input.docx")
    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="The uploaded file is empty.")

        with open(input_path, "wb") as f:
            f.write(content)

        markdown, pages = convert(input_path)

        # Save the Markdown beside the configured output directory. Use only the
        # uploaded file's base name to prevent path traversal.
        safe_stem = os.path.splitext(os.path.basename(filename))[0] or "document"
        output_filename = safe_stem + ".md"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        with open(output_path, "w", encoding="utf-8", newline="\n") as md_file:
            md_file.write(markdown)

        return JSONResponse({
            "filename": filename,
            "markdown_file": output_filename,
            "download_url": f"/download/{output_filename}",
            "page_count": len(pages),
            "markdown": markdown,
            "pages": [{"page_number": number, "markdown": body} for number, body in pages],
        })
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"DOCX conversion failed: {exc}") from exc
    finally:
        await file.close()
        shutil.rmtree(temp_dir, ignore_errors=True)


@app.get("/download/{filename}")
def download_markdown(filename: str):
    """Download a Markdown file generated by POST /convert."""
    # Only permit simple .md filenames, never arbitrary paths.
    if not filename.endswith(".md") or os.path.basename(filename) != filename:
        raise HTTPException(status_code=400, detail="Invalid Markdown filename.")
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Markdown file not found.")
    return FileResponse(path, media_type="text/markdown; charset=utf-8", filename=filename)
