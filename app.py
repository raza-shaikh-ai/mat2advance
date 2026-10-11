"""
Deep Metadata Remover
Supports: PDF (.pdf), Word (.docx)
Uses: pypdf, python-docx (no Poppler/system deps needed)
"""

import io
import os
import re
import zipfile
from typing import Dict, Any, Tuple, Optional

import streamlit as st

# ─────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Deep Metadata Remover",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────
# CSS – premium dark UI
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: linear-gradient(135deg, #0d0d14 0%, #12121f 50%, #0d1117 100%); }
#MainMenu, footer, header { visibility: hidden; }
.hero { text-align:center; padding:48px 24px 32px; }
.hero h1 { font-size:2.8rem; font-weight:700; background:linear-gradient(90deg,#a78bfa,#38bdf8,#34d399); -webkit-background-clip:text; -webkit-text-fill-color:transparent; margin-bottom:8px; }
.hero p { color:#94a3b8; font-size:1.05rem; max-width:560px; margin:0 auto; }
.section-label { font-size:0.75rem; font-weight:600; letter-spacing:0.12em; text-transform:uppercase; color:#64748b; margin-bottom:6px; }
.meta-table { width:100%; border-collapse:collapse; }
.meta-table th { text-align:left; font-size:0.72rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; color:#64748b; padding:6px 8px; border-bottom:1px solid rgba(255,255,255,0.06); }
.meta-table td { padding:7px 8px; font-size:0.84rem; color:#cbd5e1; border-bottom:1px solid rgba(255,255,255,0.04); word-break:break-all; }
.chip-ok { display:inline-block; background:rgba(52,211,153,0.15); border:1px solid rgba(52,211,153,0.35); color:#34d399; border-radius:6px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
.chip-err { display:inline-block; background:rgba(239,68,68,0.15); border:1px solid rgba(239,68,68,0.35); color:#f87171; border-radius:6px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
.divider { height:1px; background:linear-gradient(90deg,transparent,rgba(255,255,255,0.08),transparent); margin:20px 0; }
[data-testid="stFileUploader"] { background:rgba(255,255,255,0.03)!important; border:2px dashed rgba(167,139,250,0.3)!important; border-radius:12px!important; }
.stButton > button { background:linear-gradient(135deg,#7c3aed,#2563eb)!important; color:white!important; border:none!important; border-radius:10px!important; font-weight:600!important; padding:10px 28px!important; font-size:0.95rem!important; box-shadow:0 4px 20px rgba(124,58,237,0.35)!important; }
[data-testid="stDownloadButton"] button { background:linear-gradient(135deg,#059669,#0891b2)!important; color:white!important; border:none!important; border-radius:8px!important; font-weight:600!important; font-size:0.85rem!important; padding:6px 18px!important; box-shadow:0 3px 14px rgba(5,150,105,0.3)!important; }
[data-testid="stMetric"] { background:rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.07); border-radius:10px; padding:12px 16px; }
[data-testid="stMetricValue"] { color:#a78bfa!important; font-weight:700!important; }
[data-testid="stMetricLabel"] { color:#64748b!important; font-size:0.78rem!important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
# PDF helpers
# ─────────────────────────────────────────────────────────────────
def _read_pdf_meta(data: bytes) -> Dict[str, Any]:
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(data))
        meta: Dict[str, Any] = {}
        if reader.metadata:
            for k, v in reader.metadata.items():
                key = k.lstrip("/")
                if v:
                    meta[key] = str(v)
        try:
            xmp = reader.xmp_metadata
            if xmp:
                meta["XMP_raw"] = "<present>"
        except Exception:
            pass
        meta["_PageCount"] = str(len(reader.pages))
        return meta
    except Exception as e:
        return {"_error": str(e)}


def _clean_pdf(data: bytes) -> Tuple[Optional[bytes], Dict[str, str]]:
    report: Dict[str, str] = {}
    try:
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(io.BytesIO(data))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.add_metadata({
            "/Producer": "", "/Creator": "", "/Author": "",
            "/Title": "", "/Subject": "", "/Keywords": "",
            "/CreationDate": "", "/ModDate": "",
        })
        try:
            root = writer._root_object
            if "/Metadata" in root:
                del root["/Metadata"]
                report["XMP metadata"] = "removed"
        except Exception:
            pass
        try:
            root = writer._root_object
            if "/Names" in root:
                names = root["/Names"]
                if "/JavaScript" in names:
                    del names["/JavaScript"]
                    report["Embedded JavaScript"] = "removed"
        except Exception:
            pass
        annot_cleared = 0
        js_removed = 0
        for page in writer.pages:
            try:
                if "/Annots" in page:
                    for annot_ref in page["/Annots"]:
                        try:
                            annot = annot_ref.get_object()
                            for key in ["/T", "/TU", "/Contents", "/NM"]:
                                if key in annot:
                                    annot[key] = ""
                                    annot_cleared += 1
                            if "/A" in annot:
                                action = annot["/A"]
                                if "/S" in action and action["/S"] == "/JavaScript":
                                    del annot["/A"]
                                    js_removed += 1
                        except Exception:
                            pass
            except Exception:
                pass
        if annot_cleared:
            report["Annotation labels/authors"] = f"{annot_cleared} field(s) cleared"
        if js_removed:
            report["Page-level JavaScript"] = f"{js_removed} actions removed"
        report["Document info fields"] = "all cleared"
        buf = io.BytesIO()
        writer.write(buf)
        return buf.getvalue(), report
    except Exception as e:
        return None, {"_error": str(e)}


# ─────────────────────────────────────────────────────────────────
# DOCX helpers
# ─────────────────────────────────────────────────────────────────
import xml.etree.ElementTree as ET

DROP_PATTERNS = [
    re.compile(r'^customXml/'),
    re.compile(r'^docProps/custom\.xml$'),
    re.compile(r'printerSettings'),
    re.compile(r'^word/people\.xml$'),
]


def _read_docx_meta(data: bytes) -> Dict[str, Any]:
    meta: Dict[str, Any] = {}
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = z.namelist()
            if "docProps/core.xml" in names:
                tree = ET.fromstring(z.read("docProps/core.xml"))
                for child in tree:
                    tag = re.sub(r'\{[^}]+\}', '', child.tag)
                    if child.text and child.text.strip():
                        meta[f"core/{tag}"] = child.text
            if "docProps/app.xml" in names:
                tree = ET.fromstring(z.read("docProps/app.xml"))
                for child in tree:
                    tag = re.sub(r'\{[^}]+\}', '', child.tag)
                    if child.text and child.text.strip():
                        meta[f"app/{tag}"] = child.text
            if "docProps/custom.xml" in names:
                meta["Custom properties"] = "<present>"
            if "word/document.xml" in names:
                content = z.read("word/document.xml").decode("utf-8", errors="ignore")
                revisions = len(re.findall(r'<w:ins |<w:del ', content))
                if revisions:
                    meta["Tracked revisions"] = str(revisions)
                rsids = len(re.findall(r'w:rsid', content))
                if rsids:
                    meta["RSID fingerprints"] = str(rsids)
    except Exception as e:
        meta["_error"] = str(e)
    return meta


def _strip_xml(xml_bytes: bytes, filename: str) -> Tuple[bytes, int, int]:
    rsid_count = 0
    rev_count = 0
    try:
        ns_map = {}
        for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=["start-ns"]):
            ns_map[prefix] = uri
            ET.register_namespace(prefix, uri)
        tree = ET.parse(io.BytesIO(xml_bytes))
        parent_map = {c: p for p in tree.iter() for c in p}
        for elem in list(tree.iter()):
            keys_to_del = [k for k in elem.attrib if '}rsid' in k.lower()]
            for k in keys_to_del:
                del elem.attrib[k]
                rsid_count += 1
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for elem in list(tree.iterfind(f'.//{{{W}}}nsid')):
            if elem in parent_map:
                parent_map[elem].remove(elem)
        if "document.xml" in filename or "content.xml" in filename:
            for elem in list(tree.iterfind(f'.//{{{W}}}del')):
                if elem in parent_map:
                    parent_map[elem].remove(elem)
                    rev_count += 1
            for elem in list(tree.iterfind(f'.//{{{W}}}ins')):
                if elem in parent_map:
                    p = parent_map[elem]
                    idx = list(p).index(elem)
                    for child in list(elem):
                        p.insert(idx, child)
                        idx += 1
                    p.remove(elem)
                    rev_count += 1
        out = io.BytesIO()
        tree.write(out, xml_declaration=True, encoding="UTF-8")
        cleaned = out.getvalue()
        cleaned = re.sub(rb'mc:Ignorable="[^"]*"', b'', cleaned, count=1)
        return cleaned, rsid_count, rev_count
    except Exception:
        return xml_bytes, 0, 0


def _clean_docx(data: bytes) -> Tuple[Optional[bytes], Dict[str, str]]:
    report: Dict[str, str] = {}
    try:
        out_buf = io.BytesIO()
        total_rsid = 0
        total_rev = 0
        dropped = []
        CLEAN_CORE = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
            b' xmlns:dc="http://purl.org/dc/elements/1.1/"'
            b' xmlns:dcterms="http://purl.org/dc/terms/"'
            b' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            b'</cp:coreProperties>'
        )
        CLEAN_APP = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
            b'</Properties>'
        )
        with zipfile.ZipFile(io.BytesIO(data), 'r') as zin, \
             zipfile.ZipFile(out_buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if any(p.search(name) for p in DROP_PATTERNS):
                    dropped.append(name)
                    continue
                raw = zin.read(name)
                if name == "docProps/core.xml":
                    zout.writestr(item, CLEAN_CORE)
                    report["Core properties"] = "wiped (author, dates, description…)"
                elif name == "docProps/app.xml":
                    zout.writestr(item, CLEAN_APP)
                    report["App properties"] = "wiped (Company, AppVersion…)"
                elif name.endswith('.xml') and len(raw) > 0:
                    cleaned, rsid, rev = _strip_xml(raw, name)
                    total_rsid += rsid
                    total_rev += rev
                    zout.writestr(item, cleaned)
                else:
                    zout.writestr(item, raw)
        if total_rsid:
            report["RSID fingerprints"] = f"{total_rsid} removed"
        if total_rev:
            report["Tracked revisions"] = f"{total_rev} removed"
        if dropped:
            report["Dropped files"] = f"{len(dropped)} (customXml, printerSettings…)"
        return out_buf.getvalue(), report
    except Exception as e:
        return None, {"_error": str(e)}


# ─────────────────────────────────────────────────────────────────
# Dispatcher
# ─────────────────────────────────────────────────────────────────
def get_meta(filename: str, data: bytes) -> Dict[str, Any]:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf":
        return _read_pdf_meta(data)
    elif ext == ".docx":
        return _read_docx_meta(data)
    return {"_unsupported": ext}


def clean_file(filename: str, data: bytes) -> Tuple[Optional[bytes], Dict[str, str]]:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf":
        return _clean_pdf(data)
    elif ext == ".docx":
        return _clean_docx(data)
    return None, {"_error": f"Unsupported: {ext}"}


# ─────────────────────────────────────────────────────────────────
# UI helpers
# ─────────────────────────────────────────────────────────────────
def _meta_table(meta: Dict[str, Any]) -> str:
    if not meta:
        return '<p style="color:#64748b;font-size:0.85rem;">No metadata found.</p>'
    rows = "".join(
        f"<tr><td style='color:#94a3b8'>{k.lstrip('_')}</td>"
        f"<td style='color:#f87171'>{v}</td></tr>"
        for k, v in meta.items() if v
    )
    return f'<table class="meta-table"><thead><tr><th>Field</th><th>Value</th></tr></thead><tbody>{rows}</tbody></table>'


def _report_table(report: Dict[str, str]) -> str:
    rows = "".join(
        f"<tr><td style='color:#94a3b8'>{k}</td><td style='color:#34d399'>{v}</td></tr>"
        for k, v in report.items()
    )
    return f'<table class="meta-table"><thead><tr><th>Operation</th><th>Result</th></tr></thead><tbody>{rows}</tbody></table>'


def _fmt(n: int) -> str:
    if n < 1024: return f"{n} B"
    if n < 1024**2: return f"{n/1024:.1f} KB"
    return f"{n/1024**2:.2f} MB"


# ─────────────────────────────────────────────────────────────────
# Hero
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>🛡️ Deep Metadata Remover</h1>
  <p>Thoroughly strip hidden metadata from PDF &amp; Word documents —
     author info, revision history, RSID fingerprints,
     timestamps, embedded scripts &amp; more.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# Upload
# ─────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Drop files here",
    type=["pdf", "docx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown("""
    <div style="text-align:center;padding:40px 0;color:#475569;">
      <div style="font-size:3rem;margin-bottom:12px;">📂</div>
      <p style="margin:0;font-size:0.9rem;">
        Upload one or more <b style="color:#a78bfa">PDF</b> or
        <b style="color:#38bdf8">DOCX</b> files to get started
      </p>
      <p style="margin:6px 0 0;font-size:0.78rem;color:#334155;">
        All processing happens locally — nothing leaves your machine
      </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────────────────────────
# Summary bar
# ─────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
with c1: st.metric("Files queued", len(uploaded))
with c2: st.metric("Total size", _fmt(sum(f.size for f in uploaded)))
with c3: st.metric("PDF", sum(1 for f in uploaded if f.name.lower().endswith(".pdf")))
with c4: st.metric("DOCX", sum(1 for f in uploaded if f.name.lower().endswith(".docx")))

st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# Process
# ─────────────────────────────────────────────────────────────────
run = st.button("⚡ Strip All Metadata", use_container_width=True)

if run:
    results = []
    prog = st.progress(0, text="Processing…")
    for i, uf in enumerate(uploaded):
        raw = uf.read()
        before = get_meta(uf.name, raw)
        cleaned, report = clean_file(uf.name, raw)
        after = get_meta(uf.name, cleaned) if cleaned else {}
        results.append(dict(name=uf.name, raw=raw, cleaned=cleaned,
                            before=before, after=after, report=report))
        prog.progress((i + 1) / len(uploaded), text=f"Done: {uf.name}")
    prog.empty()
    st.session_state["results"] = results

# ─────────────────────────────────────────────────────────────────
# Results
# ─────────────────────────────────────────────────────────────────
if "results" in st.session_state:
    results = st.session_state["results"]
    st.markdown("### ✅ Results")

    for r in results:
        name = r["name"]
        cleaned = r["cleaned"]
        before = r["before"]
        after = r["after"]
        report = r["report"]
        ok = cleaned is not None and "_error" not in report
        ext = os.path.splitext(name)[1].upper().lstrip(".")
        badge_col = "#a78bfa" if ext == "PDF" else "#38bdf8"
        chip = '<span class="chip-ok">✓ Cleaned</span>' if ok else '<span class="chip-err">✗ Failed</span>'

        with st.expander(name, expanded=True):
            st.markdown(
                f'<span style="background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);'
                f'border-radius:6px;padding:2px 10px;font-size:0.8rem;color:{badge_col};font-weight:600;margin-right:8px;">'
                f'{ext}</span>{chip}',
                unsafe_allow_html=True,
            )
            if ok:
                m1, m2, m3 = st.columns(3)
                with m1: st.metric("Original", _fmt(len(r["raw"])))
                with m2: st.metric("Cleaned", _fmt(len(cleaned)))
                with m3: st.metric("Saved", _fmt(abs(len(r["raw"]) - len(cleaned))))

                st.markdown("<br>", unsafe_allow_html=True)
                b_col, a_col = st.columns(2)
                with b_col:
                    st.markdown("<div class='section-label'>📋 Before — metadata found</div>", unsafe_allow_html=True)
                    st.markdown(_meta_table(before), unsafe_allow_html=True)
                with a_col:
                    st.markdown("<div class='section-label'>🧹 After — remaining metadata</div>", unsafe_allow_html=True)
                    remaining = {k: v for k, v in after.items()
                                 if v and str(v) not in ("", "0") and not k.startswith("_Page")}
                    if remaining:
                        st.markdown(_meta_table(remaining), unsafe_allow_html=True)
                    else:
                        st.markdown(
                            '<div style="text-align:center;padding:24px;background:rgba(52,211,153,0.06);'
                            'border:1px solid rgba(52,211,153,0.2);border-radius:10px;">'
                            '<span style="font-size:1.8rem">✨</span><br>'
                            '<span style="color:#34d399;font-weight:600">Completely clean</span>'
                            '</div>',
                            unsafe_allow_html=True,
                        )

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("<div class='section-label'>🔧 Operations performed</div>", unsafe_allow_html=True)
                st.markdown(_report_table(report), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label=f"⬇️  Download  {name}",
                    data=cleaned,
                    file_name=name,
                    mime="application/octet-stream",
                    key=f"dl_{name}",
                )
            else:
                st.error(f"Error: {report.get('_error', 'Unknown error')}")

    # Bulk ZIP
    ok_results = [(r["name"], r["cleaned"]) for r in results if r["cleaned"]]
    if len(ok_results) > 1:
        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
        zb = io.BytesIO()
        with zipfile.ZipFile(zb, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for fname, fdata in ok_results:
                zf.writestr(fname, fdata)
        st.download_button(
            label=f"📦  Download all {len(ok_results)} cleaned files as ZIP",
            data=zb.getvalue(),
            file_name="cleaned_files.zip",
            mime="application/zip",
            use_container_width=True,
            key="dl_zip",
        )
