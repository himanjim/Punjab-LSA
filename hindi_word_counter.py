#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hindi_word_counter.py
---------------------
Recursively scans a folder (including nested .zip / .7z / .rar / .tar(.gz/.bz2/.xz)
/ .gz / .bz2 / .xz archives, and .iso/.cab/etc. via 7-Zip) for PDF and Word
documents, counts Hindi (Devanagari-script) words in each, OCRs PDF pages when
the text layer is missing / unusable, and writes an Excel report.

Examples (Windows PowerShell, inside the venv):
    python hindi_word_counter.py "D:\\Input Folder"
    python hindi_word_counter.py "D:\\Input Folder" -o D:\\report.xlsx --workers 4
    python hindi_word_counter.py "D:\\Input Folder" --ocr always --dpi 400
    python hindi_word_counter.py "D:\\Input Folder" --dump-text D:\\extracted_text

Run with -h for all options.  Requires Python 3.10+.
"""
from __future__ import annotations

import argparse
import bz2
import gzip
import hashlib
import lzma
import multiprocessing
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import time
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
import zlib
from concurrent.futures import ProcessPoolExecutor, as_completed
from concurrent.futures.process import BrokenProcessPool
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    from tqdm import tqdm
except ImportError:  # progress bar is optional
    def tqdm(it=None, **_kw):
        return it

IS_WIN = os.name == "nt"

# --------------------------------------------------------------------------- #
# File-type classification
# --------------------------------------------------------------------------- #
PDF_EXTS = {".pdf"}
OOXML_EXTS = {".docx", ".docm", ".dotx", ".dotm"}
LEGACY_WORD_EXTS = {".doc", ".dot", ".rtf"}            # converted via LibreOffice
TEXT_EXTS = {".txt"}                                  # plain-text documents
DOC_EXTS = PDF_EXTS | OOXML_EXTS | LEGACY_WORD_EXTS | TEXT_EXTS

TAR_SUFFIXES = (".tar", ".tar.gz", ".tgz", ".tar.bz2", ".tbz2", ".tbz", ".tar.xz", ".txz")
SINGLE_COMPRESSED = {".gz": gzip.open, ".bz2": bz2.open, ".xz": lzma.open}
CLI_ONLY_SUFFIXES = (".iso", ".cab", ".arj", ".lzh", ".lha", ".wim", ".001")   # need 7-Zip
SKIP_DIR_NAMES = {"__MACOSX", "$RECYCLE.BIN", "System Volume Information"}


def classify(name: str):
    """Return ('doc', ext) | ('archive', kind) | ('skipvol', None) | (None, None)."""
    low = name.lower()
    base = posixpath.basename(low.replace("\\", "/"))
    if base.startswith("~$") or base.startswith("._"):   # Word lock files / macOS forks
        return None, None
    ext = os.path.splitext(base)[1]
    if ext in DOC_EXTS:
        return "doc", ext
    m = re.search(r"\.part(\d+)\.rar$", base)
    if m and int(m.group(1)) > 1:                       # later volumes of a multi-part RAR
        return "skipvol", None
    if base.endswith(TAR_SUFFIXES):
        return "archive", "tar"
    if ext == ".zip":
        return "archive", "zip"
    if ext == ".7z":
        return "archive", "7z"
    if ext == ".rar":
        return "archive", "rar"
    if ext in SINGLE_COMPRESSED:
        return "archive", "single"
    if base.endswith(CLI_ONLY_SUFFIXES):
        return "archive", "cli"
    return None, None


# --------------------------------------------------------------------------- #
# Word counting
# --------------------------------------------------------------------------- #
# Devanagari letters + vowel signs + nasal marks + virama + nukta, Devanagari
# Extended, Vedic Extensions, ZWJ/ZWNJ.  Excludes danda (।॥), Devanagari digits
# (०-९) and the abbreviation sign (॰), so those never count as words.
_DEV_CHARS = "\u0900-\u0963\u0971-\u097F\uA8E0-\uA8FF\u1CD0-\u1CFF\u200C\u200D"
DEV_RUN_RE = re.compile(f"[{_DEV_CHARS}]+")
DEV_LETTER_RE = re.compile("[\u0904-\u0939\u0950\u0958-\u0961\u0972-\u097F\uA8F2-\uA8F7]")
APOSTROPHE_RE = re.compile(r"(?<=[^\W\d_])['\u2019](?=[^\W\d_])")


def _is_latin(c: str) -> bool:
    return c < "\u0250" or "\u1e00" <= c <= "\u1eff"


def count_words(text: str):
    """Return (hindi_words, latin_script_words, other_script_words)."""
    hindi = sum(1 for m in DEV_RUN_RE.finditer(text) if DEV_LETTER_RE.search(m.group()))
    rest = APOSTROPHE_RE.sub("", DEV_RUN_RE.sub(" ", text)) + " "
    latin = other = 0
    in_word, word_latin, has_letter = False, True, False
    for c in rest:
        cat = unicodedata.category(c)
        if cat[0] in "LM" or c in "\u200c\u200d":       # letters + combining marks (matras etc.)
            in_word = True
            if cat[0] == "L":
                has_letter = True
                word_latin = word_latin and _is_latin(c)
        elif in_word:
            if has_letter:
                if word_latin:
                    latin += 1
                else:
                    other += 1
            in_word, word_latin, has_letter = False, True, False
    return hindi, latin, other


# Legacy (non-Unicode) Hindi fonts: text looks like Latin garbage ("Hkkjr") in the
# text layer, so it must be OCRed.
LEGACY_FONT_RE = re.compile(
    r"kruti|devlys|chanakya|shusha|walkman|akruti|aps[-_ ]?dv|dv[-_ ]?tt|"
    r"shree[-_ ]?dev|ism[-_ ]?dv|shivaji\d|priyanka",
    re.I,
)
# Unicode Devanagari-capable fonts: if the page uses one but yields zero Hindi
# words, the PDF's ToUnicode map is probably broken -> OCR.
UNICODE_DEVA_FONT_RE = re.compile(
    r"mangal|devanagari|aparajita|kokila|utsaah|lohit|kalimati|nirmala|sanskrit|hindi|kiran|arialunicode",
    re.I,
)


# --------------------------------------------------------------------------- #
# Options / tool discovery
# --------------------------------------------------------------------------- #
@dataclass
class Options:
    ocr: str = "auto"                 # auto | always | never
    ocr_lang: str = "hin+eng"
    dpi: int = 300
    psm: int = 3
    ocr_timeout: int = 300
    tesseract_cmd: str | None = None
    ocr_available: bool = False
    soffice: str | None = None
    docx_pages: str = "metadata"      # metadata | libreoffice
    min_text_chars: int = 25
    image_cover: float = 0.5
    dump_dir: str | None = None


def find_tool(user_value, names, win_rel_paths):
    if user_value:
        return user_value if os.path.isfile(user_value) else shutil.which(user_value)
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    if IS_WIN:
        roots = [
            os.environ.get("ProgramFiles", r"C:\Program Files"),
            os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs"),
        ]
        for r in roots:
            for rel in win_rel_paths:
                p = os.path.join(r, rel)
                if os.path.isfile(p):
                    return p
    return None


def tesseract_langs(cmd):
    try:
        p = subprocess.run([cmd, "--list-langs"], capture_output=True, text=True,
                           errors="replace", timeout=60)
    except Exception:
        return set()
    lines = (p.stdout + "\n" + p.stderr).splitlines()
    return {ln.strip() for ln in lines if ln.strip() and not ln.lower().startswith("list of")}


# --------------------------------------------------------------------------- #
# Windows long-path helpers
# --------------------------------------------------------------------------- #
def to_long(p: str) -> str:
    """Extended-length path on Windows so deep folders (>260 chars) still work."""
    if not IS_WIN:
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def to_plain(p: str) -> str:
    if IS_WIN and p.startswith("\\\\?\\UNC\\"):
        return "\\\\" + p[8:]
    if IS_WIN and p.startswith("\\\\?\\"):
        return p[4:]
    return p


def for_native_lib(p: str) -> str:
    """Drop the \\\\?\\ prefix when the path is short enough for any library."""
    plain = to_plain(p)
    return plain if len(plain) < 250 else p


def join_logical(prefix: str, rel: str) -> str:
    return f"{prefix} > {rel}" if prefix else rel


# --------------------------------------------------------------------------- #
# Phase 1: discovery (walk folders, extract archives recursively)
# --------------------------------------------------------------------------- #
@dataclass
class DocItem:
    sno: int
    name: str
    ext: str
    found_in: str
    logical: str
    real: str
    size: int


class Discoverer:
    def __init__(self, work_dir: str, max_depth: int, sevenzip: str | None):
        self.work = work_dir
        self.max_depth = max_depth
        self.sevenzip = sevenzip
        self.docs: list[DocItem] = []
        self.skipped: list[tuple[str, str]] = []
        self.archives_opened = 0
        self._n = 0

    # -- helpers ----------------------------------------------------------- #
    def skip(self, what: str, why: str):
        self.skipped.append((what, why))

    def _new_dir(self) -> str:
        self._n += 1
        d = os.path.join(self.work, f"x{self._n}")
        os.makedirs(d, exist_ok=True)
        return d

    def _flat_name(self, member_name: str) -> str:
        """Short, ASCII-safe on-disk name (avoids MAX_PATH & illegal chars); keeps extension."""
        self._n += 1
        base = posixpath.basename(member_name.replace("\\", "/"))
        safe = re.sub(r"[^A-Za-z0-9._-]+", "_", base)[-60:]
        return f"{self._n:06d}_{safe}"

    # -- walking ----------------------------------------------------------- #
    def walk_dir(self, root: str, prefix: str, depth: int):
        def onerr(e):
            self.skip(to_plain(str(getattr(e, "filename", "") or root)), f"Cannot read folder: {e}")

        for dirpath, dirnames, filenames in os.walk(root, onerror=onerr):
            dirnames[:] = sorted(d for d in dirnames if d not in SKIP_DIR_NAMES)
            rel_dir = os.path.relpath(dirpath, root)
            rel_dir = "" if rel_dir == "." else rel_dir.replace("\\", "/")
            found_in = join_logical(prefix, rel_dir) if rel_dir else (prefix or "(input folder)")
            for fn in sorted(filenames):
                rel = f"{rel_dir}/{fn}" if rel_dir else fn
                self.handle_file(os.path.join(dirpath, fn), rel, join_logical(prefix, rel),
                                 found_in, depth)

    def handle_file(self, real: str, name: str, logical: str, found_in: str, depth: int):
        kind, sub = classify(name)
        if kind == "doc":
            try:
                size = os.path.getsize(real)
            except OSError as e:
                self.skip(logical, f"Cannot access file: {e}")
                return
            self.docs.append(DocItem(len(self.docs) + 1, posixpath.basename(name.replace("\\", "/")),
                                     sub, found_in, logical, real, size))
        elif kind == "archive":
            if depth >= self.max_depth:
                self.skip(logical, f"Archive nesting deeper than --max-depth {self.max_depth}")
                return
            self.extract_and_scan(real, logical, sub, depth + 1)

    # -- archives ---------------------------------------------------------- #
    def extract_and_scan(self, real: str, logical: str, kind: str, depth: int):
        self.archives_opened += 1
        dest = self._new_dir()
        members = None
        err = None
        if kind != "cli":
            try:
                members = getattr(self, f"_x_{kind}")(real, logical, dest)
            except Exception as e:  # archive-level failure -> try 7-Zip CLI
                err = f"{type(e).__name__}: {e}"
                shutil.rmtree(dest, ignore_errors=True)
                os.makedirs(dest, exist_ok=True)
        if members is not None:
            for mname, fpath in members:
                mdir = posixpath.dirname(mname.replace("\\", "/"))
                found_in = join_logical(logical, mdir) if mdir else logical
                self.handle_file(fpath, mname, join_logical(logical, mname), found_in, depth)
            return
        if not self.sevenzip:
            hint = " (install 7-Zip; for RAR also WinRAR/UnRAR works)"
            self.skip(logical, f"Cannot extract{': ' + err if err else ''}{hint}")
            return
        rc, msg = self._cli_extract(real, dest)
        has_files = any(files for _, _, files in os.walk(dest))
        if rc != 0:
            self.skip(logical, f"7-Zip exit code {rc}"
                               f"{' (partially extracted)' if has_files else ''}: {msg}"
                               f"{' | python extractor: ' + err if err else ''}")
        if has_files:
            self.walk_dir(dest, logical, depth)

    def _cli_extract(self, real: str, dest: str):
        # Dummy password: encrypted archives fail fast instead of waiting for input.
        cmd = [self.sevenzip, "x", "-y", "-bd", "-pNO_PASSWORD_GIVEN", f"-o{dest}", real]
        try:
            p = subprocess.run(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                               stderr=subprocess.PIPE, text=True, errors="replace", timeout=7200)
            return p.returncode, " ".join(p.stderr.split())[-300:]
        except Exception as e:
            return -1, str(e)

    def _wanted(self, name: str) -> bool:
        return classify(name)[0] in ("doc", "archive")

    def _x_zip(self, real, logical, dest):
        out = []
        with zipfile.ZipFile(real) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                mname = info.filename
                if not info.flag_bits & 0x800:        # not flagged UTF-8: fix mojibake if possible
                    try:
                        mname = mname.encode("cp437").decode("utf-8")
                    except UnicodeError:
                        pass
                if not self._wanted(mname):
                    continue
                if info.flag_bits & 0x1:
                    self.skip(join_logical(logical, mname), "Password-protected ZIP member")
                    continue
                fp = os.path.join(dest, self._flat_name(mname))
                try:
                    with z.open(info) as src, open(fp, "wb") as dst:
                        shutil.copyfileobj(src, dst, 1 << 20)
                except (zipfile.BadZipFile, zlib.error, EOFError, OSError) as e:
                    self.skip(join_logical(logical, mname), f"Corrupt ZIP member: {e}")
                    continue
                out.append((mname, fp))
        return out

    def _x_7z(self, real, logical, dest):
        import py7zr
        from py7zr.io import Py7zIO, WriterFactory

        class _FileWriter(Py7zIO):
            def __init__(self, path):
                self.f = open(path, "wb") if path else None
                self.n = 0

            def write(self, s):
                if self.f:
                    self.f.write(s)
                self.n += len(s)
                return len(s)

            def read(self, size=None):
                return b""

            def seek(self, offset, whence=0):
                return self.f.seek(offset, whence) if self.f else 0

            def flush(self):
                if self.f:
                    self.f.flush()

            def size(self):
                return self.n

            def close(self):
                if self.f and not self.f.closed:
                    self.f.close()

        class _Factory(WriterFactory):
            def __init__(self, mapping):
                self.mapping, self.writers = mapping, []

            def create(self, filename):
                w = _FileWriter(self.mapping.get(filename))
                self.writers.append(w)
                return w

        with py7zr.SevenZipFile(real, "r") as z:
            if z.needs_password():
                raise RuntimeError("password-protected 7z")
            wanted = [fi.filename for fi in z.list()
                      if not fi.is_directory and self._wanted(fi.filename)]
            if not wanted:
                return []
            mapping = {n: os.path.join(dest, self._flat_name(n)) for n in wanted}
            factory = _Factory(mapping)
            try:
                z.extract(targets=wanted, factory=factory)
            finally:
                for w in factory.writers:
                    w.close()
        return [(n, p) for n, p in mapping.items() if os.path.exists(p)]

    def _x_rar(self, real, logical, dest):
        import rarfile
        out = []
        with rarfile.RarFile(real) as r:
            if r.needs_password():
                raise RuntimeError("password-protected RAR")
            for info in r.infolist():
                if info.is_dir() or not self._wanted(info.filename):
                    continue
                if info.needs_password():
                    self.skip(join_logical(logical, info.filename), "Password-protected RAR member")
                    continue
                fp = os.path.join(dest, self._flat_name(info.filename))
                with r.open(info) as src, open(fp, "wb") as dst:
                    shutil.copyfileobj(src, dst, 1 << 20)
                out.append((info.filename, fp))
        return out

    def _x_tar(self, real, logical, dest):
        out = []
        with tarfile.open(real, "r:*") as t:
            for m in t:
                if not m.isfile() or not self._wanted(m.name):
                    continue
                fp = os.path.join(dest, self._flat_name(m.name))
                src = t.extractfile(m)
                if src is None:
                    continue
                with src, open(fp, "wb") as dst:
                    shutil.copyfileobj(src, dst, 1 << 20)
                out.append((m.name, fp))
        return out

    def _x_single(self, real, logical, dest):
        if tarfile.is_tarfile(real):                  # e.g. "backup.gz" that is really a tar.gz
            return self._x_tar(real, logical, dest)
        base = posixpath.basename(logical.split(" > ")[-1].replace("\\", "/"))
        ext = os.path.splitext(base)[1].lower()
        inner = base[: -len(ext)]
        if not self._wanted(inner):
            return []
        fp = os.path.join(dest, self._flat_name(inner))
        with SINGLE_COMPRESSED[ext](real, "rb") as src, open(fp, "wb") as dst:
            shutil.copyfileobj(src, dst, 1 << 20)
        return [(inner, fp)]


# --------------------------------------------------------------------------- #
# Phase 2: per-document processing (runs in worker processes)
# --------------------------------------------------------------------------- #
def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def run_tesseract(img, opts: Options) -> str:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = opts.tesseract_cmd
    return pytesseract.image_to_string(img, lang=opts.ocr_lang, config=f"--psm {opts.psm}",
                                       timeout=opts.ocr_timeout)


def lo_convert(src: str, fmt: str, tmp: str, soffice: str) -> str:
    """Convert with LibreOffice headless. Uses a private profile so parallel calls don't clash."""
    work = tempfile.mkdtemp(prefix="lo_", dir=tmp)
    inp = os.path.join(work, "in" + os.path.splitext(src)[1].lower())
    shutil.copyfile(src, inp)
    profile = Path(work, "profile").resolve().as_uri()
    cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
           "--convert-to", fmt, "--outdir", work, inp]
    subprocess.run(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=600, check=False)
    out = os.path.join(work, "in." + fmt.split(":")[0])
    if not os.path.isfile(out):
        raise RuntimeError(f"LibreOffice could not convert to {fmt}")
    return out


# ---- PDF ------------------------------------------------------------------ #
def open_pdf(path: str):
    import pymupdf
    try:
        return pymupdf.open(for_native_lib(path))
    except Exception:
        with open(path, "rb") as f:
            return pymupdf.open(stream=f.read(), filetype="pdf")


def page_font_names(page):
    names = []
    for f in page.get_fonts(full=False):
        names.extend(str(x) for x in (f[3], f[4]) if x)
    return names


def image_coverage(page) -> float:
    import pymupdf
    area = abs(page.rect)
    if not area:
        return 0.0
    cov = 0.0
    for info in page.get_image_info():
        b = pymupdf.Rect(info["bbox"]) & page.rect
        if not b.is_empty:
            cov += abs(b)
    return min(cov / area, 1.0)


def mixed_script_ratio(txt: str) -> float:
    """Share of Devanagari-containing tokens that also contain non-Devanagari letters.
    Broken ToUnicode maps produce tokens like 'कʏ' or 'बƓत'; genuine text almost never does."""
    dev = mixed = 0
    for tok in txt.split():
        if DEV_RUN_RE.search(tok):
            dev += 1
            if any(unicodedata.category(c)[0] in "LM" and not DEV_RUN_RE.match(c) for c in tok):
                mixed += 1
    return mixed / dev if dev >= 5 else 0.0


def ocr_reason(page, txt: str, opts: Options, fonts: list[str]):
    if any(LEGACY_FONT_RE.search(f) for f in fonts):
        return "legacy-font"
    chars = [c for c in txt if not c.isspace()]
    if chars:
        bad = sum(1 for c in chars if c == "\ufffd" or "\ue000" <= c <= "\uf8ff")
        if (bad >= 5 and bad / len(chars) > 0.05) or mixed_script_ratio(txt) > 0.10:
            return "garbled-text"
    if sum(1 for c in chars if c.isalnum()) < opts.min_text_chars:
        if page.get_images(full=False) or page.get_drawings():
            return "no-text-layer"
        return None                                   # genuinely blank page
    if count_words(txt)[0] == 0:
        if any(UNICODE_DEVA_FONT_RE.search(f) for f in fonts):
            return "devanagari-font-no-unicode"
        if image_coverage(page) >= opts.image_cover:
            return "image-heavy"
    return None


def ocr_pdf_page(page, opts: Options) -> str:
    import pymupdf
    from PIL import Image
    z = opts.dpi / 72
    longest = max(page.rect.width, page.rect.height) * z
    if longest > 8000:                                # keep huge pages within memory/Tesseract limits
        z *= 8000 / longest
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z), colorspace=pymupdf.csGRAY, alpha=False)
    img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    return run_tesseract(img, opts)


def process_pdf(path: str, opts: Options, r: dict, force_ocr: bool = False) -> str:
    doc = open_pdf(path)
    try:
        if doc.needs_pass and not doc.authenticate(""):
            raise RuntimeError("PDF is password-protected")
        r["pages"] = doc.page_count
        r["page_src"] = r["page_src"] or "PDF"
        parts, reasons, legacy = [], {}, set()
        ocr_done = ocr_missing = 0
        for page in doc:
            txt = page.get_text("text") or ""
            fonts = page_font_names(page)
            legacy.update(f for f in fonts if LEGACY_FONT_RE.search(f))
            reason = None
            if force_ocr or opts.ocr == "always":
                reason = "forced"
            elif opts.ocr == "auto":
                reason = ocr_reason(page, txt, opts, fonts)
            if reason:
                if opts.ocr_available:
                    try:
                        txt = ocr_pdf_page(page, opts)
                        ocr_done += 1
                        reasons[reason] = reasons.get(reason, 0) + 1
                    except Exception as e:
                        r["notes"].append(f"OCR failed on page {page.number + 1}: {e}")
                else:
                    ocr_missing += 1
            parts.append(txt)
        r["ocr_pages"] += ocr_done
        r["ocr_reasons"] = ", ".join(f"{k} ({v})" for k, v in sorted(reasons.items()))
        r["legacy"] = ", ".join(sorted({re.sub(r"^[A-Z]{6}\+", "", f) for f in legacy}))
        n = doc.page_count
        r["method"] = ("Text layer" if ocr_done == 0 else
                       "OCR (all pages)" if ocr_done == n else
                       f"Text layer + OCR ({ocr_done}/{n} pages)")
        if ocr_missing:
            r["notes"].append(f"{ocr_missing} page(s) needed OCR but Tesseract/Hindi data unavailable")
        return "\n".join(parts)
    finally:
        doc.close()


# ---- Word (OOXML) --------------------------------------------------------- #
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _walk_wxml(el, out: list):
    for child in el:
        tag = child.tag
        if tag == MC_FALLBACK:          # duplicate copy of text boxes -> skip (avoids double count)
            continue
        if tag == W_NS + "t":
            out.append(child.text or "")
        elif tag == W_NS + "tab":
            out.append("\t")
        elif tag in (W_NS + "br", W_NS + "cr"):
            out.append("\n")
        elif tag == W_NS + "noBreakHyphen":
            out.append("-")
        _walk_wxml(child, out)
        if tag == W_NS + "p":
            out.append("\n")


def wxml_text(data: bytes) -> str:
    out: list = []
    _walk_wxml(ET.fromstring(data), out)
    return "".join(out)


def docx_main_part(z: zipfile.ZipFile, names: set) -> str:
    try:
        for rel in ET.fromstring(z.read("_rels/.rels")):
            if rel.get("Type", "").endswith("/officeDocument"):
                target = rel.get("Target", "").lstrip("/")
                if target in names:
                    return target
    except (KeyError, ET.ParseError):
        pass
    if "word/document.xml" in names:
        return "word/document.xml"
    raise RuntimeError("Not a valid Word (OOXML) document")


def ocr_docx_images(z: zipfile.ZipFile, media: list, opts: Options, r: dict) -> str:
    import io
    from PIL import Image
    texts, done = [], 0
    for n in media:
        if os.path.splitext(n)[1].lower() not in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif"):
            continue
        try:
            img = Image.open(io.BytesIO(z.read(n)))
            img.load()
            if img.width < 100 or img.height < 100:
                continue
            img = img.convert("L")
            if img.width < 1200:                      # small screenshots OCR better upscaled
                img = img.resize((img.width * 2, img.height * 2))
            texts.append(run_tesseract(img, opts))
            done += 1
        except Exception as e:
            r["notes"].append(f"Image OCR failed ({posixpath.basename(n)}): {e}")
    r["ocr_images"] = done
    return "\n".join(texts)


def process_docx(path: str, opts: Options, r: dict, tmp: str) -> str:
    if not zipfile.is_zipfile(path):                  # e.g. an old .doc renamed to .docx
        r["notes"].append("Not a real DOCX (zip); treated as legacy Word")
        return process_legacy_word(path, opts, r, tmp, force_ext=".doc")
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        main = docx_main_part(z, names)
        base = posixpath.dirname(main)
        pat = re.compile(rf"{re.escape(base)}/(header|footer)\d*\.xml")
        parts = [main] + sorted(n for n in names if pat.fullmatch(n))
        parts += [f"{base}/{p}" for p in ("footnotes.xml", "endnotes.xml") if f"{base}/{p}" in names]
        text = "\n".join(wxml_text(z.read(p)) for p in parts)

        fonts = set()
        for p in [main, f"{base}/styles.xml"]:
            if p in names:
                fonts.update(m.decode("utf-8", "replace") for m in
                             re.findall(rb'w:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"', z.read(p)))
        legacy = sorted(f for f in fonts if LEGACY_FONT_RE.search(f))
        r["legacy"] = ", ".join(legacy)
        media = sorted(n for n in names if n.startswith(f"{base}/media/"))
        r["images"] = len(media)

        pages = None
        if "docProps/app.xml" in names:
            m = re.search(rb"<(?:\w+:)?Pages>(\d+)</", z.read("docProps/app.xml"))
            if m and int(m.group(1)) > 0:
                pages = int(m.group(1))
        r["method"] = "Word XML text"

        alnum = sum(1 for c in text if c.isalnum())
        if media and opts.ocr_available and (
                opts.ocr == "always" or (opts.ocr == "auto" and alnum < opts.min_text_chars)):
            text += "\n" + ocr_docx_images(z, media, opts, r)
            if r.get("ocr_images"):
                r["method"] += f" + image OCR ({r['ocr_images']} images)"
        elif media and alnum < opts.min_text_chars:
            r["notes"].append(f"Almost no text; {len(media)} embedded image(s) not OCRed")

    # Legacy-font Hindi (Kruti Dev etc.) is not Unicode: render via LibreOffice and OCR it.
    if legacy:
        if opts.ocr != "never" and opts.soffice and opts.ocr_available:
            pdf = lo_convert(path, "pdf", tmp, opts.soffice)
            r["page_src"] = "LibreOffice PDF render"
            text = process_pdf(pdf, opts, r, force_ocr=True)
            r["legacy"] = ", ".join(legacy)
            r["method"] = "LibreOffice PDF + OCR (legacy Hindi font)"
            r["notes"].append("Legacy-font text OCRed; accuracy depends on those fonts being "
                              "installed on this PC")
            return text
        r["notes"].append("Legacy (non-Unicode) Hindi font used - that text is NOT counted as Hindi "
                          "(needs LibreOffice + Tesseract Hindi to OCR it)")

    if pages is None or opts.docx_pages == "libreoffice":
        if opts.soffice:
            try:
                pdf = lo_convert(path, "pdf", tmp, opts.soffice)
                doc = open_pdf(pdf)
                pages = doc.page_count
                doc.close()
                r["page_src"] = "LibreOffice PDF render"
            except Exception as e:
                r["notes"].append(f"Page count via LibreOffice failed: {e}")
        if pages is None:
            r["notes"].append("Page count unavailable (no metadata; LibreOffice not found)")
    if pages is not None and not r["page_src"]:
        r["page_src"] = "Word metadata (last save)"
    r["pages"] = pages
    return text


def process_legacy_word(path: str, opts: Options, r: dict, tmp: str, force_ext: str | None = None) -> str:
    if not opts.soffice:
        raise RuntimeError("Legacy .doc/.rtf needs LibreOffice for conversion (not found)")
    src = path
    if force_ext:
        src = os.path.join(tempfile.mkdtemp(dir=tmp), "renamed" + force_ext)
        shutil.copyfile(path, src)
    docx = lo_convert(src, "docx", tmp, opts.soffice)
    text = process_docx(docx, opts, r, tmp)
    r["method"] = "LibreOffice->DOCX, " + r["method"]
    return text


def process_txt(path: str, r: dict) -> str:
    """Plain text: detect BOM / UTF-8 / UTF-16 (Notepad 'Unicode'); fall back to cp1252."""
    with open(path, "rb") as f:
        data = f.read()
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        enc = "utf-16"
    elif data[:3] == b"\xef\xbb\xbf":
        enc = "utf-8-sig"
    elif data and data.count(b"\x00") > len(data) // 4:   # BOM-less UTF-16
        enc = "utf-16-le" if data[1::2].count(b"\x00") > data[0::2].count(b"\x00") else "utf-16-be"
    else:
        try:
            data.decode("utf-8")
            enc = "utf-8"
        except UnicodeDecodeError:
            enc = "cp1252"
            r["notes"].append("Not UTF-8/UTF-16; decoded as cp1252 (Hindi only counted if Unicode)")
    r["method"] = f"Plain text ({enc})"
    r["page_src"] = "n/a (plain text)"
    return data.decode(enc, errors="replace")


def process_item(item: DocItem, opts: Options) -> dict:
    t0 = time.time()
    r = {"pages": None, "page_src": "", "hindi": None, "english": None, "other": None,
         "method": "", "ocr_pages": 0, "ocr_reasons": "", "legacy": "", "images": None,
         "sha": "", "status": "OK", "notes": []}
    try:
        try:
            import pymupdf
            pymupdf.TOOLS.mupdf_display_errors(False)
        except Exception:
            pass
        r["sha"] = sha256_file(item.real)
        with tempfile.TemporaryDirectory(prefix="hwc_w_", ignore_cleanup_errors=True) as tmp:
            if item.ext in PDF_EXTS:
                text = process_pdf(item.real, opts, r)
            elif item.ext in OOXML_EXTS:
                text = process_docx(item.real, opts, r, tmp)
            elif item.ext in TEXT_EXTS:
                text = process_txt(item.real, r)
            else:
                text = process_legacy_word(item.real, opts, r, tmp)
        r["hindi"], r["english"], r["other"] = count_words(text)
        if opts.dump_dir:
            safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", item.name)[:100]
            with open(os.path.join(opts.dump_dir, f"{item.sno:05d}_{safe}.txt"), "w",
                      encoding="utf-8") as f:
                f.write(text)
    except Exception as e:
        r["status"] = "Error"
        r["notes"].append(f"{type(e).__name__}: {e}")
    r["seconds"] = round(time.time() - t0, 2)
    r["notes"] = "; ".join(r["notes"])
    return r


def error_result(msg: str) -> dict:
    return {"pages": None, "page_src": "", "hindi": None, "english": None, "other": None,
            "method": "", "ocr_pages": 0, "ocr_reasons": "", "legacy": "", "images": None,
            "sha": "", "status": "Error", "notes": msg, "seconds": None}


def run_all(items: list[DocItem], opts: Options, workers: int) -> dict:
    results: dict[int, dict] = {}
    if workers <= 1:
        for it in tqdm(items, desc="Processing", unit="doc"):
            results[it.sno] = process_item(it, opts)
        return results
    try:
        with ProcessPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(process_item, it, opts): it for it in items}
            for fut in tqdm(as_completed(futs), total=len(futs), desc="Processing", unit="doc"):
                it = futs[fut]
                try:
                    results[it.sno] = fut.result()
                except BrokenProcessPool:
                    break
                except Exception as e:
                    results[it.sno] = error_result(f"{type(e).__name__}: {e}")
    except BrokenProcessPool:
        pass
    leftovers = [it for it in items if it.sno not in results]
    if leftovers:   # a worker crashed (e.g. a corrupt file killed MuPDF): isolate each remaining doc
        print(f"[!] Worker crash - re-running {len(leftovers)} document(s) one at a time",
              file=sys.stderr)
        for it in tqdm(leftovers, desc="Retrying", unit="doc"):
            try:
                with ProcessPoolExecutor(max_workers=1) as ex:
                    results[it.sno] = ex.submit(process_item, it, opts).result()
            except Exception as e:
                results[it.sno] = error_result(f"Worker crashed on this file: {type(e).__name__}: {e}")
    return results


# --------------------------------------------------------------------------- #
# Excel report
# --------------------------------------------------------------------------- #
COLUMNS = [  # (header, width)
    ("S.No", 7), ("Document Name", 40), ("Type", 7), ("Found In", 45), ("Relative Path", 60),
    ("Size (KB)", 10), ("Pages", 8), ("Page Count Source", 22), ("Hindi Words", 12),
    ("English/Latin Words", 12), ("Other-Script Words", 12), ("Total Words", 12), ("Hindi %", 9),
    ("Hindi Words / Page", 12), ("Extraction Method", 32), ("Pages OCRed", 10),
    ("OCR Reasons", 30), ("Legacy Hindi Fonts", 25), ("Embedded Images", 10),
    ("SHA-256 (first 16)", 20), ("Duplicate Of (S.No)", 12), ("Time (s)", 9), ("Status", 9),
    ("Notes", 70),
]


def write_report(out_path: str, items: list[DocItem], results: dict, skipped: list,
                 meta: list[tuple[str, str]], archives_opened: int) -> str:
    from openpyxl import Workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Font, PatternFill

    font = Font(name="Arial", size=10)
    bold = Font(name="Arial", size=10, bold=True)
    hdr_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    err_fill = PatternFill("solid", fgColor="FCE4D6")

    def clean(v):
        return ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v

    def put(ws, row, col, value, f=font, fmt=None):
        c = ws.cell(row=row, column=col, value=clean(value))
        c.font = f
        if fmt:
            c.number_format = fmt
        return c

    def header(ws, cols, row=1):
        for i, (h, w) in enumerate(cols, 1):
            c = put(ws, row, i, h, hdr_font)
            c.fill = hdr_fill
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[c.column_letter].width = w

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws = wb.create_sheet("Documents")
    ws_skip = wb.create_sheet("Skipped")

    # ---- Documents sheet ----
    header(ws, COLUMNS)
    ws.freeze_panes = "C2"
    for i, it in enumerate(items):
        row = i + 2
        r = results[it.sno]
        vals = [it.sno, it.name, it.ext.lstrip(".").upper(), it.found_in, it.logical,
                round(it.size / 1024, 1), r["pages"], r["page_src"], r["hindi"], r["english"],
                r["other"], None, None, None, r["method"], r["ocr_pages"] or 0, r["ocr_reasons"],
                r["legacy"], r["images"], (r["sha"] or "")[:16], r.get("dup_of"), r["seconds"],
                r["status"], r["notes"]]
        for col, v in enumerate(vals, 1):
            put(ws, row, col, v)
        put(ws, row, 6, vals[5], fmt="#,##0.0")
        put(ws, row, 12, f"=SUM(I{row}:K{row})", fmt="#,##0")
        put(ws, row, 13, f"=IF(L{row}>0,I{row}/L{row},0)", fmt="0.0%")
        put(ws, row, 14, f'=IF(AND(ISNUMBER(G{row}),G{row}>0),I{row}/G{row},"")', fmt="#,##0.0")
        for col in (9, 10, 11):
            ws.cell(row=row, column=col).number_format = "#,##0"
        if r["status"] != "OK":
            for col in range(1, len(COLUMNS) + 1):
                ws.cell(row=row, column=col).fill = err_fill
    last = max(len(items) + 1, 2)
    ws.auto_filter.ref = f"A1:{ws.cell(row=last, column=len(COLUMNS)).coordinate}"

    # ---- Skipped sheet ----
    header(ws_skip, [("Item", 90), ("Reason", 90)])
    for i, (what, why) in enumerate(skipped, 2):
        put(ws_skip, i, 1, what)
        put(ws_skip, i, 2, why)
    last_s = max(len(skipped) + 1, 2)

    # ---- Summary sheet (formulas reference the Documents sheet) ----
    ws_sum.column_dimensions["A"].width = 42
    ws_sum.column_dimensions["B"].width = 70
    ws_sum.column_dimensions["C"].width = 14
    ws_sum.column_dimensions["D"].width = 14
    put(ws_sum, 1, 1, "Hindi Word Count Report", Font(name="Arial", size=14, bold=True))
    row = 3
    for k, v in meta:
        put(ws_sum, row, 1, k, bold)
        put(ws_sum, row, 2, v)
        row += 1
    row += 1
    D = f"Documents!$A$2:$A${last}"
    col = lambda L: f"Documents!${L}$2:${L}${last}"   # noqa: E731
    put(ws_sum, row, 1, "Metric", hdr_font).fill = hdr_fill
    put(ws_sum, row, 2, "Value", hdr_font).fill = hdr_fill
    row += 1
    start = row
    metrics = [
        ("Documents found", f"=COUNTA({D})", "#,##0"),
        ("Processed OK", f'=COUNTIF({col("W")},"OK")', "#,##0"),
        ("Errors", f'=COUNTIF({col("W")},"Error")', "#,##0"),
        ("Duplicate copies (same SHA-256)", f"=COUNT({col('U')})", "#,##0"),
        ("Unique documents", f"=B{start}-B{start + 3}", "#,##0"),
        ("Archives opened (incl. nested)", archives_opened, "#,##0"),
        ("Skipped items (see Skipped sheet)", f"=COUNTA(Skipped!$A$2:$A${last_s})", "#,##0"),
        ("Total pages", f"=SUM({col('G')})", "#,##0"),
        ("Pages OCRed", f"=SUM({col('P')})", "#,##0"),
        ("Total Hindi words (all files)", f"=SUM({col('I')})", "#,##0"),
        ("Total Hindi words (excluding duplicates)", f'=SUMIFS({col("I")},{col("U")},"")', "#,##0"),
        ("Total words (all scripts)", f"=SUM({col('L')})", "#,##0"),
        ("Hindi share of all words", f"=IF(B{start + 11}>0,B{start + 9}/B{start + 11},0)", "0.0%"),
    ]
    for name, formula, fmt in metrics:
        put(ws_sum, row, 1, name)
        c = put(ws_sum, row, 2, formula, fmt=fmt)
        c.alignment = Alignment(horizontal="left")
        if name.startswith("Archives opened"):
            c.comment = Comment("Counted by the script during extraction (not a formula).", "script")
        row += 1

    row += 1
    for i, h in enumerate(["Type", "Files", "Pages", "Hindi Words"], 1):
        put(ws_sum, row, i, h, hdr_font).fill = hdr_fill
    row += 1
    for t in sorted({it.ext.lstrip(".").upper() for it in items}):
        put(ws_sum, row, 1, t)
        put(ws_sum, row, 2, f'=COUNTIF({col("C")},A{row})', fmt="#,##0").alignment = Alignment(horizontal="left")
        put(ws_sum, row, 3, f'=SUMIF({col("C")},A{row},{col("G")})', fmt="#,##0")
        put(ws_sum, row, 4, f'=SUMIF({col("C")},A{row},{col("I")})', fmt="#,##0")
        row += 1

    row += 1
    put(ws_sum, row, 1, "How words are counted", bold)
    notes = [
        "Hindi word = a run of Devanagari letters/vowel-signs (U+0900-097F + extensions) containing "
        "at least one letter. Danda (। ॥) and Devanagari digits are not words.",
        "Devanagari is also used by Marathi, Nepali, Sanskrit etc.; those words are counted as Hindi too.",
        "English/Latin words = words in Latin script (incl. accented). Other-Script Words = any other script (Urdu, Gurmukhi...).",
        "PDF pages are OCRed (auto mode) when there is no text layer, legacy Hindi fonts (Kruti Dev, "
        "Chanakya...) are used, text is garbled, or an image-heavy page has no Hindi text.",
        "OCR counts are approximate (typically within a few % on clean scans).",
        "Word page counts come from the metadata Word saved last time; use --docx-pages libreoffice "
        "for rendered counts.",
    ]
    for n in notes:
        row += 1
        c = put(ws_sum, row, 1, n)
        ws_sum.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws_sum.row_dimensions[row].height = 28

    try:
        wb.save(out_path)
        return out_path
    except PermissionError:   # typically: the file is open in Excel
        p = Path(out_path)
        alt = str(p.with_name(f"{p.stem}_{datetime.now():%Y%m%d_%H%M%S}{p.suffix}"))
        wb.save(alt)
        print(f"[!] '{out_path}' is locked (open in Excel?). Saved to '{alt}' instead.")
        return alt


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description="Count Hindi words in PDF/Word files inside a folder and nested archives.")
    ap.add_argument("input", help="Input folder")
    ap.add_argument("-o", "--output", help="Output .xlsx (default: <folder>_hindi_word_count.xlsx)")
    ap.add_argument("--ocr", choices=["auto", "always", "never"], default="auto",
                    help="auto: OCR only pages that need it (default); always: OCR every PDF page "
                         "and Word images; never: text layer only")
    ap.add_argument("--ocr-lang", default="hin+eng", help="Tesseract languages (default hin+eng)")
    ap.add_argument("--dpi", type=int, default=300, help="OCR render DPI (default 300)")
    ap.add_argument("--psm", type=int, default=3, help="Tesseract page segmentation mode (default 3)")
    ap.add_argument("--ocr-timeout", type=int, default=300, help="Seconds per OCR page/image")
    ap.add_argument("--tesseract", help="Path to tesseract.exe (auto-detected if omitted)")
    ap.add_argument("--tessdata-dir", help="Folder containing hin.traineddata (e.g. tessdata_best)")
    ap.add_argument("--soffice", help="Path to LibreOffice soffice (auto-detected if omitted)")
    ap.add_argument("--sevenzip", help="Path to 7z.exe (auto-detected if omitted)")
    ap.add_argument("--unrar", help="Path to UnRAR.exe (auto-detected if omitted)")
    ap.add_argument("--docx-pages", choices=["metadata", "libreoffice"], default="metadata",
                    help="Word page count source (default: metadata, LibreOffice fallback)")
    ap.add_argument("--workers", type=int, default=max(1, (os.cpu_count() or 2) // 2),
                    help="Parallel worker processes (default: half your CPU cores)")
    ap.add_argument("--max-depth", type=int, default=10, help="Max archive nesting depth")
    ap.add_argument("--min-text-chars", type=int, default=25,
                    help="PDF page with fewer letters than this is treated as scanned")
    ap.add_argument("--work-dir", help="Where archives are extracted (default: system temp). "
                                       "Use a drive with enough free space.")
    ap.add_argument("--keep-temp", action="store_true", help="Keep extracted files")
    ap.add_argument("--dump-text", help="Folder to save extracted text per document (for checking)")
    return ap.parse_args(argv)


def main(argv=None):
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    a = parse_args(argv)

    in_dir = os.path.abspath(a.input)
    if not os.path.isdir(in_dir):
        sys.exit(f"Input folder not found: {in_dir}")
    out = a.output or f"{Path(in_dir).name or 'input'}_hindi_word_count.xlsx"
    if not out.lower().endswith(".xlsx"):
        out += ".xlsx"

    # ---- external tools ----
    if a.tessdata_dir:
        os.environ["TESSDATA_PREFIX"] = os.path.abspath(a.tessdata_dir)
    tess = find_tool(a.tesseract, ["tesseract"], [r"Tesseract-OCR\tesseract.exe"])
    soffice = find_tool(a.soffice, ["soffice", "libreoffice"],
                        [r"LibreOffice\program\soffice.com", r"LibreOffice\program\soffice.exe"])
    sevenzip = find_tool(a.sevenzip, ["7z", "7za", "7zz"], [r"7-Zip\7z.exe"])
    unrar = find_tool(a.unrar, ["unrar"], [r"WinRAR\UnRAR.exe"])
    try:
        import rarfile
        if unrar:
            rarfile.UNRAR_TOOL = unrar
        if sevenzip:
            rarfile.SEVENZIP_TOOL = sevenzip
        rarfile.tool_setup(force=True)
    except Exception:
        pass   # no RAR backend -> 7-Zip CLI fallback / reported in Skipped sheet

    ocr_ok = False
    if a.ocr != "never":
        if not tess:
            print("[!] Tesseract not found - scanned PDFs will NOT be OCRed (see --tesseract).")
        else:
            need = set(a.ocr_lang.split("+"))
            missing = need - tesseract_langs(tess)
            if missing:
                print(f"[!] Tesseract language data missing: {', '.join(sorted(missing))}. "
                      "Install Hindi (hin.traineddata) or pass --tessdata-dir. OCR disabled.")
            else:
                ocr_ok = True
    if not soffice:
        print("[i] LibreOffice not found: .doc/.rtf files will be reported as errors; "
              "Word page counts rely on saved metadata.")
    if not sevenzip:
        print("[i] 7-Zip not found: fallback for unusual archives (Deflate64 zip, RAR without "
              "UnRAR, .iso/.cab) is unavailable.")

    opts = Options(ocr=a.ocr, ocr_lang=a.ocr_lang, dpi=a.dpi, psm=a.psm, ocr_timeout=a.ocr_timeout,
                   tesseract_cmd=tess, ocr_available=ocr_ok, soffice=soffice,
                   docx_pages=a.docx_pages, min_text_chars=a.min_text_chars)
    if a.dump_text:
        os.makedirs(a.dump_text, exist_ok=True)
        opts.dump_dir = os.path.abspath(a.dump_text)
    if a.workers > 1:
        os.environ.setdefault("OMP_THREAD_LIMIT", "1")   # one Tesseract thread per worker

    if a.work_dir:
        os.makedirs(a.work_dir, exist_ok=True)
    work = tempfile.mkdtemp(prefix="hwc_", dir=a.work_dir)
    t0 = time.time()
    try:
        print(f"[1/3] Scanning {in_dir} (folders + archives) ...")
        disc = Discoverer(to_long(work), a.max_depth, sevenzip)
        disc.walk_dir(to_long(in_dir), "", 0)
        items = disc.docs
        print(f"      {len(items)} document(s) found, {disc.archives_opened} archive(s) opened, "
              f"{len(disc.skipped)} item(s) skipped")

        print(f"[2/3] Extracting text / OCR with {a.workers} worker(s) ...")
        results = run_all(items, opts, a.workers) if items else {}

        first_by_hash: dict[str, int] = {}
        for it in items:
            r = results[it.sno]
            if r["sha"]:
                if r["sha"] in first_by_hash:
                    r["dup_of"] = first_by_hash[r["sha"]]
                else:
                    first_by_hash[r["sha"]] = it.sno

        meta = [
            ("Input folder", in_dir),
            ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("OCR mode / language / DPI", f"{a.ocr} / {a.ocr_lang} / {a.dpi}"),
            ("Tesseract", f"{tess} ({'Hindi OCR ready' if ocr_ok else 'OCR unavailable'})"
                          if tess else "not found - no OCR"),
            ("LibreOffice", soffice or "not found"),
            ("7-Zip / UnRAR", f"{sevenzip or 'not found'} / {unrar or 'not found'}"),
            ("Run time (s)", f"{time.time() - t0:.1f}"),
        ]
        print("[3/3] Writing Excel report ...")
        saved = write_report(out, items, results, disc.skipped, meta, disc.archives_opened)
        ok = sum(1 for r in results.values() if r["status"] == "OK")
        hindi = sum(r["hindi"] or 0 for r in results.values())
        print(f"Done: {ok}/{len(items)} OK, {hindi:,} Hindi words total -> {os.path.abspath(saved)}")
    finally:
        if a.keep_temp:
            print(f"[i] Extracted files kept in: {work}")
        else:
            shutil.rmtree(to_long(work), ignore_errors=True)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
