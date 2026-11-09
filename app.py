"""Servidor de documentos de Convertilo.

Convierte Word, PowerPoint y Excel con LibreOffice, y PDF -> Word con pdf2docx.
Los archivos viven en una carpeta temporal que se borra al terminar cada conversión.
"""

import asyncio
import os
import re
import shutil
import subprocess
import tempfile
import uuid
from html.parser import HTMLParser
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

MAX_BYTES = int(os.getenv("MAX_MB", "50")) * 1024 * 1024
TIMEOUT = int(os.getenv("TIMEOUT_SECONDS", "180"))
SOFFICE = os.getenv("SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice") or "soffice"
ORIGINS = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "*").split(",") if o.strip()]

WORD = {"doc", "docx", "docm", "dot", "dotx", "dotm", "odt", "ott", "fodt", "rtf", "txt", "md", "html", "htm", "xhtml",
        "wpd", "wps", "pages", "abw", "lwp", "sxw"}
SLIDES = {"ppt", "pptx", "pptm", "pot", "potx", "pps", "ppsx", "odp", "otp", "fodp", "key", "sxi"}
SHEET = {"xls", "xlsx", "xlsm", "xlsb", "xlt", "xltx", "ods", "ots", "fods", "csv", "tsv", "numbers", "sxc", "dbf"}

# Qué destinos acepta cada tipo de entrada.
ALLOWED = {
    **{e: {"pdf", "docx", "doc", "odt", "txt", "rtf", "html", "epub"} for e in WORD},
    **{e: {"pdf", "pptx", "ppt", "odp"} for e in SLIDES},
    **{e: {"pdf", "xlsx", "xls", "csv", "ods"} for e in SHEET},
    "pdf": {"docx", "txt", "pptx"},
}

# Filtros para ABRIR ciertos archivos (por defecto LibreOffice adivina).
IMPORT_FILTERS = {
    ("pdf", "pptx"): "impress_pdf_import",  # cada página del PDF pasa a ser una diapositiva
    ("tsv", "*"): "Text - txt - csv (StarCalc):9,34,76",
}

# Filtros de exportación de LibreOffice.
# Sin comillas: cada filtro va como un solo argumento, así que no hacen falta (y en Linux romperían el nombre).
FILTERS = {
    "pdf": "pdf",
    "docx": "docx:MS Word 2007 XML",
    "odt": "odt",
    "rtf": "rtf",
    "doc": "doc:MS Word 97",
    # Página web con las imágenes metidas dentro del mismo archivo.
    "html": "html:HTML (StarWriter):EmbedImages",
    # LibreOffice se cuelga exportando a texto plano documentos con tablas; se pasa por HTML y se extrae el texto.
    "html_txt": "html:XHTML Writer File:UTF8",
    "epub": "epub",
    "pptx": "pptx:Impress MS PowerPoint 2007 XML",
    "ppt": "ppt:MS PowerPoint 97",
    "odp": "odp",
    "xlsx": "xlsx:Calc MS Excel 2007 XML",
    "xls": "xls:MS Excel 97",
    "ods": "ods",
    "csv": "csv:Text - txt - csv (StarCalc):44,34,76",
}

MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "odt": "application/vnd.oasis.opendocument.text",
    "odp": "application/vnd.oasis.opendocument.presentation",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "txt": "text/plain; charset=utf-8",
    "rtf": "application/rtf",
    "csv": "text/csv; charset=utf-8",
    "doc": "application/msword",
    "ppt": "application/vnd.ms-powerpoint",
    "xls": "application/vnd.ms-excel",
    "html": "text/html; charset=utf-8",
    "epub": "application/epub+zip",
}

app = FastAPI(title="Convertilo - documentos", docs_url=None, redoc_url=None)
app.add_middleware(CORSMiddleware, allow_origins=ORIGINS, allow_methods=["GET", "POST"], allow_headers=["*"])

# LibreOffice consume mucha memoria: máximo 2 conversiones a la vez.
slots = asyncio.Semaphore(int(os.getenv("MAX_PARALLEL", "2")))


@app.get("/")
@app.get("/health")
def health():
    return {"ok": True}


@app.post("/convert")
async def convert(file: UploadFile = File(...), to: str = Form(...)):
    to = to.lower().strip()
    name = Path(file.filename or "archivo").name
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if ext not in ALLOWED:
        raise HTTPException(415, "Este tipo de documento no se puede convertir.")
    if to not in ALLOWED[ext]:
        raise HTTPException(400, f"No se puede pasar de {ext.upper()} a {to.upper()}.")

    work = Path(tempfile.mkdtemp(prefix="conv-"))
    cleanup = BackgroundTask(shutil.rmtree, work, ignore_errors=True)
    try:
        src = work / f"entrada.{ext}"
        size = 0
        with src.open("wb") as fh:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_BYTES:
                    raise HTTPException(413, f"El documento pesa más de {MAX_BYTES // 1024 // 1024} MB.")
                fh.write(chunk)
        if size == 0:
            raise HTTPException(400, "El archivo está vacío.")

        async with slots:
            infilter = IMPORT_FILTERS.get((ext, to)) or IMPORT_FILTERS.get((ext, "*"))
            if ext == "pdf" and to == "txt":
                out = await asyncio.wait_for(asyncio.to_thread(pdf_to_txt, src, work), TIMEOUT)
            elif ext == "pdf" and to == "docx":
                out = await asyncio.wait_for(asyncio.to_thread(pdf_to_docx, src, work), TIMEOUT)
            elif to == "txt":
                html = await asyncio.to_thread(run_soffice, src, "html_txt", work, infilter)
                out = await asyncio.to_thread(html_to_txt, html, work)
            else:
                out = await asyncio.to_thread(run_soffice, src, to, work, infilter)

        stem = name.rsplit(".", 1)[0]
        return FileResponse(out, media_type=MIME.get(to), filename=f"{stem}.{to}", background=cleanup)
    except HTTPException:
        shutil.rmtree(work, ignore_errors=True)
        raise
    except asyncio.TimeoutError:
        shutil.rmtree(work, ignore_errors=True)
        raise HTTPException(504, "El documento tardó demasiado en convertirse. Prueba con uno más pequeño.")
    except Exception as e:  # noqa: BLE001
        shutil.rmtree(work, ignore_errors=True)
        print("ERROR", type(e).__name__, e, flush=True)
        raise HTTPException(500, "No se pudo convertir este documento. Puede que esté dañado o protegido con contraseña.")


def run_soffice(src: Path, to: str, work: Path, infilter: str | None = None) -> Path:
    # Un perfil de LibreOffice por conversión permite convertir varias a la vez sin chocar.
    profile = (work / f"perfil-{uuid.uuid4().hex}").resolve().as_uri()
    outdir = work / f"salida-{uuid.uuid4().hex[:6]}"
    outdir.mkdir()
    cmd = [SOFFICE, f"-env:UserInstallation={profile}", "--headless", "--norestore", "--nologo", "--nodefault", "--nolockcheck"]
    if infilter:
        cmd.append(f"--infilter={infilter}")
    cmd += ["--convert-to", FILTERS[to], "--outdir", str(outdir), str(src)]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=TIMEOUT)
    except subprocess.TimeoutExpired as e:
        raise asyncio.TimeoutError() from e
    ext = FILTERS[to].split(":", 1)[0]
    files = list(outdir.glob(f"*.{ext}"))
    if not files:
        raise RuntimeError(f"LibreOffice no produjo salida: {proc.returncode} {proc.stderr.strip()[:300]}")
    return files[0]


class _TextExtractor(HTMLParser):
    BLOCKS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "tr", "br", "table", "ul", "ol"}

    def __init__(self):
        super().__init__()
        self.parts: list[str] = []
        self.skip = 0
        self.cell = 0  # dentro de una celda los párrafos no parten la fila

    def handle_starttag(self, tag, attrs):
        if tag in ("head", "style", "script"):
            self.skip += 1
        elif tag in ("td", "th"):
            self.cell += 1
            self.parts.append("\t")
        elif tag in self.BLOCKS and not (self.cell and tag != "table"):
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("head", "style", "script"):
            self.skip -= 1
        elif tag in ("td", "th"):
            self.cell -= 1
            self.parts.append(" ")
        elif tag in self.BLOCKS and not self.cell:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.skip:
            self.parts.append(data.replace("\n", " "))


def html_to_txt(html: Path, work: Path) -> Path:
    p = _TextExtractor()
    p.feed(html.read_text(encoding="utf-8", errors="replace"))
    lines = [re.sub(r"[  ]+", " ", l).strip(" ") for l in "".join(p.parts).split("\n")]
    lines = [re.sub(r"\s*\t\s*", "\t", l).strip("\t ").replace("\t", "  |  ") if "\t" in l else l for l in lines]
    text = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip() + "\n"
    out = work / "salida.txt"
    out.write_text(text, encoding="utf-8-sig")  # con BOM para que el Bloc de notas muestre bien las tildes
    return out


def pdf_to_txt(src: Path, work: Path) -> Path:
    # PyMuPDF (viene con pdf2docx) lee bien las tildes de PDF hechos con Calibri, donde pdf.js se equivoca.
    import pymupdf

    with pymupdf.open(src) as doc:
        if doc.needs_pass:
            raise HTTPException(400, "Este PDF tiene contraseña. Quítasela primero y vuelve a intentarlo.")
        pages = [page.get_text().strip() for page in doc]
    text = "\n\n".join(pages).strip()
    if not text:
        raise HTTPException(422, "Este PDF no tiene texto (parece una foto escaneada). Conviértelo a JPG o a Word.")
    out = work / "salida.txt"
    out.write_text(text + "\n", encoding="utf-8-sig")
    return out


def pdf_to_docx(src: Path, work: Path) -> Path:
    from pdf2docx import Converter

    out = work / "salida.docx"
    cv = Converter(str(src))
    try:
        cv.convert(str(out))
    finally:
        cv.close()
    if not out.exists() or out.stat().st_size == 0:
        raise RuntimeError("pdf2docx no produjo salida")
    return out
