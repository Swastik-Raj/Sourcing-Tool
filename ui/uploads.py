"""Files a person hands to the UI. A product sheet is checked by content, not by its name, and saved under a name the
server builds; the client's file name is only looked at for its extension and is never part of a path. An uploaded file is
never edited afterwards."""
import csv
import io
import os
import secrets
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional

from . import search
from .commands import JobRejected
from .config import MAX_FACTS_BYTES, MAX_PRODUCT_ROWS, MAX_UPLOAD_BYTES, Settings

MAX_UNZIPPED_BYTES = 50 * 1024 * 1024     # a workbook that unpacks to more than this is refused (zip bomb)
MAX_ZIP_ENTRIES = 2000
PREVIEW_ROWS = 10


def _mb(n: int) -> str:
    return f"{n / (1024 * 1024):.0f} MB"


def check_workbook(filename: str, data: bytes) -> dict:
    """Refuses anything that is not a plain .xlsx product sheet; returns its preview (see preview_of)."""
    if not isinstance(filename, str) or not filename.lower().endswith(".xlsx"):
        raise JobRejected("Only .xlsx files can be uploaded (not .xlsm, .xls or .csv).")
    if len(data) > MAX_UPLOAD_BYTES:
        raise JobRejected(f"That file is larger than {_mb(MAX_UPLOAD_BYTES)}.")
    if not data.startswith(b"PK\x03\x04"):
        raise JobRejected("That file is not an Excel .xlsx workbook, whatever its name says.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            infos = z.infolist()
            names = {i.filename for i in infos}
            if len(infos) > MAX_ZIP_ENTRIES or sum(i.file_size for i in infos) > MAX_UNZIPPED_BYTES:
                raise JobRejected(f"That workbook unpacks to more than {_mb(MAX_UNZIPPED_BYTES)} and was refused.")
            if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
                raise JobRejected("That file is not an Excel .xlsx workbook, whatever its name says.")
            types = z.read("[Content_Types].xml").decode("utf-8", "replace").lower()
    except zipfile.BadZipFile:
        raise JobRejected("That file is not an Excel .xlsx workbook, whatever its name says.")
    if "macroenabled" in types or any(n.lower().endswith("vbaproject.bin") for n in names):
        raise JobRejected("That workbook contains macros (.xlsm). Save it as a plain .xlsx and upload that.")
    return preview_of(search.sheet_rows(io.BytesIO(data)))


def preview_of(rows: list) -> dict:
    """{'count', 'products' [(sku, description)], 'first', 'columns' {keyword, product, price: found?}} of a sheet's rows."""
    products, columns = search.products_from_rows(rows)
    if not products:
        raise JobRejected("That file has no products under its 'SKU' heading.")
    if len(products) > MAX_PRODUCT_ROWS:
        raise JobRejected(f"That file has {len(products)} products. One upload takes at most {MAX_PRODUCT_ROWS}.")
    return {"count": len(products), "products": products, "first": products[:PREVIEW_ROWS], "columns": columns}


def _real_folder(folder: Path) -> Path:
    """The folder, created if needed, refused if it is a link to somewhere else."""
    folder.mkdir(parents=True, exist_ok=True)
    if folder.resolve() != folder.parent.resolve() / folder.name:      # only this folder itself: its parents may be reached any way
        raise JobRejected(f"The {folder.name} folder is a link to another place, so nothing was saved.")
    return folder


def _write_new(target: Path, data: bytes) -> None:
    tmp = target.with_name(target.name + ".part")
    tmp.write_bytes(data)
    os.replace(tmp, target)


def save_workbook(s: Settings, filename: str, data: bytes, now: Optional[datetime] = None) -> str:
    """Checks, then saves under Uploads\\input_<date>_<time>_<random>.xlsx. Returns the value the forms use for it."""
    check_workbook(filename, data)
    name = f"input_{now or datetime.now():%Y%m%d_%H%M%S}_{secrets.token_hex(3)}.xlsx"
    _write_new(_real_folder(s.uploads_dir) / name, data)
    return search.UPLOAD_PREFIX + name


def save_facts(s: Settings, filename: str, data: bytes, columns: list, now: Optional[datetime] = None) -> str:
    """A filled facts file: .csv, small, text, and with exactly the columns of the content agent's template. It is saved
    byte for byte as Uploads\\facts_<date>_<time>_<random>.csv and never opened for writing again: what a fact's status
    says is only ever what the person who filled the file wrote. Returns the saved name."""
    if not isinstance(filename, str) or not filename.lower().endswith(".csv"):
        raise JobRejected("Only a .csv product details sheet can be uploaded.")
    if len(data) > MAX_FACTS_BYTES:
        raise JobRejected(f"That product details sheet is larger than {_mb(MAX_FACTS_BYTES)}.")
    if b"\x00" in data or data.startswith(b"PK"):
        raise JobRejected("That file is not a .csv text file, whatever its name says.")
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        raise JobRejected("That product details sheet is not saved as UTF-8 text. In Excel, save it as 'CSV UTF-8' and upload it again.")
    header = next(csv.reader(io.StringIO(text)), [])
    if [h.strip() for h in header] != list(columns):
        raise JobRejected(f"That file's first row is not the product details sheet's columns ({', '.join(columns)}). "
                          "Fill in the sheet from step 1 on this page and upload that.")
    name = f"facts_{now or datetime.now():%Y%m%d_%H%M%S}_{secrets.token_hex(3)}.csv"
    _write_new(_real_folder(s.uploads_dir) / name, data)
    return name


def preview(s: Settings, value: str) -> dict:
    """The preview of a file already in the listings (an upload or a sheet in the project folder)."""
    return preview_of(search.sheet_rows(s.project / search.input_path(s, value)))
