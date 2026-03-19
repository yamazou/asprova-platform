import csv
import io
import re
from datetime import datetime
from typing import Any, Dict, Iterable, Iterator, List, Mapping, Optional, Sequence, Tuple, Union


def decode_csv_content(content_bytes: bytes) -> str:
    """
    Decode uploaded CSV bytes to text.
    Uses utf-8-sig to gracefully handle BOM.
    """
    return content_bytes.decode("utf-8-sig", errors="replace")


def detect_delimiter(content: str, *, sample_size: int = 2000) -> str:
    """Detect delimiter by max occurrence in the first sample_size characters."""
    sample = content[:sample_size]
    delimiter = ","
    for d in [",", "\t", ";", "|"]:
        if sample.count(d) > sample.count(delimiter):
            delimiter = d
    return delimiter


def csv_dict_reader_from_content(
    content: str, *, delimiter: Optional[str] = None
) -> Tuple[csv.DictReader, str]:
    """
    Return a csv.DictReader and the delimiter used.
    """
    used_delimiter = delimiter or detect_delimiter(content)
    reader = csv.DictReader(io.StringIO(content), delimiter=used_delimiter)
    return reader, used_delimiter


def csv_dict_reader_from_bytes(
    content_bytes: bytes, *, delimiter: Optional[str] = None
) -> Tuple[csv.DictReader, str, str]:
    """
    Decode bytes, detect delimiter (if not provided), and return a DictReader.
    Returns: (reader, used_delimiter, decoded_text)
    """
    content = decode_csv_content(content_bytes)
    reader, used_delimiter = csv_dict_reader_from_content(content, delimiter=delimiter)
    return reader, used_delimiter, content


def rows_to_csv(rows: Iterable[Sequence[Any]], headers: Sequence[str]) -> str:
    """
    Convert DB/query rows to CSV string.
    Matches the previous bridge behavior:
    - datetime -> YYYY/MM/DD
    - None -> empty
    """
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(list(headers))

    for row in rows:
        values = list(row)
        formatted: List[str] = []
        for v in values:
            if v is None:
                formatted.append("")
            elif isinstance(v, datetime):
                formatted.append(v.strftime("%Y/%m/%d"))
            else:
                formatted.append(str(v))
        writer.writerow(formatted)

    return output.getvalue()


_ANSI_ESCAPE_RE = re.compile(r"\x1b\[[0-9;]*m")


def strip_ansi(text: str) -> str:
    """Utility if some sources include ANSI escape sequences."""
    return _ANSI_ESCAPE_RE.sub("", text)

