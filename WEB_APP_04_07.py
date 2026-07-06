
"""
WEB APP 04_07 - FinancePlus Email Azienda Streamlit PRO
File unico Python/Streamlit per ricerca azienda, archivio email, allegati,
anteprima intelligente, download selettivo e cartella cliente automatica.

Esecuzione:
    pip install streamlit pandas
    streamlit run WEB_APP_04_07.py

Dipendenze opzionali consigliate:
    pip install pymupdf pdfplumber pillow pytesseract openpyxl

Configurazione email:
    - Da interfaccia laterale, oppure
    - Variabili ambiente / Streamlit Secrets:
        FP_IMAP_HOST
        FP_IMAP_PORT
        FP_IMAP_USER
        FP_IMAP_PASSWORD
        FP_IMAP_FOLDER
"""

from __future__ import annotations

import base64
import dataclasses
import datetime as dt
import email
from email.header import decode_header, make_header
import hashlib
import html
import imaplib
import io
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import textwrap
import zipfile
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import pandas as pd
except Exception:  # pragma: no cover
    pd = None

try:
    import streamlit as st
except Exception:  # pragma: no cover
    st = None


APP_NAME = "WEB APP 04_07"
APP_VERSION = "1.0"
ROOT_DIR = Path(os.getenv("FP_DATA_DIR", "financeplus_data")).resolve()
DB_PATH = ROOT_DIR / "financeplus_04_07.db"
CLIENTI_DIR = ROOT_DIR / "clienti"
EXPORT_DIR = ROOT_DIR / "export"
TMP_DIR = ROOT_DIR / "tmp"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg", ".ico", ".tif", ".tiff"}
DOC_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".txt", ".rtf",
    ".xml", ".p7m", ".eml", ".msg", ".zip", ".rar", ".7z", ".odt", ".ods"
}

KEYWORDS_TIPOLOGIA = {
    "visura": ["visura", "camera di commercio", "camerale", "cciaa", "registro imprese"],
    "bilancio": ["bilancio", "stato patrimoniale", "conto economico", "nota integrativa", "xbrl"],
    "centrale_rischi": ["centrale rischi", "crif", "banca d'italia", "cr banca", "segnalazione"],
    "estratto_conto": ["estratto conto", "movimenti", "conto corrente", "saldo", "iban"],
    "contratto": ["contratto", "fornitura", "ordine", "accordo", "mandato"],
    "documentazione": ["documentazione", "documenti", "allegati", "file richiesti"],
    "preventivo": ["preventivo", "offerta", "proposta economica", "quotazione"],
    "fattura": ["fattura", "invoice", "proforma", "nota credito", "ricevuta"],
    "identita": ["carta identita", "carta d'identita", "patente", "passaporto", "documento identita"],
    "report": ["report", "relazione", "dossier", "analisi", "valutazione"],
}

COMPANY_STOPWORDS = {
    "srl", "s.r.l", "spa", "s.p.a", "societa", "azienda", "documentazione",
    "documenti", "fattura", "visura", "bilancio", "preventivo", "contratto",
    "pec", "mail", "richiesta", "allegati", "fw", "fwd", "re", "rif"
}


@dataclasses.dataclass
class ImapConfig:
    host: str
    port: int
    username: str
    password: str
    folder: str = "INBOX"
    use_ssl: bool = True


@dataclasses.dataclass
class AttachmentInfo:
    filename: str
    content_type: str
    payload: bytes
    size: int
    sha256: str
    extension: str
    is_image: bool
    kind: str = "documentazione"


@dataclasses.dataclass
class ParsedEmail:
    uid: str
    message_id: str
    subject: str
    sender: str
    recipients: str
    date: str
    body_text: str
    body_html: str
    raw_bytes: bytes
    attachments: List[AttachmentInfo]
    summary: str
    matched_company: str


def ensure_dirs() -> None:
    for path in (ROOT_DIR, CLIENTI_DIR, EXPORT_DIR, TMP_DIR):
        path.mkdir(parents=True, exist_ok=True)


def init_db() -> None:
    ensure_dirs()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS aziende (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ragione_sociale TEXT UNIQUE NOT NULL,
                partita_iva TEXT,
                codice_fiscale TEXT,
                amministratore TEXT,
                note TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS email_archiviate (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                uid TEXT,
                message_id TEXT,
                azienda TEXT NOT NULL,
                mittente TEXT,
                destinatari TEXT,
                oggetto TEXT,
                data_email TEXT,
                sintesi TEXT,
                cartella TEXT,
                eml_path TEXT,
                body_path TEXT,
                sha256 TEXT,
                created_at TEXT NOT NULL,
                UNIQUE(message_id, azienda)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS allegati (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_id INTEGER,
                azienda TEXT NOT NULL,
                filename_originale TEXT,
                filename_archivio TEXT,
                path TEXT,
                tipo TEXT,
                content_type TEXT,
                size INTEGER,
                sha256 TEXT,
                created_at TEXT NOT NULL,
                UNIQUE(sha256, path),
                FOREIGN KEY(email_id) REFERENCES email_archiviate(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS documenti_locali (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                azienda TEXT NOT NULL,
                filename_originale TEXT,
                filename_archivio TEXT,
                path TEXT,
                tipo TEXT,
                size INTEGER,
                sha256 TEXT,
                testo_estratto TEXT,
                created_at TEXT NOT NULL,
                UNIQUE(sha256, azienda)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                azione TEXT NOT NULL,
                dettagli TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def log_action(action: str, details: str = "") -> None:
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO audit_log (azione, dettagli, created_at) VALUES (?, ?, ?)",
            (action, details, now_iso()),
        )
        conn.commit()


def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def ddmmyyyy(value: Optional[str] = None) -> str:
    parsed = parse_date(value) if value else dt.datetime.now()
    return parsed.strftime("%d-%m-%Y")


def parse_date(value: Any) -> dt.datetime:
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time.min)
    text = str(value or "").strip()
    if not text:
        return dt.datetime.now()
    try:
        return email.utils.parsedate_to_datetime(text).replace(tzinfo=None)
    except Exception:
        pass
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(text[:10], fmt)
        except Exception:
            continue
    return dt.datetime.now()


def sanitize_name(value: str, default: str = "SENZA_NOME", max_len: int = 90) -> str:
    value = html.unescape(value or "").strip()
    value = re.sub(r"[^\w\s\-.&]", " ", value, flags=re.UNICODE)
    value = re.sub(r"\s+", "_", value)
    value = value.strip("._- ")
    if not value:
        value = default
    return value[:max_len]


def normalize_company(value: str) -> str:
    text = (value or "").strip()
    text = re.sub(r"\s+", " ", text)
    return text.upper()


def sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload or b"").hexdigest()


def decode_mime_header(value: Any) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        try:
            return value.decode("utf-8", errors="ignore") if isinstance(value, bytes) else str(value)
        except Exception:
            return ""


def get_text_from_part(part: email.message.Message) -> str:
    payload = part.get_payload(decode=True)
    if not payload:
        return ""
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except Exception:
        return payload.decode("utf-8", errors="replace")


def strip_html(value: str) -> str:
    if not value:
        return ""
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?is)<br\s*/?>", "\n", value)
    value = re.sub(r"(?is)</p>", "\n", value)
    value = re.sub(r"(?is)<.*?>", " ", value)
    value = html.unescape(value)
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def classify_text(text: str, filename: str = "") -> str:
    haystack = f"{filename} {text}".lower()
    for tipo, words in KEYWORDS_TIPOLOGIA.items():
        for word in words:
            if word in haystack:
                return tipo
    return "documentazione"


def smart_summary(text: str, subject: str = "", max_chars: int = 700) -> str:
    clean = re.sub(r"\s+", " ", strip_html(text or "")).strip()
    if not clean:
        return f"Email senza testo leggibile. Oggetto: {subject}".strip()
    sentences = re.split(r"(?<=[.!?])\s+", clean)
    selected = []
    priority_words = [
        "alleg", "document", "bilancio", "visura", "fattura", "contratto",
        "preventivo", "centrale rischi", "scaden", "pagamento", "finanziamento",
        "richiesta", "azienda", "cliente"
    ]
    for sentence in sentences:
        if any(word in sentence.lower() for word in priority_words):
            selected.append(sentence)
        if len(" ".join(selected)) > max_chars:
            break
    if not selected:
        selected = sentences[:3]
    summary = " ".join(selected)
    summary = summary[:max_chars].strip()
    if len(summary) == max_chars:
        summary = summary.rsplit(" ", 1)[0] + "..."
    return summary


def detect_company_from_text(text: str, fallback: str = "") -> str:
    """
    Priorita:
    1. Denominazioni con S.r.l./S.p.A. ecc.
    2. Stringhe dopo parole chiave: documentazione, cliente, azienda, societa.
    3. Fallback digitato dall'utente.
    """
    clean = html.unescape(text or "")
    patterns = [
        r"\b([A-Z0-9][A-Z0-9&\-\.\s]{2,70}\s+(?:S\.?R\.?L\.?|SRL|S\.?P\.?A\.?|SPA|SNC|SAS|S\.?A\.?S\.?|S\.?N\.?C\.?))\b",
        r"(?:documentazione|documenti|pratica|cliente|azienda|societ[àa]|rif\.?)\s+([A-Z0-9][A-Z0-9&\-\.\s]{2,55})",
    ]
    upper = clean.upper()
    for pattern in patterns:
        found = re.findall(pattern, upper, flags=re.IGNORECASE)
        for item in found:
            candidate = item if isinstance(item, str) else item[0]
            candidate = re.sub(r"\s+", " ", candidate).strip(" -_.,;:")
            if candidate and len(candidate) >= 3:
                return normalize_company(candidate)
    return normalize_company(fallback)


def get_client_folder(company: str) -> Path:
    company_safe = sanitize_name(normalize_company(company), default="CLIENTE")
    folder = CLIENTI_DIR / company_safe
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def subject_folder(company: str, subject: str, date_text: str, sender: str = "") -> Path:
    """
    Struttura robusta:
    financeplus_data/clienti/AZIENDA/GG-MM-AAAA_OGGETTO/
    L'oggetto email governa la cartella operativa, come richiesto.
    """
    base = get_client_folder(company)
    d = ddmmyyyy(date_text)
    subj = sanitize_name(subject or "EMAIL", max_len=80)
    folder = base / f"{d}_{subj}"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def attachment_archive_name(company: str, original_name: str, kind: str, date_text: str) -> str:
    ext = Path(original_name or "").suffix.lower()
    if not ext:
        ext = ".bin"
    company_safe = sanitize_name(normalize_company(company), default="CLIENTE", max_len=60)
    kind_safe = sanitize_name(kind or "documentazione", default="documentazione", max_len=35)
    return f"{company_safe}_{kind_safe}_{ddmmyyyy(date_text)}{ext}"


def is_document_file(filename: str, skip_images: bool = True) -> bool:
    ext = Path(filename or "").suffix.lower()
    if skip_images and ext in IMAGE_EXTENSIONS:
        return False
    if ext in DOC_EXTENSIONS:
        return True
    if not ext:
        return True
    return not skip_images


def parse_email(uid: str, raw_bytes: bytes, fallback_company: str = "") -> ParsedEmail:
    msg = email.message_from_bytes(raw_bytes)
    subject = decode_mime_header(msg.get("Subject"))
    sender = decode_mime_header(msg.get("From"))
    recipients = decode_mime_header(msg.get("To"))
    date_text = decode_mime_header(msg.get("Date"))
    message_id = decode_mime_header(msg.get("Message-ID")) or sha256_bytes(raw_bytes)[:32]

    body_text_parts: List[str] = []
    body_html_parts: List[str] = []
    attachments: List[AttachmentInfo] = []

    if msg.is_multipart():
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition") or "").lower()
            content_type = part.get_content_type() or ""
            filename = decode_mime_header(part.get_filename())
            is_attachment = "attachment" in content_disposition or bool(filename)

            if is_attachment:
                payload = part.get_payload(decode=True) or b""
                if not filename:
                    filename = f"allegato_{len(attachments) + 1}.bin"
                ext = Path(filename).suffix.lower()
                kind = classify_text("", filename)
                attachments.append(
                    AttachmentInfo(
                        filename=filename,
                        content_type=content_type,
                        payload=payload,
                        size=len(payload),
                        sha256=sha256_bytes(payload),
                        extension=ext,
                        is_image=ext in IMAGE_EXTENSIONS or content_type.startswith("image/"),
                        kind=kind,
                    )
                )
                continue

            if content_type == "text/plain":
                body_text_parts.append(get_text_from_part(part))
            elif content_type == "text/html":
                body_html_parts.append(get_text_from_part(part))
    else:
        content_type = msg.get_content_type()
        if content_type == "text/html":
            body_html_parts.append(get_text_from_part(msg))
        else:
            body_text_parts.append(get_text_from_part(msg))

    body_html = "\n".join(body_html_parts).strip()
    body_text = "\n".join(body_text_parts).strip() or strip_html(body_html)
    search_text = f"{subject}\n{sender}\n{body_text}"
    matched_company = detect_company_from_text(search_text, fallback_company)
    summary = smart_summary(body_text or body_html, subject)

    return ParsedEmail(
        uid=str(uid),
        message_id=message_id,
        subject=subject,
        sender=sender,
        recipients=recipients,
        date=date_text,
        body_text=body_text,
        body_html=body_html,
        raw_bytes=raw_bytes,
        attachments=attachments,
        summary=summary,
        matched_company=matched_company,
    )


def save_parsed_email(parsed: ParsedEmail, company: str, skip_images: bool = True) -> Dict[str, Any]:
    init_db()
    final_company = detect_company_from_text(f"{parsed.subject}\n{parsed.body_text}", company) or company
    folder = subject_folder(final_company, parsed.subject, parsed.date, parsed.sender)

    email_hash = sha256_bytes(parsed.raw_bytes)
    eml_path = unique_path(folder / "email_originale.eml")
    eml_path.write_bytes(parsed.raw_bytes)

    summary_path = unique_path(folder / "sintesi_email.txt")
    summary_path.write_text(
        "\n".join([
            f"Progetto: {APP_NAME}",
            f"Azienda: {final_company}",
            f"Data email: {parsed.date}",
            f"Mittente: {parsed.sender}",
            f"Destinatari: {parsed.recipients}",
            f"Oggetto: {parsed.subject}",
            "",
            "SINTESI INTELLIGENTE",
            parsed.summary,
            "",
            "TESTO EMAIL",
            parsed.body_text[:12000],
        ]),
        encoding="utf-8",
    )

    saved_attachments = []
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            """
            INSERT OR IGNORE INTO email_archiviate
            (uid, message_id, azienda, mittente, destinatari, oggetto, data_email,
             sintesi, cartella, eml_path, body_path, sha256, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                parsed.uid, parsed.message_id, final_company, parsed.sender, parsed.recipients,
                parsed.subject, parsed.date, parsed.summary, str(folder), str(eml_path),
                str(summary_path), email_hash, now_iso()
            )
        )
        conn.commit()
        email_db_id = cur.lastrowid
        if not email_db_id:
            row = conn.execute(
                "SELECT id FROM email_archiviate WHERE message_id = ? AND azienda = ?",
                (parsed.message_id, final_company)
            ).fetchone()
            email_db_id = row[0] if row else None

        for attachment in parsed.attachments:
            if skip_images and attachment.is_image:
                continue
            if not is_document_file(attachment.filename, skip_images=skip_images):
                continue
            archive_name = attachment_archive_name(final_company, attachment.filename, attachment.kind, parsed.date)
            out_path = unique_path(folder / archive_name)
            out_path.write_bytes(attachment.payload)
            saved_attachments.append(str(out_path))

            conn.execute(
                """
                INSERT OR IGNORE INTO allegati
                (email_id, azienda, filename_originale, filename_archivio, path, tipo,
                 content_type, size, sha256, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    email_db_id, final_company, attachment.filename, out_path.name, str(out_path),
                    attachment.kind, attachment.content_type, attachment.size, attachment.sha256, now_iso()
                )
            )
        conn.commit()

    log_action("SCARICA_EMAIL", f"{final_company} - {parsed.subject}")
    return {
        "azienda": final_company,
        "cartella": str(folder),
        "eml_path": str(eml_path),
        "summary_path": str(summary_path),
        "allegati_salvati": saved_attachments,
    }


def secret_or_env(name: str, default: str = "") -> str:
    """Legge prima le variabili ambiente, poi Streamlit Secrets.

    Supporta sia chiavi dirette, es. FP_IMAP_HOST, sia sezione [imap]
    in .streamlit/secrets.toml o nei Secrets di Streamlit Cloud.
    """
    value = os.getenv(name, "").strip()
    if value:
        return value
    if st is None:
        return default
    try:
        if name in st.secrets:
            return str(st.secrets.get(name, default)).strip()
    except Exception:
        pass
    try:
        imap = st.secrets.get("imap", {})
        key_map = {
            "FP_IMAP_HOST": "host",
            "FP_IMAP_PORT": "port",
            "FP_IMAP_USER": "user",
            "FP_IMAP_PASSWORD": "password",
            "FP_IMAP_FOLDER": "folder",
        }
        key = key_map.get(name, name)
        if key in imap:
            return str(imap.get(key, default)).strip()
    except Exception:
        pass
    return default


def get_imap_config_from_env() -> Optional[ImapConfig]:
    host = secret_or_env("FP_IMAP_HOST", "")
    username = secret_or_env("FP_IMAP_USER", "")
    password = secret_or_env("FP_IMAP_PASSWORD", "")
    if not host or not username or not password:
        return None
    try:
        port = int(secret_or_env("FP_IMAP_PORT", "993") or "993")
    except ValueError:
        port = 993
    folder = secret_or_env("FP_IMAP_FOLDER", "INBOX") or "INBOX"
    return ImapConfig(host=host, port=port, username=username, password=password, folder=folder)


def imap_connect(cfg: ImapConfig) -> imaplib.IMAP4:
    if cfg.use_ssl:
        client: imaplib.IMAP4 = imaplib.IMAP4_SSL(cfg.host, cfg.port)
    else:
        client = imaplib.IMAP4(cfg.host, cfg.port)
    client.login(cfg.username, cfg.password)
    client.select(cfg.folder)
    return client


def imap_date(value: Optional[dt.date]) -> Optional[str]:
    if not value:
        return None
    return value.strftime("%d-%b-%Y")


def build_imap_criteria(keyword: str, date_from: Optional[dt.date] = None, date_to: Optional[dt.date] = None) -> List[Tuple[str, List[str]]]:
    """
    Ricerche separate per compatibilita IMAP:
    - SUBJECT
    - FROM
    - BODY
    - TEXT
    """
    keyword = (keyword or "").strip()
    criteria: List[Tuple[str, List[str]]] = []
    base: List[str] = ["ALL"]
    if date_from:
        base += ["SINCE", imap_date(date_from) or ""]
    if date_to:
        # BEFORE e' esclusivo: aggiungo un giorno.
        before = date_to + dt.timedelta(days=1)
        base += ["BEFORE", imap_date(before) or ""]
    for field in ("SUBJECT", "FROM", "BODY", "TEXT"):
        if keyword:
            criteria.append((field, base + [field, f'"{keyword}"']))
    if not keyword:
        criteria.append(("ALL", base))
    return criteria


def imap_search_uids(cfg: ImapConfig, keyword: str, date_from: Optional[dt.date] = None, date_to: Optional[dt.date] = None, max_results: int = 100) -> List[str]:
    client = imap_connect(cfg)
    found: List[str] = []
    seen = set()
    try:
        for label, criteria in build_imap_criteria(keyword, date_from, date_to):
            try:
                status, data = client.uid("SEARCH", None, *criteria)
            except Exception:
                status, data = client.uid("SEARCH", None, "TEXT", f'"{keyword}"')
            if status != "OK":
                continue
            for raw in (data[0] or b"").split():
                uid = raw.decode()
                if uid not in seen:
                    seen.add(uid)
                    found.append(uid)
                if len(found) >= max_results:
                    return found
    finally:
        try:
            client.logout()
        except Exception:
            pass
    return found


def imap_fetch_raw(cfg: ImapConfig, uid: str) -> bytes:
    client = imap_connect(cfg)
    try:
        status, data = client.uid("FETCH", str(uid), "(RFC822)")
        if status != "OK" or not data:
            raise RuntimeError(f"Email UID {uid} non letta")
        for item in data:
            if isinstance(item, tuple):
                return item[1]
        raise RuntimeError(f"Payload email UID {uid} non disponibile")
    finally:
        try:
            client.logout()
        except Exception:
            pass


def search_and_parse_emails(
    cfg: ImapConfig,
    company: str,
    date_from: Optional[dt.date] = None,
    date_to: Optional[dt.date] = None,
    max_results: int = 50,
) -> List[ParsedEmail]:
    uids = imap_search_uids(cfg, company, date_from, date_to, max_results=max_results)
    parsed: List[ParsedEmail] = []
    for uid in uids:
        try:
            raw = imap_fetch_raw(cfg, uid)
            parsed.append(parse_email(uid, raw, fallback_company=company))
        except Exception as exc:
            log_action("ERRORE_LETTURA_EMAIL", f"UID {uid}: {exc}")
    return parsed


def dataframe(rows: List[Dict[str, Any]]):
    if pd is not None:
        return pd.DataFrame(rows)
    return rows


def db_query(sql: str, params: Sequence[Any] = ()) -> List[Dict[str, Any]]:
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return [dict(row) for row in conn.execute(sql, params).fetchall()]


def list_companies() -> List[str]:
    rows = db_query(
        """
        SELECT azienda, COUNT(*) AS n
        FROM (
            SELECT azienda FROM email_archiviate
            UNION ALL
            SELECT azienda FROM documenti_locali
        )
        GROUP BY azienda
        ORDER BY azienda
        """
    )
    return [r["azienda"] for r in rows]


def archive_stats() -> Dict[str, int]:
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        return {
            "aziende": conn.execute("SELECT COUNT(DISTINCT azienda) FROM email_archiviate").fetchone()[0],
            "email": conn.execute("SELECT COUNT(*) FROM email_archiviate").fetchone()[0],
            "allegati": conn.execute("SELECT COUNT(*) FROM allegati").fetchone()[0],
            "documenti": conn.execute("SELECT COUNT(*) FROM documenti_locali").fetchone()[0],
        }


def zip_paths(paths: Sequence[str], zip_name: str = "export.zip") -> Path:
    ensure_dirs()
    out = unique_path(EXPORT_DIR / zip_name)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in paths:
            p = Path(item)
            if not p.exists():
                continue
            if p.is_file():
                zf.write(p, arcname=p.name)
            elif p.is_dir():
                for child in p.rglob("*"):
                    if child.is_file():
                        zf.write(child, arcname=str(child.relative_to(p.parent)))
    return out


def extract_text_from_file(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in {".txt", ".csv", ".xml", ".rtf"}:
        try:
            return path.read_text(encoding="utf-8", errors="ignore")[:15000]
        except Exception:
            return ""
    if ext == ".pdf":
        # PyMuPDF opzionale
        try:
            import fitz  # type: ignore
            doc = fitz.open(str(path))
            text = "\n".join(page.get_text() for page in doc)
            return text[:20000]
        except Exception:
            return ""
    return ""


def save_uploaded_documents(company: str, uploaded_files: Sequence[Any], skip_images: bool = True) -> List[Dict[str, Any]]:
    init_db()
    company_final = normalize_company(company)
    base = get_client_folder(company_final) / f"{ddmmyyyy()}_IMPORT_DOCUMENTI"
    base.mkdir(parents=True, exist_ok=True)
    saved: List[Dict[str, Any]] = []

    with sqlite3.connect(DB_PATH) as conn:
        for file in uploaded_files:
            name = getattr(file, "name", "documento.bin")
            if not is_document_file(name, skip_images=skip_images):
                continue
            payload = file.getvalue()
            digest = sha256_bytes(payload)
            kind = classify_text("", name)
            archive_name = attachment_archive_name(company_final, name, kind, now_iso())
            out_path = unique_path(base / archive_name)
            out_path.write_bytes(payload)
            text = extract_text_from_file(out_path)
            if text:
                detected = detect_company_from_text(text, company_final)
                if detected and detected != company_final:
                    company_final = detected
                    corrected_base = get_client_folder(company_final) / f"{ddmmyyyy()}_IMPORT_DOCUMENTI"
                    corrected_base.mkdir(parents=True, exist_ok=True)
                    corrected_path = unique_path(corrected_base / out_path.name)
                    shutil.move(str(out_path), corrected_path)
                    out_path = corrected_path
            conn.execute(
                """
                INSERT OR IGNORE INTO documenti_locali
                (azienda, filename_originale, filename_archivio, path, tipo, size, sha256, testo_estratto, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (company_final, name, out_path.name, str(out_path), kind, len(payload), digest, text[:5000], now_iso())
            )
            saved.append({
                "azienda": company_final,
                "file_originale": name,
                "file_archivio": out_path.name,
                "tipo": kind,
                "path": str(out_path),
                "size": len(payload),
            })
        conn.commit()
    log_action("IMPORT_DOCUMENTI", f"{company_final} - {len(saved)} file")
    return saved


def import_local_folder(company: str, folder_path: str, recursive: bool = True, skip_images: bool = True) -> List[Dict[str, Any]]:
    """
    Funzione utile in locale: Streamlit Cloud non puo aprire cartelle del PC.
    Inserendo un percorso locale, importa anche le sottocartelle.
    """
    source = Path(folder_path).expanduser()
    if not source.exists() or not source.is_dir():
        raise FileNotFoundError(f"Cartella non trovata: {source}")
    class LocalUpload:
        def __init__(self, path: Path):
            self.name = path.name
            self._payload = path.read_bytes()
        def getvalue(self) -> bytes:
            return self._payload

    files = []
    iterator = source.rglob("*") if recursive else source.glob("*")
    for p in iterator:
        if p.is_file() and is_document_file(p.name, skip_images=skip_images):
            files.append(LocalUpload(p))
    return save_uploaded_documents(company, files, skip_images=skip_images)


def render_css() -> None:
    if st is None:
        return
    st.markdown(
        """
        <style>
        :root {
          --fp-blue: #10253f;
          --fp-blue2: #173b63;
          --fp-copper: #b97935;
          --fp-bg: #f4f7fb;
          --fp-card: #ffffff;
        }
        .stApp {
          background: linear-gradient(180deg, #f6f8fb 0%, #eef3f8 100%);
        }
        [data-testid="stSidebar"] {
          background: linear-gradient(180deg, #10253f 0%, #173b63 100%);
        }
        [data-testid="stSidebar"] * {
          color: white !important;
        }
        .main-title {
          padding: 1.1rem 1.2rem;
          border-radius: 18px;
          background: linear-gradient(135deg, #10253f 0%, #173b63 65%, #b97935 100%);
          color: white;
          box-shadow: 0 8px 28px rgba(16,37,63,.18);
          margin-bottom: 1rem;
        }
        .main-title h1 { margin: 0; font-size: 2rem; }
        .main-title p { margin: .25rem 0 0 0; opacity: .92; }
        .fp-card {
          background: white;
          border: 1px solid rgba(16,37,63,.08);
          border-radius: 16px;
          padding: 1rem;
          box-shadow: 0 6px 18px rgba(16,37,63,.08);
          margin-bottom: 1rem;
        }
        .metric-card {
          background: white;
          border-top: 4px solid #b97935;
          border-radius: 15px;
          padding: .85rem;
          box-shadow: 0 6px 18px rgba(16,37,63,.08);
        }
        .small-muted { color: #5c6b7a; font-size: .9rem; }
        .badge {
          display: inline-block;
          padding: .2rem .55rem;
          border-radius: 999px;
          background: rgba(185,121,53,.13);
          color: #7b4b1d;
          font-weight: 700;
          font-size: .78rem;
          margin-right: .25rem;
        }
        .email-box {
          border: 1px solid #d9e2ec;
          border-left: 5px solid #b97935;
          border-radius: 12px;
          padding: .9rem;
          background: #fff;
          margin: .6rem 0;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        f"""
        <div class="main-title">
          <h1>{APP_NAME}</h1>
          <p>FinancePlus - ricerca azienda, email, allegati, archivio cliente e anteprima intelligente</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def sidebar_config() -> Tuple[Optional[ImapConfig], Dict[str, Any]]:
    st.sidebar.title("FinancePlus")
    st.sidebar.caption(f"{APP_NAME} v{APP_VERSION}")

    st.sidebar.subheader("Configurazione email")
    env_cfg = get_imap_config_from_env()
    use_env = bool(env_cfg)
    if use_env:
        st.sidebar.success("Configurazione email caricata da variabili ambiente.")

    host = st.sidebar.text_input("IMAP host", value=env_cfg.host if env_cfg else "imap.gmail.com")
    port = st.sidebar.number_input("Porta", value=int(env_cfg.port if env_cfg else 993), min_value=1, max_value=65535)
    username = st.sidebar.text_input("Email utente", value=env_cfg.username if env_cfg else "")
    password = st.sidebar.text_input("Password / App password", value=env_cfg.password if env_cfg else "", type="password")
    folder = st.sidebar.text_input("Cartella IMAP", value=env_cfg.folder if env_cfg else "INBOX")
    use_ssl = st.sidebar.checkbox("Usa SSL", value=True)

    cfg = None
    if host and username and password:
        cfg = ImapConfig(host=host, port=int(port), username=username, password=password, folder=folder, use_ssl=use_ssl)

    st.sidebar.subheader("Opzioni archivio")
    options = {
        "skip_images": st.sidebar.checkbox("Non scaricare immagini automatiche", value=True),
        "max_results": st.sidebar.slider("Numero massimo email", 10, 300, 80, 10),
        "data_root": str(ROOT_DIR),
    }

    st.sidebar.info("La password non viene salvata nel database. Per Streamlit Cloud usa Secrets o variabili ambiente.")
    return cfg, options


def page_dashboard() -> None:
    stats = archive_stats()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Aziende", stats["aziende"])
    c2.metric("Email archiviate", stats["email"])
    c3.metric("Allegati", stats["allegati"])
    c4.metric("Documenti locali", stats["documenti"])

    st.markdown("### Stato progetto")
    st.markdown(
        """
        <div class="fp-card">
        <span class="badge">Streamlit</span>
        <span class="badge">SQLite</span>
        <span class="badge">IMAP</span>
        <span class="badge">Archivio Cliente</span>
        <span class="badge">Download ZIP</span>
        <p class="small-muted">
        Il file unico include dashboard, comando <b>Cerca Azienda</b>, anteprima email,
        sintesi testuale, scarico massivo, scarico selettivo, esclusione automatica immagini,
        cartella cliente e database locale.
        </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    recent = db_query(
        """
        SELECT azienda, oggetto, mittente, data_email, cartella
        FROM email_archiviate
        ORDER BY id DESC
        LIMIT 15
        """
    )
    st.markdown("### Ultime email archiviate")
    if recent:
        st.dataframe(dataframe(recent), use_container_width=True)
    else:
        st.info("Nessuna email archiviata. Usa il comando Cerca Azienda.")


def render_email_preview(parsed: ParsedEmail, index: int) -> None:
    st.markdown(
        f"""
        <div class="email-box">
        <b>{index}. {html.escape(parsed.subject or '(senza oggetto)')}</b><br>
        <span class="small-muted">Da: {html.escape(parsed.sender)} | Data: {html.escape(parsed.date)} | Azienda rilevata: {html.escape(parsed.matched_company)}</span><br><br>
        <b>Sintesi:</b> {html.escape(parsed.summary)}
        </div>
        """,
        unsafe_allow_html=True,
    )
    if parsed.attachments:
        rows = []
        for att in parsed.attachments:
            rows.append({
                "file": att.filename,
                "tipo": att.kind,
                "content_type": att.content_type,
                "dimensione_kb": round(att.size / 1024, 1),
                "immagine": "SI" if att.is_image else "NO",
            })
        st.dataframe(dataframe(rows), use_container_width=True, hide_index=True)
    else:
        st.caption("Nessun allegato rilevato.")


def page_cerca_azienda(cfg: Optional[ImapConfig], options: Dict[str, Any]) -> None:
    st.markdown("## CERCA AZIENDA")
    st.caption("Scrivi il nome azienda. Puoi vedere tutto, scaricare tutto o selezionare singole email/allegati.")

    col1, col2, col3 = st.columns([2, 1, 1])
    company = col1.text_input("Nome azienda / ragione sociale", placeholder="Esempio: BELGARDEN, PELCOM, ETS GROUP")
    date_from = col2.date_input("Data inizio", value=None, format="DD/MM/YYYY")
    date_to = col3.date_input("Data fine", value=None, format="DD/MM/YYYY")

    if "search_results" not in st.session_state:
        st.session_state["search_results"] = []
    if "search_company" not in st.session_state:
        st.session_state["search_company"] = ""

    b1, b2, b3 = st.columns(3)
    cerca = b1.button("🔎 CERCA AZIENDA", type="primary", use_container_width=True)
    vedi = b2.button("👁️ VEDI TUTTO", use_container_width=True)
    scarica = b3.button("⬇️ SCARICA TUTTO", use_container_width=True)

    if (cerca or vedi or scarica) and not company:
        st.warning("Inserisci il nome azienda.")
        return
    if (cerca or vedi or scarica) and cfg is None:
        st.error("Configura l'accesso IMAP nella barra laterale prima di cercare le email.")
        return

    if cerca or vedi or scarica:
        with st.spinner("Ricerca email e allegati in corso..."):
            results = search_and_parse_emails(
                cfg=cfg,
                company=company,
                date_from=date_from,
                date_to=date_to,
                max_results=int(options["max_results"]),
            )
            st.session_state["search_results"] = results
            st.session_state["search_company"] = company
        st.success(f"Trovate {len(st.session_state['search_results'])} email per {company}.")

    results: List[ParsedEmail] = st.session_state.get("search_results", [])
    if results:
        st.markdown("### Risultati")
        for i, parsed in enumerate(results, start=1):
            with st.expander(f"{i}. {parsed.subject or '(senza oggetto)'}", expanded=i <= 3):
                render_email_preview(parsed, i)

        st.markdown("### Scarico selettivo")
        labels = [
            f"{i+1}. {p.subject[:70] or '(senza oggetto)'} - {p.date}"
            for i, p in enumerate(results)
        ]
        selected_labels = st.multiselect("Seleziona email da scaricare", labels, default=[])
        if st.button("⬇️ SCARICA SELEZIONATI", use_container_width=True):
            selected_idx = [labels.index(label) for label in selected_labels]
            saved_paths = []
            for idx in selected_idx:
                out = save_parsed_email(results[idx], company, skip_images=bool(options["skip_images"]))
                saved_paths.append(out["cartella"])
            if saved_paths:
                zip_file = zip_paths(saved_paths, f"{sanitize_name(company)}_selezionati_{ddmmyyyy()}.zip")
                st.success(f"Scaricate {len(saved_paths)} email selezionate.")
                st.download_button(
                    "Scarica ZIP selezionati",
                    data=zip_file.read_bytes(),
                    file_name=zip_file.name,
                    mime="application/zip",
                    use_container_width=True,
                )
            else:
                st.info("Nessuna email selezionata.")

    if scarica and results:
        saved_paths = []
        progress = st.progress(0)
        for i, parsed in enumerate(results, start=1):
            out = save_parsed_email(parsed, company, skip_images=bool(options["skip_images"]))
            saved_paths.append(out["cartella"])
            progress.progress(i / len(results))
        zip_file = zip_paths(saved_paths, f"{sanitize_name(company)}_scarica_tutto_{ddmmyyyy()}.zip")
        st.success(f"Scaricate e archiviate {len(saved_paths)} email nella cartella cliente.")
        st.download_button(
            "Scarica ZIP completo",
            data=zip_file.read_bytes(),
            file_name=zip_file.name,
            mime="application/zip",
            use_container_width=True,
        )


def page_archivio() -> None:
    st.markdown("## Archivio clienti")
    companies = list_companies()
    selected = st.selectbox("Cliente / azienda", [""] + companies)
    if not selected:
        st.info("Seleziona un cliente per vedere email, allegati e documenti.")
        return

    folder = get_client_folder(selected)
    st.markdown(f"**Cartella cliente:** `{folder}`")

    t1, t2, t3 = st.tabs(["Email", "Allegati", "Documenti locali"])

    with t1:
        rows = db_query(
            """
            SELECT id, data_email, mittente, oggetto, sintesi, cartella
            FROM email_archiviate
            WHERE azienda = ?
            ORDER BY id DESC
            """,
            (selected,),
        )
        if rows:
            st.dataframe(dataframe(rows), use_container_width=True)
        else:
            st.info("Nessuna email archiviata per questo cliente.")

    with t2:
        rows = db_query(
            """
            SELECT id, filename_originale, filename_archivio, tipo, size, path
            FROM allegati
            WHERE azienda = ?
            ORDER BY id DESC
            """,
            (selected,),
        )
        if rows:
            st.dataframe(dataframe(rows), use_container_width=True)
            paths = [r["path"] for r in rows]
            if st.button("Scarica tutti gli allegati del cliente", use_container_width=True):
                zip_file = zip_paths(paths, f"{sanitize_name(selected)}_allegati_{ddmmyyyy()}.zip")
                st.download_button(
                    "Download ZIP allegati",
                    data=zip_file.read_bytes(),
                    file_name=zip_file.name,
                    mime="application/zip",
                    use_container_width=True,
                )
        else:
            st.info("Nessun allegato archiviato per questo cliente.")

    with t3:
        rows = db_query(
            """
            SELECT id, filename_originale, filename_archivio, tipo, size, path
            FROM documenti_locali
            WHERE azienda = ?
            ORDER BY id DESC
            """,
            (selected,),
        )
        if rows:
            st.dataframe(dataframe(rows), use_container_width=True)
        else:
            st.info("Nessun documento locale archiviato per questo cliente.")

    if folder.exists() and st.button("Prepara ZIP cartella completa cliente", use_container_width=True):
        zip_file = zip_paths([str(folder)], f"{sanitize_name(selected)}_cartella_cliente_{ddmmyyyy()}.zip")
        st.download_button(
            "Download ZIP cartella cliente",
            data=zip_file.read_bytes(),
            file_name=zip_file.name,
            mime="application/zip",
            use_container_width=True,
        )


def page_importa_documenti(options: Dict[str, Any]) -> None:
    st.markdown("## Importa documenti")
    st.caption("Importa PDF, Excel, Word, TXT, EML e altri file documentali. Le immagini possono essere escluse.")

    company = st.text_input("Azienda di destinazione", placeholder="Ragione sociale cliente")
    uploaded = st.file_uploader(
        "Carica file singoli o multipli",
        accept_multiple_files=True,
        type=None,
    )
    if st.button("Importa file caricati", type="primary", use_container_width=True):
        if not company:
            st.warning("Inserisci l'azienda di destinazione.")
        elif not uploaded:
            st.warning("Carica almeno un file.")
        else:
            saved = save_uploaded_documents(company, uploaded, skip_images=bool(options["skip_images"]))
            st.success(f"Importati {len(saved)} documenti.")
            if saved:
                st.dataframe(dataframe(saved), use_container_width=True)

    st.divider()
    st.markdown("### Importa cartella locale con sottocartelle")
    st.caption("Disponibile quando l'app gira sul tuo PC/server. Su Streamlit Cloud il browser non puo leggere cartelle locali.")
    local_company = st.text_input("Azienda per import cartella", key="local_company")
    folder_path = st.text_input("Percorso cartella locale", placeholder=r"C:\Users\...\Documenti\Cliente")
    recursive = st.checkbox("Leggi anche sottocartelle", value=True)
    if st.button("Importa cartella locale", use_container_width=True):
        try:
            if not local_company or not folder_path:
                st.warning("Inserisci azienda e percorso cartella.")
            else:
                saved = import_local_folder(local_company, folder_path, recursive=recursive, skip_images=bool(options["skip_images"]))
                st.success(f"Importati {len(saved)} file dalla cartella.")
                if saved:
                    st.dataframe(dataframe(saved), use_container_width=True)
        except Exception as exc:
            st.error(str(exc))


def page_database() -> None:
    st.markdown("## Database e controlli")
    stats = archive_stats()
    st.json(stats)

    st.markdown("### Audit log")
    rows = db_query("SELECT created_at, azione, dettagli FROM audit_log ORDER BY id DESC LIMIT 100")
    if rows:
        st.dataframe(dataframe(rows), use_container_width=True)
    else:
        st.info("Audit log vuoto.")

    st.markdown("### Esporta database")
    if DB_PATH.exists():
        st.download_button(
            "Scarica SQLite DB",
            data=DB_PATH.read_bytes(),
            file_name=DB_PATH.name,
            mime="application/octet-stream",
            use_container_width=True,
        )


def page_guida() -> None:
    st.markdown("## Guida rapida")
    st.markdown(
        """
        ### 1. Avvio
        ```bash
        pip install streamlit pandas
        streamlit run WEB_APP_04_07.py
        ```

        ### 2. Configurazione email
        Inserisci host IMAP, email e password applicativa nella barra laterale.
        Per Gmail/Google Workspace usa una **App Password**, non la password normale.

        ### 3. Cerca azienda
        Apri **Cerca Azienda**, scrivi la ragione sociale e usa:
        - **CERCA AZIENDA** per trovare le email;
        - **VEDI TUTTO** per visualizzare anteprima, sintesi e allegati;
        - **SCARICA TUTTO** per creare automaticamente la cartella cliente.

        ### 4. Regola cartella cliente
        L'archivio viene creato così:

        ```text
        financeplus_data/
          clienti/
            AZIENDA/
              GG-MM-AAAA_OGGETTO_EMAIL/
                email_originale.eml
                sintesi_email.txt
                AZIENDA_tipologia_GG-MM-AAAA.pdf
        ```

        ### 5. Regola immagini
        Per impostazione predefinita le immagini non vengono scaricate automaticamente.
        Puoi cambiare questa opzione nella barra laterale.

        ### 6. GitHub e Streamlit Cloud
        Carica questo file su GitHub come `WEB_APP_04_07.py`.
        Su Streamlit Cloud imposta eventuali Secrets/variabili ambiente:

        ```text
        FP_IMAP_HOST=imap.gmail.com
        FP_IMAP_PORT=993
        FP_IMAP_USER=nome@email.it
        FP_IMAP_PASSWORD=app_password
        FP_IMAP_FOLDER=INBOX
        ```

        ### 7. Sicurezza
        Il database salva indici, sintesi, percorsi e hash. Non salva la password email.
        """
    )


def main() -> None:
    if st is None:
        print("Streamlit non installato. Esegui: pip install streamlit pandas")
        return

    ensure_dirs()
    init_db()
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="📁",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    render_css()
    render_header()
    cfg, options = sidebar_config()

    page = st.sidebar.radio(
        "Menu",
        [
            "Dashboard",
            "Cerca Azienda",
            "Archivio Clienti",
            "Importa Documenti",
            "Database",
            "Guida",
        ],
    )

    if page == "Dashboard":
        page_dashboard()
    elif page == "Cerca Azienda":
        page_cerca_azienda(cfg, options)
    elif page == "Archivio Clienti":
        page_archivio()
    elif page == "Importa Documenti":
        page_importa_documenti(options)
    elif page == "Database":
        page_database()
    elif page == "Guida":
        page_guida()


if __name__ == "__main__":
    main()
