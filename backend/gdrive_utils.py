import re
import requests


def gdrive_file_id(url: str) -> str | None:
    patterns = [
        r"/file/d/([a-zA-Z0-9_-]+)",
        r"[?&]id=([a-zA-Z0-9_-]+)",
        r"/presentation/d/([a-zA-Z0-9_-]+)",
        r"/document/d/([a-zA-Z0-9_-]+)",
        r"/spreadsheets/d/([a-zA-Z0-9_-]+)",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None


def gdrive_is_slides(url: str) -> bool:
    return bool(re.search(r"/presentation/d/", url))


def gdrive_is_docs(url: str) -> bool:
    return bool(re.search(r"/document/d/", url))


def is_zip(data: bytes) -> bool:
    """PPTX and DOCX are ZIP files — both start with PK magic bytes."""
    return len(data) >= 4 and data[:4] == b"PK\x03\x04"


def _gdrive_confirm_url(html_text: str, file_id: str) -> str | None:
    """Extract confirmed download URL from Google's virus-scan warning page."""
    m = re.search(r"confirm=([0-9A-Za-z_-]+)", html_text)
    if m:
        return f"https://drive.google.com/uc?export=download&id={file_id}&confirm={m.group(1)}"
    return None


def download_gdrive(file_id: str, is_slides: bool = False, is_docs: bool = False) -> bytes:
    session = requests.Session()
    errors = []

    if is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and is_zip(resp.content):
                return resp.content
            errors.append(f"Slides export: HTTP {resp.status_code}, size={len(resp.content)}")
        except Exception as e:
            errors.append(f"Slides export: {e}")

    if is_docs:
        url = f"https://docs.google.com/document/d/{file_id}/export?format=docx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and is_zip(resp.content):
                return resp.content
            errors.append(f"Docs export: HTTP {resp.status_code}, size={len(resp.content)}")
        except Exception as e:
            errors.append(f"Docs export: {e}")

    # Direct usercontent download (works for regular Drive files)
    url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&authuser=0&confirm=t"
    try:
        resp = session.get(url, timeout=300)
        if resp.status_code == 200 and is_zip(resp.content):
            return resp.content
        # Google may return an HTML virus-scan confirmation page
        if resp.status_code == 200 and resp.content[:1] == b"<":
            confirm = _gdrive_confirm_url(resp.text, file_id)
            if confirm:
                resp2 = session.get(confirm, timeout=300)
                if resp2.status_code == 200 and is_zip(resp2.content):
                    return resp2.content
        errors.append(f"usercontent: HTTP {resp.status_code}, size={len(resp.content)}")
    except Exception as e:
        errors.append(f"usercontent: {e}")

    # Slides export fallback (try even for non-Slides URLs — works if it happens to be Slides)
    if not is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and is_zip(resp.content):
                return resp.content
            errors.append(f"Slides fallback: HTTP {resp.status_code}")
        except Exception as e:
            errors.append(f"Slides fallback: {e}")

    # Legacy /uc endpoint with cookie + HTML confirm handling
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    try:
        resp = session.get(url, timeout=300)
        for key, value in resp.cookies.items():
            if key.startswith("download_warning"):
                url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={value}"
                resp = session.get(url, timeout=300)
                break
        if not is_zip(resp.content) and resp.content[:1] == b"<":
            confirm = _gdrive_confirm_url(resp.text, file_id)
            if confirm:
                resp = session.get(confirm, timeout=300)
        if resp.status_code == 200 and is_zip(resp.content):
            return resp.content
        errors.append(f"legacy uc: HTTP {resp.status_code}, size={len(resp.content)}")
    except Exception as e:
        errors.append(f"legacy uc: {e}")
    raise RuntimeError(f"모든 다운로드 방법 실패: {'; '.join(errors)}")
