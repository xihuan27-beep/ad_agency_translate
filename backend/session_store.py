"""In-memory session store for the multi-step upload → classify → review → download flow.

Holds only what must live server-side (the source file bytes, rendered slide
images, and extracted text units). Everything else — classification edits,
merge/exclude state, translations, copy selections — is owned by the frontend
and passed directly to the stateless /api/translate/* endpoints.

Single-process, in-memory. Fine for an internal tool with a handful of
concurrent users; sessions are lost on server restart and expire after
SESSION_TTL_SECONDS of inactivity.
"""
import threading
import time
import uuid

SESSION_TTL_SECONDS = 2 * 60 * 60  # 2 hours


class Session:
    def __init__(self, session_id: str):
        self.id = session_id
        self.created_at = time.time()
        self.last_used = time.time()
        self.file_bytes: bytes | None = None
        self.file_type: str = "pptx"  # "pptx" | "docx"
        self.file_name: str = "translated.pptx"
        self.text_units: list[dict] = []
        self.slide_images: list[bytes] = []
        self.slide_count: int = 0
        self.output_bytes: bytes | None = None
        self.output_file_name: str = "output.pptx"

    def touch(self):
        self.last_used = time.time()


class SessionStore:
    def __init__(self):
        self._sessions: dict[str, Session] = {}
        self._lock = threading.Lock()

    def create(self) -> Session:
        sid = uuid.uuid4().hex
        session = Session(sid)
        with self._lock:
            self._sessions[sid] = session
        return session

    def get(self, session_id: str) -> Session | None:
        with self._lock:
            session = self._sessions.get(session_id)
        if session:
            session.touch()
        return session

    def delete(self, session_id: str):
        with self._lock:
            self._sessions.pop(session_id, None)

    def sweep_expired(self):
        now = time.time()
        with self._lock:
            expired = [
                sid for sid, s in self._sessions.items()
                if now - s.last_used > SESSION_TTL_SECONDS
            ]
            for sid in expired:
                del self._sessions[sid]


store = SessionStore()
