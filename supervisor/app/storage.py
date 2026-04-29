"""In-memory file storage for the MVP.

Stores uploaded + processed files keyed by a UUID. Files are auto-purged
after `MAX_AGE_SECONDS`. No DB, no disk persistence — fits in a single
process and matches the privacy story ("we don't store your files").

For production:
- Replace with disk-backed temp storage (tempfile.NamedTemporaryFile)
- Add explicit cleanup endpoint
- Consider object storage (S3/R2/etc) if we need horizontal scaling
"""

import time
import uuid
from dataclasses import dataclass


MAX_AGE_SECONDS = 60 * 60  # 1 hour


@dataclass
class StoredFile:
    """A file held in memory with metadata."""
    file_id: str
    filename: str           # original upload name
    content: bytes          # the raw xlsx bytes
    created_at: float       # unix timestamp
    review_summary: dict    # counts by issue type, for the API response


_store: dict[str, StoredFile] = {}


def put(filename: str, content: bytes, review_summary: dict | None = None) -> str:
    """Store bytes, return a UUID handle."""
    _purge_expired()
    file_id = str(uuid.uuid4())
    _store[file_id] = StoredFile(
        file_id=file_id,
        filename=filename,
        content=content,
        created_at=time.time(),
        review_summary=review_summary or {},
    )
    return file_id


def get(file_id: str) -> StoredFile | None:
    _purge_expired()
    return _store.get(file_id)


def remove(file_id: str) -> None:
    _store.pop(file_id, None)


def _purge_expired() -> None:
    now = time.time()
    expired = [fid for fid, f in _store.items() if now - f.created_at > MAX_AGE_SECONDS]
    for fid in expired:
        _store.pop(fid, None)
