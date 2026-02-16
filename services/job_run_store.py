import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from uuid import uuid4


class JobRunStore:
    """Persistência de execuções de geração em SQLite."""

    def __init__(self, db_path: str = "logs/jobs.db"):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = Lock()
        self._init_db()

    @staticmethod
    def _agora_iso() -> str:
        return datetime.now(timezone.utc).isoformat()

    def _init_db(self):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS job_runs (
                    id TEXT PRIMARY KEY,
                    created_at TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    finished_at TEXT,
                    status TEXT NOT NULL,
                    formato TEXT,
                    tipo_codigo TEXT,
                    modo TEXT,
                    destino TEXT,
                    total_entradas INTEGER NOT NULL,
                    total_invalidos INTEGER NOT NULL,
                    total_processado INTEGER NOT NULL DEFAULT 0,
                    erro TEXT
                )
                """
            )
            conn.commit()

    def create_run(
        self,
        *,
        formato: str,
        tipo_codigo: str,
        modo: str,
        destino: str,
        total_entradas: int,
        total_invalidos: int,
    ) -> str:
        job_id = str(uuid4())
        now = self._agora_iso()
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO job_runs (
                    id, created_at, started_at, status, formato, tipo_codigo, modo,
                    destino, total_entradas, total_invalidos, total_processado
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    job_id,
                    now,
                    now,
                    "started",
                    formato,
                    tipo_codigo,
                    modo,
                    destino,
                    int(total_entradas),
                    int(total_invalidos),
                    0,
                ),
            )
            conn.commit()
        return job_id

    def update_progress(self, job_id: str, processado: int):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "UPDATE job_runs SET total_processado = ?, status = ? WHERE id = ?",
                (int(processado), "running", job_id),
            )
            conn.commit()

    def finish_run(self, job_id: str, status: str, erro: str = "", processado: int | None = None):
        with self._lock, sqlite3.connect(self.db_path) as conn:
            if processado is None:
                conn.execute(
                    "UPDATE job_runs SET finished_at = ?, status = ?, erro = ? WHERE id = ?",
                    (self._agora_iso(), status, erro, job_id),
                )
            else:
                conn.execute(
                    """
                    UPDATE job_runs
                    SET finished_at = ?, status = ?, erro = ?, total_processado = ?
                    WHERE id = ?
                    """,
                    (self._agora_iso(), status, erro, int(processado), job_id),
                )
            conn.commit()
