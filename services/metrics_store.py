import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock


class MetricsStore:
    """Camada simples de métricas locais para saúde/performance."""

    def __init__(self, db_path: str = "logs/metrics.db"):
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
                CREATE TABLE IF NOT EXISTS run_metrics (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ts TEXT NOT NULL,
                    formato TEXT NOT NULL,
                    status TEXT NOT NULL,
                    total_entradas INTEGER NOT NULL,
                    total_invalidos INTEGER NOT NULL,
                    total_processado INTEGER NOT NULL,
                    duracao_s REAL NOT NULL,
                    throughput_itens_s REAL NOT NULL,
                    erro TEXT
                )
                """
            )
            conn.commit()

    def record_run(
        self,
        *,
        formato: str,
        status: str,
        total_entradas: int,
        total_invalidos: int,
        total_processado: int,
        duracao_s: float,
        erro: str = "",
    ):
        duracao = max(0.0, float(duracao_s))
        processado = max(0, int(total_processado))
        throughput = processado / duracao if duracao > 0 else 0.0
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO run_metrics (
                    ts, formato, status, total_entradas, total_invalidos, total_processado,
                    duracao_s, throughput_itens_s, erro
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    self._agora_iso(),
                    formato,
                    status,
                    int(total_entradas),
                    int(total_invalidos),
                    processado,
                    duracao,
                    throughput,
                    erro,
                ),
            )
            conn.commit()

    def get_health_snapshot(self) -> dict:
        with self._lock, sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            total_runs = conn.execute("SELECT COUNT(*) AS c FROM run_metrics").fetchone()["c"]
            avg_duration = conn.execute(
                "SELECT COALESCE(AVG(duracao_s), 0) AS avg_duracao FROM run_metrics"
            ).fetchone()["avg_duracao"]
            avg_throughput = conn.execute(
                "SELECT COALESCE(AVG(throughput_itens_s), 0) AS avg_tp FROM run_metrics"
            ).fetchone()["avg_tp"]
            err_rate = conn.execute(
                """
                SELECT COALESCE(AVG(CASE WHEN status = 'error' THEN 1.0 ELSE 0.0 END), 0) AS err_rate
                FROM run_metrics
                """
            ).fetchone()["err_rate"]
            by_formato = conn.execute(
                """
                SELECT formato,
                       COUNT(*) AS runs,
                       COALESCE(AVG(duracao_s), 0) AS avg_duracao,
                       COALESCE(AVG(throughput_itens_s), 0) AS avg_throughput,
                       COALESCE(AVG(CASE WHEN status = 'error' THEN 1.0 ELSE 0.0 END), 0) AS erro_rate
                FROM run_metrics
                GROUP BY formato
                ORDER BY runs DESC
                """
            ).fetchall()

        return {
            "total_runs": int(total_runs),
            "avg_duration_s": float(avg_duration),
            "avg_throughput_itens_s": float(avg_throughput),
            "error_rate": float(err_rate),
            "by_formato": [dict(row) for row in by_formato],
        }
