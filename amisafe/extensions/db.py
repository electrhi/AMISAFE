"""DB 접속 및 스키마 초기화."""
import psycopg
from amisafe.config import Config


def get_conn():
    """psycopg 커넥션 반환. DATABASE_URL 미설정 시 명확한 에러."""
    if not Config.DATABASE_URL:
        raise ValueError("DATABASE_URL 환경변수가 설정되지 않았습니다.")
    return psycopg.connect(Config.DATABASE_URL)


def init_db():
    """필수 테이블 및 유니크 인덱스 생성 (멱등)."""
    sql_list = [
        """
        CREATE TABLE IF NOT EXISTS documents (
            id SERIAL PRIMARY KEY,
            form_id TEXT NOT NULL,
            form_type TEXT NOT NULL,
            group_name TEXT,
            user_id TEXT,
            work_date DATE NOT NULL,
            status TEXT NOT NULL DEFAULT 'in_progress',
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            completed_at TIMESTAMP
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS document_participants (
            id SERIAL PRIMARY KEY,
            document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
            user_id TEXT NOT NULL,
            user_name TEXT NOT NULL,
            group_name TEXT,
            role TEXT NOT NULL,
            slot_index INTEGER,
            is_done BOOLEAN NOT NULL DEFAULT FALSE,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS document_values (
            id SERIAL PRIMARY KEY,
            document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
            field_id TEXT NOT NULL,
            user_id TEXT,
            role TEXT,
            slot_index INTEGER,
            value_text TEXT,
            value_json JSONB,
            value_image TEXT,
            is_completed BOOLEAN NOT NULL DEFAULT FALSE,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        """,
        """
        CREATE UNIQUE INDEX IF NOT EXISTS uq_group_document
        ON documents (form_id, group_name, work_date)
        WHERE form_type = 'group';
        """,
        """
        CREATE UNIQUE INDEX IF NOT EXISTS uq_individual_document
        ON documents (form_id, user_id, work_date)
        WHERE form_type = 'individual';
        """,
    ]

    with get_conn() as conn:
        with conn.cursor() as cur:
            for sql in sql_list:
                cur.execute(sql)
        conn.commit()
