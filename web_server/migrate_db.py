from database import engine, Base
from sqlalchemy import text

def migrate():
    with engine.connect() as conn:
        # Check if columns exist, if not add them (SQLite specific)
        # SQLite doesn't support IF NOT EXISTS in ALTER TABLE, so we try and catch
        try:
            conn.execute(text("ALTER TABLE conversion_history ADD COLUMN source_file_path VARCHAR"))
            print("Added source_file_path column")
        except Exception as e:
            print(f"source_file_path column might already exist: {e}")
            
        try:
            conn.execute(text("ALTER TABLE conversion_history ADD COLUMN note VARCHAR"))
            print("Added note column")
        except Exception as e:
            print(f"note column might already exist: {e}")

if __name__ == "__main__":
    migrate()
