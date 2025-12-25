from sqlalchemy import Column, Integer, String, DateTime, JSON
from sqlalchemy.sql import func
from database import Base

class ConversionHistory(Base):
    __tablename__ = "conversion_history"

    id = Column(Integer, primary_key=True, index=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    original_filename = Column(String, index=True)
    mode = Column(String)  # allocation, delivery_note, assortment
    status = Column(String)  # processing, success, failed
    output_filename = Column(String, nullable=True)
    file_path = Column(String, nullable=True)  # Path to the output file
    source_file_path = Column(String, nullable=True) # Path to the uploaded source file
    note = Column(String, nullable=True) # User notes
    stats = Column(JSON, nullable=True)  # JSON field for statistics
    error_message = Column(String, nullable=True)

class SystemSetting(Base):
    __tablename__ = "system_settings"

    key = Column(String, primary_key=True, index=True)
    value = Column(String)
    description = Column(String, nullable=True)
