import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent

class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "biblioteca-central-local-dev-key-change-in-production")
    # Sem configuração adicional, o sistema usa SQLite local e cria o arquivo automaticamente.
    # Para MySQL, basta definir DATABASE_URL no ambiente/.env; as tabelas são criadas no primeiro start.
    SQLALCHEMY_DATABASE_URI = os.getenv(
        "DATABASE_URL",
        f"sqlite:///{(BASE_DIR / 'instance' / 'biblioteca.db').as_posix()}"
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {"pool_pre_ping": True}
    WTF_CSRF_TIME_LIMIT = None
    MAX_CONTENT_LENGTH = 6 * 1024 * 1024
    DEBUG = os.getenv("FLASK_DEBUG", "0") == "1"
    AUTO_CREATE_DB = True
