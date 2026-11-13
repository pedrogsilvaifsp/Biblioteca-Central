from datetime import datetime, date, timedelta
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from .extensions import db

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(140), nullable=False)
    username = db.Column(db.String(60), unique=True, nullable=False, index=True)
    email = db.Column(db.String(160), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="servidor")
    active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    def set_password(self, password): self.password_hash = generate_password_hash(password)
    def check_password(self, password): return check_password_hash(self.password_hash, password)
    @property
    def is_active(self): return self.active

class Book(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    internal_code = db.Column(db.String(40), unique=True, nullable=False, index=True)
    isbn = db.Column(db.String(20), unique=True, nullable=True)
    title = db.Column(db.String(220), nullable=False, index=True)
    author = db.Column(db.String(180), nullable=False, index=True)
    publisher = db.Column(db.String(140))
    category = db.Column(db.String(100), index=True)
    publication_year = db.Column(db.Integer)
    acquisition_date = db.Column(db.Date)
    shelf = db.Column(db.String(80))
    section = db.Column(db.String(80))
    notes = db.Column(db.Text)
    archived = db.Column(db.Boolean, default=False, nullable=False)
    created_by_id = db.Column(db.Integer, db.ForeignKey("user.id"))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    copies = db.relationship("Copy", backref="book", cascade="all, delete-orphan")
    @property
    def copy_count(self): return len(self.copies)

class Copy(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    book_id = db.Column(db.Integer, db.ForeignKey("book.id"), nullable=False, index=True)
    barcode = db.Column(db.String(60), unique=True, nullable=False, index=True)
    status = db.Column(db.String(30), default="Disponível", nullable=False, index=True)
    condition = db.Column(db.String(40), default="Bom", nullable=False)
    acquired_at = db.Column(db.Date)
    last_evaluation = db.Column(db.Date)
    next_evaluation = db.Column(db.Date, index=True)
    evaluation_interval_days = db.Column(db.Integer, default=180, nullable=False)
    useful_life_end = db.Column(db.Date)
    damage_notes = db.Column(db.Text)
    location = db.Column(db.String(120))
    version = db.Column(db.Integer, default=1, nullable=False)
    __table_args__ = (db.UniqueConstraint("book_id", "barcode", name="uq_copy_book_barcode"),)
    __mapper_args__ = {"version_id_col": version}

class Borrower(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(140), nullable=False, index=True)
    registration = db.Column(db.String(60), unique=True, nullable=True)
    email = db.Column(db.String(160))
    phone = db.Column(db.String(40))
    active = db.Column(db.Boolean, default=True, nullable=False)

class Loan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    copy_id = db.Column(db.Integer, db.ForeignKey("copy.id"), nullable=False, index=True)
    borrower_id = db.Column(db.Integer, db.ForeignKey("borrower.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    loan_date = db.Column(db.Date, nullable=False, default=date.today)
    due_date = db.Column(db.Date, nullable=False)
    returned_at = db.Column(db.DateTime)
    notes = db.Column(db.Text)
    copy = db.relationship("Copy")
    borrower = db.relationship("Borrower")
    operator = db.relationship("User")
    @property
    def overdue(self): return self.returned_at is None and self.due_date < date.today()

class Evaluation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    copy_id = db.Column(db.Integer, db.ForeignKey("copy.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    evaluated_at = db.Column(db.Date, nullable=False, default=date.today)
    condition = db.Column(db.String(40), nullable=False)
    notes = db.Column(db.Text)
    copy = db.relationship("Copy")
    operator = db.relationship("User")

class Recommendation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    copy_id = db.Column(db.Integer, db.ForeignKey("copy.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    reason = db.Column(db.Text, nullable=False)
    due_date = db.Column(db.Date)
    status = db.Column(db.String(20), default="Pendente", nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    copy = db.relationship("Copy")
    operator = db.relationship("User")

class Notification(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    kind = db.Column(db.String(40), nullable=False, index=True)
    message = db.Column(db.String(255), nullable=False)
    copy_id = db.Column(db.Integer, db.ForeignKey("copy.id"))
    loan_id = db.Column(db.Integer, db.ForeignKey("loan.id"))
    due_date = db.Column(db.Date)
    resolved_at = db.Column(db.DateTime)
    resolved_by_id = db.Column(db.Integer, db.ForeignKey("user.id"))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    copy = db.relationship("Copy")
    loan = db.relationship("Loan")
    resolved_by = db.relationship("User")
    reads = db.relationship("NotificationRead", cascade="all, delete-orphan", backref="notification")

class NotificationRead(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    notification_id = db.Column(db.Integer, db.ForeignKey("notification.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    read_at = db.Column(db.DateTime, default=datetime.utcnow)
    __table_args__ = (db.UniqueConstraint("notification_id", "user_id", name="uq_notification_user"),)

class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"))
    action = db.Column(db.String(60), nullable=False, index=True)
    entity = db.Column(db.String(60), nullable=False)
    entity_id = db.Column(db.Integer)
    details = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    user = db.relationship("User")


class Reservation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    book_id = db.Column(db.Integer, db.ForeignKey("book.id"), nullable=False, index=True)
    borrower_id = db.Column(db.Integer, db.ForeignKey("borrower.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    status = db.Column(db.String(20), default="Pendente", nullable=False)
    book = db.relationship("Book")
    borrower = db.relationship("Borrower")
    operator = db.relationship("User")


class WorkRequest(db.Model):
    """Solicitação interna criada e acompanhada pelos servidores."""
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(180), nullable=False, index=True)
    description = db.Column(db.Text, nullable=False)
    category = db.Column(db.String(40), nullable=False, default="Tarefa")
    priority = db.Column(db.String(20), nullable=False, default="Normal", index=True)
    status = db.Column(db.String(30), nullable=False, default="Aberta", index=True)
    due_date = db.Column(db.Date, index=True)
    created_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    assigned_to_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    creator = db.relationship("User", foreign_keys=[created_by_id])
    assignee = db.relationship("User", foreign_keys=[assigned_to_id])
    reads = db.relationship("WorkRequestRead", cascade="all, delete-orphan", backref="work_request")
    messages = db.relationship(
        "WorkRequestMessage",
        backref="work_request",
        cascade="all, delete-orphan",
        order_by="WorkRequestMessage.created_at",
    )


class WorkRequestMessage(db.Model):
    """Mensagem enviada na conversa de uma solicitação."""
    id = db.Column(db.Integer, primary_key=True)
    work_request_id = db.Column(db.Integer, db.ForeignKey("work_request.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    body = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)

    author = db.relationship("User")


class WorkRequestRead(db.Model):
    """Última visualização individual de cada solicitação por servidor."""
    id = db.Column(db.Integer, primary_key=True)
    work_request_id = db.Column(db.Integer, db.ForeignKey("work_request.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    last_viewed_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    __table_args__ = (db.UniqueConstraint("work_request_id", "user_id", name="uq_work_request_user_read"),)
