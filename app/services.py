from datetime import date, datetime, timedelta
from sqlalchemy import or_
from .extensions import db
from .models import *

def audit(user, action, entity, entity_id=None, details=None):
    db.session.add(AuditLog(user_id=user.id if user else None, action=action,
                            entity=entity, entity_id=entity_id, details=details))

def refresh_notifications():
    today = date.today()
    candidates = []
    for c in Copy.query.filter(Copy.status != "Descartado").all():
        if c.next_evaluation:
            delta = (c.next_evaluation - today).days
            if delta < 0: candidates.append(("avaliacao_atrasada", f"Avaliação atrasada: {c.book.title} ({c.barcode})", c.id, None, c.next_evaluation))
            elif delta <= 14: candidates.append(("avaliacao_proxima", f"Avaliação próxima: {c.book.title} ({c.barcode})", c.id, None, c.next_evaluation))
        if c.condition in ("Ruim", "Necessita de restauração") or c.status == "Em manutenção":
            candidates.append(("manutencao", f"Verificar conservação: {c.book.title} ({c.barcode})", c.id, None, None))
        if c.useful_life_end and c.useful_life_end <= today + timedelta(days=60):
            candidates.append(("substituicao", f"Vida útil próxima do fim: {c.book.title} ({c.barcode})", c.id, None, c.useful_life_end))
    for loan in Loan.query.filter(Loan.returned_at.is_(None)).all():
        if loan.due_date <= today + timedelta(days=3):
            kind = "emprestimo_atrasado" if loan.due_date < today else "emprestimo_proximo"
            candidates.append((kind, f"Empréstimo {kind.replace('_',' ')}: {loan.copy.book.title} — {loan.borrower.name}", loan.copy_id, loan.id, loan.due_date))
    for kind, msg, copy_id, loan_id, due in candidates:
        exists = Notification.query.filter_by(kind=kind, copy_id=copy_id, loan_id=loan_id).first()
        if not exists:
            db.session.add(Notification(kind=kind, message=msg, copy_id=copy_id, loan_id=loan_id, due_date=due))
    db.session.commit()

def search_books(args):
    q = Book.query.filter_by(archived=(args.get("archived") == "1"))
    term = args.get("q", "").strip()
    if term:
        like = f"%{term}%"
        q = q.outerjoin(Copy).filter(or_(Book.title.ilike(like), Book.author.ilike(like),
            Book.internal_code.ilike(like), Book.isbn.ilike(like), Book.category.ilike(like),
            Copy.barcode.ilike(like), Copy.location.ilike(like)))
    for field in ("category",):
        if args.get(field): q = q.filter(getattr(Book, field) == args[field])
    if args.get("status"): q = q.join(Copy).filter(Copy.status == args["status"])
    if args.get("condition"): q = q.join(Copy).filter(Copy.condition == args["condition"])
    return q.distinct().order_by(Book.title)
