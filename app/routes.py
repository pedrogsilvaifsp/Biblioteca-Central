from datetime import date, datetime, timedelta
from io import BytesIO, StringIO
import csv, json, os, re, uuid
from pathlib import Path
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort, send_file, make_response, current_app
from flask_login import login_user, logout_user, login_required, current_user
from sqlalchemy import func, or_
from sqlalchemy.exc import IntegrityError
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from .extensions import db
from .models import *
from .services import audit, refresh_notifications, search_books

bp = Blueprint("main", __name__)

def admin_required():
    if not current_user.is_authenticated or current_user.role != "admin":
        abort(403)

@bp.route("/")
@login_required
def dashboard():
    refresh_notifications()
    stats = {
      "titles": Book.query.filter_by(archived=False).count(),
      "copies": Copy.query.join(Book).filter(Book.archived.is_(False)).count(),
      "available": Copy.query.filter_by(status="Disponível").count(),
      "loaned": Copy.query.filter_by(status="Emprestado").count(),
      "maintenance": Copy.query.filter_by(status="Em manutenção").count(),
      "poor": Copy.query.filter(Copy.condition.in_(["Ruim","Necessita de restauração"])).count(),
      "overdue": Loan.query.filter(Loan.returned_at.is_(None), Loan.due_date < date.today()).count(),
      "pending_replacements": Recommendation.query.filter_by(status="Pendente").count(),
      "notifications": Notification.query.filter_by(resolved_at=None).count()
    }
    categories = db.session.query(Book.category, func.count(Copy.id)).join(Copy).filter(Book.archived.is_(False)).group_by(Book.category).all()
    conditions = db.session.query(Copy.condition, func.count(Copy.id)).group_by(Copy.condition).all()
    recent = AuditLog.query.order_by(AuditLog.created_at.desc()).limit(8).all()
    return render_template("dashboard.html", stats=stats, categories=categories, conditions=conditions, recent=recent)

@bp.before_app_request
def require_initial_admin():
    # A tela de configuração é obrigatória até existir um Administrador Geral ativo.
    if request.endpoint in ("static", "main.setup_admin"):
        return None
    if User.query.filter_by(role="admin", active=True).count() == 0:
        return redirect(url_for("main.setup_admin"))


@bp.route("/setup-admin", methods=["GET", "POST"])
def setup_admin():
    if User.query.filter_by(role="admin", active=True).count() > 0:
        return redirect(url_for("main.login"))
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        if not all((name, username, email, password)):
            flash("Preencha todos os campos.", "danger")
        elif len(password) < 10:
            flash("A senha precisa ter pelo menos 10 caracteres.", "danger")
        else:
            user = User(name=name, username=username, email=email, role="admin", active=True)
            user.set_password(password)
            db.session.add(user)
            try:
                db.session.commit()
                login_user(user)
                flash("Administrador Geral criado. Bem-vindo ao sistema!", "success")
                return redirect(url_for("main.dashboard"))
            except IntegrityError:
                db.session.rollback()
                flash("Não foi possível criar a conta. Verifique se o login ou e-mail já está em uso.", "danger")
    return render_template("setup_admin.html")


@bp.route("/login", methods=["GET","POST"])
def login():
    if current_user.is_authenticated: return redirect(url_for("main.dashboard"))
    if request.method == "POST":
        user = User.query.filter_by(username=request.form.get("username","").strip()).first()
        if user and user.active and user.check_password(request.form.get("password","")):
            login_user(user)
            return redirect(url_for("main.dashboard"))
        flash("Usuário ou senha inválidos.", "danger")
    return render_template("login.html")

@bp.route("/logout", methods=["POST"])
@login_required
def logout():
    logout_user()
    return redirect(url_for("main.login"))

IMPORT_ALIASES = {
    "internal_code": ["codigo interno", "código interno", "codigo", "código", "tombo", "patrimonio", "patrimônio", "codigo do livro", "código do livro"],
    "title": ["titulo", "título", "nome", "nome do livro", "obra", "livro"],
    "author": ["autor", "autores", "escritor"],
    "isbn": ["isbn", "isbn 10", "isbn 13"],
    "publisher": ["editora", "publicadora"],
    "category": ["categoria", "categoria do livro", "assunto", "genero", "gênero"],
    "publication_year": ["ano", "ano de publicacao", "ano de publicação", "publicacao", "publicação"],
    "quantity": ["quantidade", "qtd", "exemplares", "quantidade de exemplares", "copias", "cópias"],
    "shelf": ["prateleira", "estante", "localizacao", "localização"],
    "section": ["secao", "seção", "setor"],
    "condition": ["conservacao", "conservação", "estado de conservacao", "estado de conservação", "estado"],
    "notes": ["observacoes", "observações", "notas", "descricao", "descrição"],
}

def _norm_header(value):
    import unicodedata
    value = unicodedata.normalize("NFKD", str(value or "").strip().lower())
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", value).strip()

def _read_import_file(file_storage):
    filename = (file_storage.filename or "").lower()
    raw = file_storage.read(5 * 1024 * 1024 + 1)
    if len(raw) > 5 * 1024 * 1024:
        raise ValueError("A planilha ultrapassa o limite de 5 MB.")
    if not raw:
        raise ValueError("O arquivo está vazio.")
    if filename.endswith(".xlsx"):
        wb = load_workbook(BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        try: headers = next(rows)
        except StopIteration: raise ValueError("A planilha não possui linhas.")
        records = [dict(zip(headers, row)) for row in rows if any(v is not None and str(v).strip() for v in row)]
        wb.close()
    elif filename.endswith(".csv"):
        try: content = raw.decode("utf-8-sig")
        except UnicodeDecodeError: content = raw.decode("latin-1")
        sample = content[:4096]
        try: dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
        except csv.Error: dialect = csv.excel
        reader = csv.DictReader(StringIO(content), dialect=dialect)
        records = [row for row in reader if any(str(v or "").strip() for v in row.values())]
    else:
        raise ValueError("Formato não aceito. Envie uma planilha .xlsx ou .csv.")
    if not records:
        raise ValueError("Não encontrei registros abaixo do cabeçalho.")
    if len(records) > 5000:
        raise ValueError("A planilha possui mais de 5.000 registros. Divida o arquivo em partes menores.")
    headers = list(records[0].keys())
    if len(headers) > 50:
        raise ValueError("A planilha possui colunas demais. Limite o arquivo a 50 colunas.")
    mapping = {}
    normalized = {_norm_header(h): h for h in headers if h is not None}
    for field, aliases in IMPORT_ALIASES.items():
        for alias in aliases:
            key = _norm_header(alias)
            if key in normalized:
                mapping[field] = normalized[key]
                break
    if "title" not in mapping or "author" not in mapping:
        raise ValueError("Não identifiquei as colunas obrigatórias Título e Autor. Confira o cabeçalho da planilha.")
    return records, mapping

def _preview_import(records, mapping):
    preview, errors = [], []
    seen_codes, seen_isbns = set(), set()
    stamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")
    for index, row in enumerate(records, start=2):
        def val(field):
            key = mapping.get(field)
            value = row.get(key) if key is not None else None
            return str(value).strip() if value is not None else ""
        title, author = val("title"), val("author")
        code = val("internal_code") or f"IMP-{stamp}-{index-1:05d}"
        isbn = val("isbn") or None
        try:
            qty = int(float(val("quantity") or 1))
            if qty < 1 or qty > 500: raise ValueError
        except ValueError:
            qty = 1; errors.append(f"Linha {index}: quantidade inválida (use de 1 a 500).")
        try:
            year = int(float(val("publication_year"))) if val("publication_year") else None
            if year and (year < 1400 or year > date.today().year + 1): raise ValueError
        except ValueError:
            year = None; errors.append(f"Linha {index}: ano de publicação inválido.")
        row_errors = []
        if not title: row_errors.append("Título ausente")
        if not author: row_errors.append("Autor ausente")
        if code in seen_codes or Book.query.filter_by(internal_code=code).first(): row_errors.append("Código interno duplicado")
        if isbn and (isbn in seen_isbns or Book.query.filter_by(isbn=isbn).first()): row_errors.append("ISBN já cadastrado")
        seen_codes.add(code)
        if isbn: seen_isbns.add(isbn)
        preview.append({
            "line": index, "internal_code": code, "title": title, "author": author, "isbn": isbn,
            "publisher": val("publisher"), "category": val("category"), "publication_year": year,
            "quantity": qty, "shelf": val("shelf"), "section": val("section"),
            "condition": val("condition") or "Bom", "notes": val("notes"), "errors": row_errors,
        })
    return preview, errors

@bp.route("/acervo/modelo-planilha")
@login_required
def import_template():
    admin_required()
    wb = Workbook()
    ws = wb.active
    ws.title = "Acervo"
    ws.append(["Código interno", "Título", "Autor", "ISBN", "Editora", "Categoria", "Ano", "Quantidade", "Estante", "Seção", "Conservação", "Observações"])
    ws.append(["BIB-0001", "O Pequeno Príncipe", "Antoine de Saint-Exupéry", "9788595081512", "HarperCollins", "Literatura", 2018, 3, "Estante A", "Prateleira 02", "Bom", "Exemplo: remover esta linha antes de importar"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    out = BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name="modelo_importacao_acervo.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@bp.route("/acervo/importar", methods=["GET", "POST"])
@login_required
def import_acervo():
    admin_required()
    if request.method == "POST":
        action = request.form.get("action")
        token = request.form.get("token", "")
        preview_dir = Path(current_app.instance_path) / "import_previews"
        preview_dir.mkdir(parents=True, exist_ok=True)
        if action == "confirm":
            if not re.fullmatch(r"[a-f0-9]{32}", token): abort(400)
            path = preview_dir / f"{token}.json"
            if not path.is_file():
                flash("Prévia expirada. Envie a planilha novamente.", "warning")
                return redirect(url_for("main.import_acervo"))
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
                if payload.get("user_id") != current_user.id or datetime.utcnow().timestamp() - payload.get("created_at", 0) > 1800:
                    path.unlink(missing_ok=True)
                    raise ValueError("Prévia expirada ou não pertence à sua sessão. Envie a planilha novamente.")
                rows = payload.get("rows", [])
                if not rows or any(row.get("errors") for row in rows) or payload.get("errors"):
                    raise ValueError("A prévia contém erros. Corrija a planilha antes de importar.")
                for row in rows:
                    book = Book(internal_code=row["internal_code"], isbn=row["isbn"], title=row["title"], author=row["author"],
                        publisher=row["publisher"] or None, category=row["category"] or None, publication_year=row["publication_year"],
                        shelf=row["shelf"] or None, section=row["section"] or None, notes=row["notes"] or None, created_by_id=current_user.id)
                    db.session.add(book); db.session.flush()
                    for number in range(1, row["quantity"] + 1):
                        barcode = f"{book.internal_code}-{number:03d}"
                        db.session.add(Copy(book_id=book.id, barcode=barcode, status="Disponível",
                            condition=row["condition"], location=" / ".join(filter(None, [row["section"], row["shelf"]])) or None))
                    audit(current_user, "IMPORTAÇÃO DE ACERVO", "Book", book.id, f"{row['quantity']} exemplar(es)")
                db.session.commit()
                path.unlink(missing_ok=True)
                flash(f"Importação concluída: {len(rows)} título(s) cadastrado(s).", "success")
                return redirect(url_for("main.books"))
            except Exception:
                db.session.rollback()
                flash("Importação cancelada. Nenhum registro foi gravado. Confira códigos internos, ISBNs e os dados da planilha e tente novamente.", "danger")
                return redirect(url_for("main.import_acervo"))
        upload = request.files.get("spreadsheet")
        if not upload or not upload.filename:
            flash("Selecione uma planilha .xlsx ou .csv.", "warning")
            return redirect(url_for("main.import_acervo"))
        try:
            # Remove prévias antigas para não acumular arquivos temporários no servidor.
            for old_file in preview_dir.glob("*.json"):
                try:
                    if datetime.utcnow().timestamp() - old_file.stat().st_mtime > 1800:
                        old_file.unlink(missing_ok=True)
                except OSError:
                    pass
            records, mapping = _read_import_file(upload)
            rows, errors = _preview_import(records, mapping)
            token = uuid.uuid4().hex
            preview_path = preview_dir / f"{token}.json"
            preview_path.write_text(json.dumps({"rows": rows, "errors": errors, "user_id": current_user.id, "created_at": datetime.utcnow().timestamp()}, ensure_ascii=False), encoding="utf-8")
            try: os.chmod(preview_path, 0o600)
            except OSError: pass
            return render_template("import_acervo.html", rows=rows, errors=errors, token=token, mapping=mapping)
        except Exception as exc:
            flash(str(exc), "danger")
    return render_template("import_acervo.html", rows=None, errors=[], token=None, mapping={})

@bp.route("/books")
@login_required
def books():
    rows = search_books(request.args).all()
    statuses = [row[0] for row in Copy.query.with_entities(Copy.status).distinct().order_by(Copy.status).all()]
    return render_template("books.html", books=rows, statuses=statuses)

@bp.route("/books/new", methods=["GET","POST"])
@login_required
def book_new():
    if request.method == "POST":
        try:
            book = Book(internal_code=request.form["internal_code"].strip(), isbn=request.form.get("isbn") or None,
                title=request.form["title"].strip(), author=request.form["author"].strip(),
                publisher=request.form.get("publisher"), category=request.form.get("category"),
                publication_year=int(request.form["publication_year"]) if request.form.get("publication_year") else None,
                acquisition_date=date.fromisoformat(request.form["acquisition_date"]) if request.form.get("acquisition_date") else None,
                shelf=request.form.get("shelf"), section=request.form.get("section"), notes=request.form.get("notes"),
                created_by_id=current_user.id)
            qty = max(1, min(500, int(request.form.get("quantity",1))))
            db.session.add(book); db.session.flush()
            for i in range(qty):
                barcode = f"{book.internal_code}-{i+1:03d}"
                db.session.add(Copy(book_id=book.id, barcode=barcode, acquired_at=book.acquisition_date,
                    location=" / ".join(filter(None,[book.section,book.shelf]))))
            audit(current_user, "CADASTRO", "Book", book.id, f"{qty} exemplar(es)")
            db.session.commit()
            flash("Livro cadastrado com sucesso.", "success")
            return redirect(url_for("main.books"))
        except (ValueError, KeyError, IntegrityError):
            db.session.rollback(); flash("Não foi possível cadastrar. Verifique os campos e se o código/ISBN já existe.", "danger")
    return render_template("book_form.html")


@bp.route("/books/<int:book_id>/edit", methods=["GET","POST"])
@login_required
def book_edit(book_id):
    book = db.get_or_404(Book, book_id)
    if request.method == "POST":
        try:
            book.title=request.form["title"].strip(); book.author=request.form["author"].strip()
            book.publisher=request.form.get("publisher"); book.category=request.form.get("category")
            book.isbn=request.form.get("isbn") or None
            book.publication_year=int(request.form["publication_year"]) if request.form.get("publication_year") else None
            book.acquisition_date=date.fromisoformat(request.form["acquisition_date"]) if request.form.get("acquisition_date") else None
            book.shelf=request.form.get("shelf"); book.section=request.form.get("section"); book.notes=request.form.get("notes")
            audit(current_user,"EDIÇÃO","Book",book.id); db.session.commit()
            flash("Cadastro atualizado.","success"); return redirect(url_for("main.books"))
        except (ValueError, IntegrityError):
            db.session.rollback(); flash("Não foi possível salvar. Verifique os dados e o ISBN.","danger")
    return render_template("book_edit.html", book=book)

@bp.route("/reservations", methods=["GET","POST"])
@login_required
def reservations():
    if request.method == "POST":
        borrower=db.session.get(Borrower,int(request.form.get("borrower_id",0) or 0))
        book=db.session.get(Book,int(request.form.get("book_id",0) or 0))
        if not borrower or not book or book.archived:
            flash("Leitor ou título inválido.","danger")
        else:
            r=Reservation(book_id=book.id,borrower_id=borrower.id,user_id=current_user.id)
            db.session.add(r); db.session.flush(); audit(current_user,"RESERVA","Reservation",r.id); db.session.commit()
            flash("Reserva registrada.","success")
        return redirect(url_for("main.reservations"))
    return render_template("reservations.html", rows=Reservation.query.order_by(Reservation.created_at.desc()).all(),
        books=Book.query.filter_by(archived=False).order_by(Book.title).all(), borrowers=Borrower.query.filter_by(active=True).all())

@bp.route("/reservations/<int:rid>/cancel", methods=["POST"])
@login_required
def reservation_cancel(rid):
    r=db.get_or_404(Reservation,rid)
    if r.status == "Pendente":
        r.status="Cancelada"; audit(current_user,"CANCELAMENTO DE RESERVA","Reservation",r.id); db.session.commit()
    return redirect(url_for("main.reservations"))

@bp.route("/books/<int:book_id>/copies/new", methods=["POST"])
@login_required
def copy_new(book_id):
    book=db.get_or_404(Book,book_id)
    if book.archived:
        flash("Reative o título antes de adicionar exemplares.","warning")
        return redirect(url_for("main.books"))
    try:
        quantity=max(1,min(100,int(request.form.get("quantity",1))))
        existing=len(book.copies)
        for offset in range(1,quantity+1):
            barcode=f"{book.internal_code}-{existing+offset:03d}"
            db.session.add(Copy(book_id=book.id,barcode=barcode,location=" / ".join(filter(None,[book.section,book.shelf])) or None))
        audit(current_user,"ADIÇÃO DE EXEMPLARES","Book",book.id,f"{quantity} exemplar(es)")
        db.session.commit(); flash(f"{quantity} exemplar(es) adicionado(s).","success")
    except (ValueError,IntegrityError):
        db.session.rollback(); flash("Não foi possível adicionar exemplares. Verifique a quantidade e os códigos.","danger")
    return redirect(url_for("main.books"))


@bp.route("/books/<int:book_id>/unarchive", methods=["POST"])
@login_required
def book_unarchive(book_id):
    book=db.get_or_404(Book,book_id); book.archived=False
    audit(current_user,"REATIVAÇÃO","Book",book.id); db.session.commit()
    flash("Título reativado no acervo.","success")
    return redirect(url_for("main.books"))


@bp.route("/books/<int:book_id>/archive", methods=["POST"])
@login_required
def book_archive(book_id):
    book = db.get_or_404(Book, book_id)
    if any(c.status == "Emprestado" for c in book.copies):
        flash("Não é possível arquivar um título com exemplar emprestado.", "warning")
    else:
        book.archived = True; audit(current_user,"ARQUIVAMENTO","Book",book.id); db.session.commit()
        flash("Título arquivado.", "success")
    return redirect(url_for("main.books"))

@bp.route("/copies/<int:copy_id>", methods=["GET","POST"])
@login_required
def copy_detail(copy_id):
    copy = db.get_or_404(Copy, copy_id)
    if request.method == "POST":
        requested_status = request.form.get("status", copy.status)
        active_loan = Loan.query.filter_by(copy_id=copy.id, returned_at=None).first()
        if requested_status == "Emprestado" and not active_loan:
            flash("O status Emprestado só pode ser definido pelo registro de empréstimo.", "warning")
            return redirect(url_for("main.copy_detail", copy_id=copy.id))
        if requested_status == "Disponível" and active_loan:
            flash("Registre a devolução antes de liberar este exemplar.", "warning")
            return redirect(url_for("main.copy_detail", copy_id=copy.id))
        copy.status = requested_status
        copy.condition = request.form.get("condition", copy.condition)
        copy.location = request.form.get("location","").strip()
        copy.damage_notes = request.form.get("damage_notes")
        copy.version += 1
        audit(current_user,"ATUALIZAÇÃO","Copy",copy.id,"Status/conservação/localização")
        db.session.commit(); flash("Exemplar atualizado.", "success")
        return redirect(url_for("main.copy_detail", copy_id=copy.id))
    return render_template("copy_detail.html", copy=copy)

@bp.route("/borrowers", methods=["GET","POST"])
@login_required
def borrowers():
    if request.method == "POST":
        b = Borrower(name=request.form["name"].strip(), registration=request.form.get("registration") or None,
                     email=request.form.get("email"), phone=request.form.get("phone"))
        db.session.add(b)
        try:
            db.session.flush(); audit(current_user,"CADASTRO","Borrower",b.id); db.session.commit()
            flash("Leitor cadastrado.", "success")
        except IntegrityError:
            db.session.rollback(); flash("Matrícula já cadastrada.", "danger")
    return render_template("borrowers.html", borrowers=Borrower.query.order_by(Borrower.name).all())

@bp.route("/loans", methods=["GET","POST"])
@login_required
def loans():
    if request.method == "POST":
        copy = Copy.query.filter_by(barcode=request.form.get("barcode","").strip()).with_for_update().first()
        borrower = db.session.get(Borrower, int(request.form.get("borrower_id",0) or 0))
        if not copy or copy.status != "Disponível" or not borrower or not borrower.active:
            flash("Exemplar indisponível ou leitor inválido.", "danger")
        else:
            loan = Loan(copy_id=copy.id, borrower_id=borrower.id, user_id=current_user.id,
                loan_date=date.today(), due_date=date.fromisoformat(request.form["due_date"]))
            copy.status = "Emprestado"
            db.session.add(loan); db.session.flush(); audit(current_user,"EMPRÉSTIMO","Loan",loan.id,copy.barcode)
            db.session.commit(); flash("Empréstimo registrado.", "success")
        return redirect(url_for("main.loans"))
    rows = Loan.query.order_by(Loan.loan_date.desc()).limit(200).all()
    return render_template("loans.html", loans=rows, borrowers=Borrower.query.filter_by(active=True).order_by(Borrower.name).all())

@bp.route("/loans/<int:loan_id>/return", methods=["POST"])
@login_required
def loan_return(loan_id):
    loan = db.get_or_404(Loan, loan_id)
    if loan.returned_at:
        flash("Este empréstimo já foi devolvido.", "info")
    else:
        loan.returned_at = datetime.utcnow()
        if loan.copy.status == "Emprestado": loan.copy.status = "Disponível"
        audit(current_user,"DEVOLUÇÃO","Loan",loan.id,loan.copy.barcode)
        db.session.commit(); flash("Devolução registrada.", "success")
    return redirect(url_for("main.loans"))

@bp.route("/evaluations", methods=["GET","POST"])
@login_required
def evaluations():
    if request.method == "POST":
        copy = db.get_or_404(Copy, int(request.form["copy_id"]))
        d = date.fromisoformat(request.form.get("evaluated_at") or date.today().isoformat())
        condition = request.form["condition"]
        ev = Evaluation(copy_id=copy.id,user_id=current_user.id,evaluated_at=d,condition=condition,notes=request.form.get("notes"))
        copy.last_evaluation=d; copy.next_evaluation=d+timedelta(days=max(1,int(request.form.get("interval",180))))
        copy.evaluation_interval_days=max(1,int(request.form.get("interval",180))); copy.condition=condition
        if condition == "Necessita de restauração": copy.status="Em manutenção"
        db.session.add(ev); audit(current_user,"AVALIAÇÃO","Evaluation",None,copy.barcode); db.session.commit()
        flash("Avaliação registrada e próxima data calculada.", "success")
        return redirect(url_for("main.evaluations"))
    due = Copy.query.filter(Copy.next_evaluation <= date.today()+timedelta(days=14)).order_by(Copy.next_evaluation).all()
    history = Evaluation.query.order_by(Evaluation.evaluated_at.desc()).limit(100).all()
    return render_template("evaluations.html", copies=Copy.query.filter(Copy.status!="Descartado").all(), due=due, history=history)

@bp.route("/recommendations", methods=["GET","POST"])
@login_required
def recommendations():
    if request.method == "POST":
        r = Recommendation(copy_id=int(request.form["copy_id"]),user_id=current_user.id,
            reason=request.form["reason"],due_date=date.fromisoformat(request.form["due_date"]) if request.form.get("due_date") else None)
        db.session.add(r); db.session.flush(); audit(current_user,"RECOMENDAÇÃO","Recommendation",r.id); db.session.commit()
        flash("Recomendação registrada.", "success"); return redirect(url_for("main.recommendations"))
    rows=Recommendation.query.order_by(Recommendation.created_at.desc()).all()
    return render_template("recommendations.html", rows=rows, copies=Copy.query.all())

@bp.route("/recommendations/<int:rid>/resolve", methods=["POST"])
@login_required
def recommendation_resolve(rid):
    r=db.get_or_404(Recommendation,rid); r.status="Concluída"; audit(current_user,"RESOLUÇÃO","Recommendation",r.id); db.session.commit()
    return redirect(url_for("main.recommendations"))

REQUEST_CATEGORIES = ["Tarefa", "Manutenção", "Aquisição", "Organização", "Tecnologia", "Outro"]
REQUEST_PRIORITIES = ["Baixa", "Normal", "Alta", "Urgente"]
REQUEST_STATUSES = ["Aberta", "Em andamento", "Aguardando", "Concluída", "Cancelada"]


@bp.route("/solicitacoes", methods=["GET", "POST"])
@login_required
def work_requests():
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        description = request.form.get("description", "").strip()
        category = request.form.get("category", "Tarefa")
        priority = request.form.get("priority", "Normal")
        due_raw = request.form.get("due_date", "").strip()
        assigned_raw = request.form.get("assigned_to_id", "").strip()

        if not title or not description:
            flash("Informe o título e descreva o que precisa ser feito.", "danger")
            return redirect(url_for("main.work_requests"))
        if category not in REQUEST_CATEGORIES or priority not in REQUEST_PRIORITIES:
            abort(400)
        try:
            due_date = date.fromisoformat(due_raw) if due_raw else None
            assigned_to_id = int(assigned_raw) if assigned_raw else None
        except ValueError:
            flash("Prazo ou responsável inválido.", "danger")
            return redirect(url_for("main.work_requests"))

        assignee = db.session.get(User, assigned_to_id) if assigned_to_id else None
        if assigned_to_id and (not assignee or not assignee.active):
            flash("O responsável selecionado não está disponível.", "danger")
            return redirect(url_for("main.work_requests"))

        item = WorkRequest(
            title=title, description=description, category=category, priority=priority,
            due_date=due_date, created_by_id=current_user.id, assigned_to_id=assigned_to_id,
        )
        db.session.add(item)
        db.session.flush()
        audit(current_user, "NOVA SOLICITAÇÃO", "WorkRequest", item.id, title)
        db.session.commit()
        flash("Solicitação criada. A equipe já pode acompanhar e conversar nela.", "success")
        return redirect(url_for("main.work_request_detail", request_id=item.id))

    query = WorkRequest.query
    status_filter = request.args.get("status", "").strip()
    term = request.args.get("q", "").strip()
    if status_filter in REQUEST_STATUSES:
        query = query.filter_by(status=status_filter)
    if term:
        like = f"%{term}%"
        query = query.filter(or_(WorkRequest.title.ilike(like), WorkRequest.description.ilike(like)))
    rows = query.order_by(WorkRequest.updated_at.desc(), WorkRequest.created_at.desc()).all()
    counts = {status: WorkRequest.query.filter_by(status=status).count() for status in REQUEST_STATUSES}
    return render_template(
        "work_requests.html", rows=rows, counts=counts, status_filter=status_filter, term=term,
        categories=REQUEST_CATEGORIES, priorities=REQUEST_PRIORITIES,
        users=User.query.filter_by(active=True).order_by(User.name).all(),
    )


@bp.route("/solicitacoes/<int:request_id>", methods=["GET"])
@login_required
def work_request_detail(request_id):
    item = db.get_or_404(WorkRequest, request_id)
    read = WorkRequestRead.query.filter_by(work_request_id=item.id, user_id=current_user.id).first()
    if read:
        read.last_viewed_at = datetime.utcnow()
    else:
        db.session.add(WorkRequestRead(work_request_id=item.id, user_id=current_user.id, last_viewed_at=datetime.utcnow()))
    db.session.commit()
    return render_template(
        "work_request_detail.html", item=item, statuses=REQUEST_STATUSES,
        priorities=REQUEST_PRIORITIES, users=User.query.filter_by(active=True).order_by(User.name).all(),
    )


@bp.route("/solicitacoes/<int:request_id>/mensagens", methods=["POST"])
@login_required
def work_request_message(request_id):
    item = db.get_or_404(WorkRequest, request_id)
    body = request.form.get("body", "").strip()
    if not body:
        flash("Digite uma mensagem antes de enviar.", "warning")
    elif len(body) > 10000:
        flash("A mensagem deve ter no máximo 10.000 caracteres.", "danger")
    else:
        message = WorkRequestMessage(work_request_id=item.id, user_id=current_user.id, body=body)
        db.session.add(message)
        item.updated_at = datetime.utcnow()
        audit(current_user, "MENSAGEM EM SOLICITAÇÃO", "WorkRequest", item.id)
        db.session.commit()
        flash("Mensagem enviada para a equipe.", "success")
    return redirect(url_for("main.work_request_detail", request_id=item.id) + "#conversa")


@bp.route("/solicitacoes/<int:request_id>/atualizar", methods=["POST"])
@login_required
def work_request_update(request_id):
    item = db.get_or_404(WorkRequest, request_id)
    status = request.form.get("status", item.status)
    priority = request.form.get("priority", item.priority)
    assigned_raw = request.form.get("assigned_to_id", "").strip()
    due_raw = request.form.get("due_date", "").strip()
    if status not in REQUEST_STATUSES or priority not in REQUEST_PRIORITIES:
        abort(400)
    try:
        assigned_to_id = int(assigned_raw) if assigned_raw else None
        due_date = date.fromisoformat(due_raw) if due_raw else None
    except ValueError:
        flash("Prazo ou responsável inválido.", "danger")
        return redirect(url_for("main.work_request_detail", request_id=item.id))
    assignee = db.session.get(User, assigned_to_id) if assigned_to_id else None
    if assigned_to_id and (not assignee or not assignee.active):
        flash("O responsável selecionado não está disponível.", "danger")
        return redirect(url_for("main.work_request_detail", request_id=item.id))

    title = request.form.get("title", item.title).strip()
    description = request.form.get("description", item.description).strip()
    category = request.form.get("category", item.category)
    if not title or not description or category not in REQUEST_CATEGORIES:
        flash("Título, descrição ou categoria inválidos.", "danger")
        return redirect(url_for("main.work_request_detail", request_id=item.id))
    item.title = title
    item.description = description
    item.category = category
    item.status = status
    item.priority = priority
    item.assigned_to_id = assigned_to_id
    item.due_date = due_date
    item.updated_at = datetime.utcnow()
    audit(current_user, "ATUALIZAÇÃO DE SOLICITAÇÃO", "WorkRequest", item.id, status)
    db.session.commit()
    flash("Solicitação atualizada.", "success")
    return redirect(url_for("main.work_request_detail", request_id=item.id))


@bp.route("/notifications")
@login_required
def notifications():
    refresh_notifications()
    rows=Notification.query.order_by(Notification.resolved_at.is_(None).desc(),Notification.created_at.desc()).all()
    return render_template("notifications.html", rows=rows)

@bp.route("/notifications/<int:nid>/read", methods=["POST"])
@login_required
def notification_read(nid):
    n=db.get_or_404(Notification,nid)
    if not NotificationRead.query.filter_by(notification_id=n.id,user_id=current_user.id).first():
        db.session.add(NotificationRead(notification_id=n.id,user_id=current_user.id))
    db.session.commit(); return redirect(url_for("main.notifications"))

@bp.route("/notifications/<int:nid>/resolve", methods=["POST"])
@login_required
def notification_resolve(nid):
    n=db.get_or_404(Notification,nid)
    if not n.resolved_at:
        n.resolved_at=datetime.utcnow(); n.resolved_by_id=current_user.id
        audit(current_user,"NOTIFICAÇÃO RESOLVIDA","Notification",n.id); db.session.commit()
    return redirect(url_for("main.notifications"))

@bp.route("/users", methods=["GET","POST"])
@login_required
def users():
    admin_required()
    if request.method=="POST":
        password=request.form["password"]
        if len(password) < 10:
            flash("A senha precisa ter pelo menos 10 caracteres.","danger")
            return redirect(url_for("main.users"))
        u=User(name=request.form["name"],username=request.form["username"].strip(),email=request.form["email"].strip(),
               role=request.form.get("role","servidor"),active=True)
        u.set_password(password); db.session.add(u)
        try:
            db.session.flush(); audit(current_user,"CADASTRO","User",u.id); db.session.commit(); flash("Usuário criado.","success")
        except IntegrityError:
            db.session.rollback(); flash("Login ou e-mail já utilizado.","danger")
    return render_template("users.html", users=User.query.order_by(User.name).all())

@bp.route("/users/<int:uid>/toggle", methods=["POST"])
@login_required
def user_toggle(uid):
    admin_required(); u=db.get_or_404(User,uid)
    if u.id==current_user.id: flash("Não é possível desativar a própria conta.","warning")
    else:
        u.active=not u.active; audit(current_user,"STATUS USUÁRIO","User",u.id); db.session.commit()
    return redirect(url_for("main.users"))



# ---------------------------------------------------------------------------
# Operações de edição e exclusão segura (CRUD complementar)
# ---------------------------------------------------------------------------
def _record_edit(entity, record, fields, title):
    return render_template("record_edit.html", entity=entity, record=record, fields=fields, page_title=title)


@bp.route("/books/<int:book_id>/delete", methods=["POST"])
@login_required
def book_delete(book_id):
    book = db.get_or_404(Book, book_id)
    copy_ids = [c.id for c in book.copies]
    has_history = bool(copy_ids and (Loan.query.filter(Loan.copy_id.in_(copy_ids)).first() or Evaluation.query.filter(Evaluation.copy_id.in_(copy_ids)).first() or Recommendation.query.filter(Recommendation.copy_id.in_(copy_ids)).first() or Notification.query.filter(Notification.copy_id.in_(copy_ids)).first()))
    has_reservations = Reservation.query.filter_by(book_id=book.id).first() is not None
    if has_history or has_reservations:
        flash("Este título possui movimentações ou reservas vinculadas. Use Arquivar para preservá-lo no histórico.", "warning")
    else:
        audit(current_user, "EXCLUSÃO", "Book", book.id, book.internal_code)
        db.session.delete(book); db.session.commit()
        flash("Título excluído do acervo.", "success")
    return redirect(url_for("main.books"))


@bp.route("/copies/<int:copy_id>/delete", methods=["POST"])
@login_required
def copy_delete(copy_id):
    copy = db.get_or_404(Copy, copy_id)
    linked = (Loan.query.filter_by(copy_id=copy.id).first() or Evaluation.query.filter_by(copy_id=copy.id).first() or Recommendation.query.filter_by(copy_id=copy.id).first() or Notification.query.filter_by(copy_id=copy.id).first())
    if linked or copy.status == "Emprestado":
        flash("Não é possível excluir exemplar com histórico. Altere seu status para Descartado ou mantenha o registro.", "warning")
    else:
        book_id = copy.book_id
        audit(current_user, "EXCLUSÃO", "Copy", copy.id, copy.barcode)
        db.session.delete(copy); db.session.commit()
        flash("Exemplar excluído.", "success")
        return redirect(url_for("main.books"))
    return redirect(url_for("main.copy_detail", copy_id=copy.id))


@bp.route("/borrowers/<int:bid>/edit", methods=["GET", "POST"])
@login_required
def borrower_edit(bid):
    b = db.get_or_404(Borrower, bid)
    if request.method == "POST":
        b.name=request.form.get("name", "").strip(); b.registration=request.form.get("registration", "").strip() or None
        b.email=request.form.get("email", "").strip() or None; b.phone=request.form.get("phone", "").strip() or None
        b.active = request.form.get("active") == "on"
        if not b.name:
            flash("Informe o nome do leitor.", "danger")
        else:
            try:
                audit(current_user, "EDIÇÃO", "Borrower", b.id); db.session.commit()
                flash("Leitor atualizado.", "success"); return redirect(url_for("main.borrowers"))
            except IntegrityError:
                db.session.rollback(); flash("Esta matrícula já está cadastrada.", "danger")
    fields=[("name","Nome", "text", b.name), ("registration","Matrícula", "text", b.registration or ""), ("email","E-mail", "email", b.email or ""), ("phone","Telefone", "text", b.phone or ""),("active","Leitor ativo", "checkbox", b.active)]
    return _record_edit("Leitor", b, fields, "Editar leitor")


@bp.route("/borrowers/<int:bid>/delete", methods=["POST"])
@login_required
def borrower_delete(bid):
    b=db.get_or_404(Borrower,bid)
    if Loan.query.filter_by(borrower_id=b.id).first() or Reservation.query.filter_by(borrower_id=b.id).first():
        b.active=False; audit(current_user,"DESATIVAÇÃO","Borrower",b.id); db.session.commit()
        flash("O leitor possui histórico. Foi desativado, sem apagar seus registros.","warning")
    else:
        audit(current_user,"EXCLUSÃO","Borrower",b.id); db.session.delete(b); db.session.commit(); flash("Leitor excluído.","success")
    return redirect(url_for("main.borrowers"))


@bp.route("/loans/<int:loan_id>/edit", methods=["GET", "POST"])
@login_required
def loan_edit(loan_id):
    loan=db.get_or_404(Loan,loan_id)
    if request.method=="POST":
        try:
            due=date.fromisoformat(request.form.get("due_date", ""))
            loan.due_date=due; loan.notes=request.form.get("notes", "").strip() or None
            audit(current_user,"EDIÇÃO","Loan",loan.id,"Prazo/observações"); db.session.commit()
            flash("Empréstimo atualizado.","success"); return redirect(url_for("main.loans"))
        except ValueError:
            flash("Informe uma data de vencimento válida.","danger")
    fields=[("due_date","Data de vencimento","date",loan.due_date.isoformat()),("notes","Observações","textarea",loan.notes or "")]
    return _record_edit("Empréstimo",loan,fields,"Editar empréstimo")


@bp.route("/loans/<int:loan_id>/delete", methods=["POST"])
@login_required
def loan_delete(loan_id):
    loan=db.get_or_404(Loan,loan_id)
    if loan.returned_at:
        flash("Empréstimos devolvidos fazem parte do histórico permanente e não podem ser apagados.","warning")
    else:
        copy=loan.copy
        if copy.status=="Emprestado": copy.status="Disponível"
        audit(current_user,"CANCELAMENTO/EXCLUSÃO","Loan",loan.id,copy.barcode)
        db.session.delete(loan); db.session.commit(); flash("Empréstimo cancelado e exemplar liberado.","success")
    return redirect(url_for("main.loans"))


@bp.route("/evaluations/<int:eid>/edit", methods=["GET", "POST"])
@login_required
def evaluation_edit(eid):
    ev=db.get_or_404(Evaluation,eid)
    if request.method=="POST":
        try:
            ev.evaluated_at=date.fromisoformat(request.form.get("evaluated_at", ""))
            ev.condition=request.form.get("condition", ev.condition); ev.notes=request.form.get("notes", "").strip() or None
            interval=max(1,min(3650,int(request.form.get("interval",180))))
            copy=ev.copy; copy.condition=ev.condition; copy.last_evaluation=ev.evaluated_at
            copy.evaluation_interval_days=interval; copy.next_evaluation=ev.evaluated_at+timedelta(days=interval)
            if ev.condition=="Necessita de restauração": copy.status="Em manutenção"
            audit(current_user,"EDIÇÃO","Evaluation",ev.id); db.session.commit()
            flash("Avaliação atualizada.","success"); return redirect(url_for("main.evaluations"))
        except ValueError:
            flash("Confira a data e a periodicidade.","danger")
    fields=[("evaluated_at","Data da avaliação","date",ev.evaluated_at.isoformat()),("condition","Conservação","select",ev.condition),("interval","Periodicidade (dias)","number",ev.copy.evaluation_interval_days),("notes","Observações","textarea",ev.notes or "")]
    return _record_edit("Avaliação",ev,fields,"Editar avaliação")


@bp.route("/evaluations/<int:eid>/delete", methods=["POST"])
@login_required
def evaluation_delete(eid):
    ev=db.get_or_404(Evaluation,eid); copy=ev.copy
    audit(current_user,"EXCLUSÃO","Evaluation",ev.id); db.session.delete(ev); db.session.flush()
    latest=Evaluation.query.filter_by(copy_id=copy.id).order_by(Evaluation.evaluated_at.desc(),Evaluation.id.desc()).first()
    if latest:
        copy.condition=latest.condition; copy.last_evaluation=latest.evaluated_at
        copy.next_evaluation=latest.evaluated_at+timedelta(days=max(1,copy.evaluation_interval_days or 180))
    else:
        copy.last_evaluation=None; copy.next_evaluation=None
    db.session.commit(); flash("Avaliação excluída. O resumo do exemplar foi recalculado.","success")
    return redirect(url_for("main.evaluations"))


@bp.route("/recommendations/<int:rid>/edit", methods=["GET", "POST"])
@login_required
def recommendation_edit(rid):
    r=db.get_or_404(Recommendation,rid)
    if request.method=="POST":
        r.reason=request.form.get("reason", "").strip()
        try: r.due_date=date.fromisoformat(request.form["due_date"]) if request.form.get("due_date") else None
        except ValueError: flash("Prazo inválido.","danger"); return redirect(url_for("main.recommendation_edit",rid=r.id))
        r.status=request.form.get("status",r.status)
        if r.status not in ["Pendente","Concluída","Cancelada"]: abort(400)
        audit(current_user,"EDIÇÃO","Recommendation",r.id); db.session.commit(); flash("Recomendação atualizada.","success")
        return redirect(url_for("main.recommendations"))
    fields=[("reason","Justificativa","textarea",r.reason),("due_date","Prazo","date",r.due_date.isoformat() if r.due_date else ""),("status","Status","select",r.status)]
    return _record_edit("Recomendação",r,fields,"Editar recomendação")


@bp.route("/recommendations/<int:rid>/delete", methods=["POST"])
@login_required
def recommendation_delete(rid):
    r=db.get_or_404(Recommendation,rid); audit(current_user,"EXCLUSÃO","Recommendation",r.id); db.session.delete(r); db.session.commit()
    flash("Recomendação excluída.","success"); return redirect(url_for("main.recommendations"))


@bp.route("/reservations/<int:rid>/edit", methods=["GET", "POST"])
@login_required
def reservation_edit(rid):
    r=db.get_or_404(Reservation,rid)
    if request.method=="POST":
        r.status=request.form.get("status",r.status)
        book=db.session.get(Book,int(request.form.get("book_id",0) or 0))
        borrower=db.session.get(Borrower,int(request.form.get("borrower_id",0) or 0))
        if r.status not in ["Pendente","Atendida","Cancelada"]: abort(400)
        if not book or not borrower:
            flash("Selecione um título e um leitor válidos.","danger")
        else:
            r.book_id=book.id; r.borrower_id=borrower.id
            audit(current_user,"EDIÇÃO","Reservation",r.id); db.session.commit(); flash("Reserva atualizada.","success")
            return redirect(url_for("main.reservations"))
    return render_template("reservation_edit.html", r=r, books=Book.query.filter_by(archived=False).order_by(Book.title).all(), borrowers=Borrower.query.filter_by(active=True).order_by(Borrower.name).all())


@bp.route("/reservations/<int:rid>/delete", methods=["POST"])
@login_required
def reservation_delete(rid):
    r=db.get_or_404(Reservation,rid)
    if r.status=="Pendente": r.status="Cancelada"; audit(current_user,"CANCELAMENTO","Reservation",r.id); db.session.commit(); flash("Reserva cancelada e mantida no histórico.","success")
    else: flash("A reserva já finalizada permanece no histórico.","warning")
    return redirect(url_for("main.reservations"))


@bp.route("/solicitacoes/<int:request_id>/delete", methods=["POST"])
@login_required
def work_request_delete(request_id):
    item=db.get_or_404(WorkRequest,request_id)
    if current_user.role!="admin" and item.created_by_id!=current_user.id:
        abort(403)
    audit(current_user,"EXCLUSÃO","WorkRequest",item.id,item.title)
    db.session.delete(item); db.session.commit(); flash("Solicitação excluída.","success")
    return redirect(url_for("main.work_requests"))


@bp.route("/solicitacoes/mensagens/<int:message_id>/edit", methods=["GET", "POST"])
@login_required
def work_request_message_edit(message_id):
    message=db.get_or_404(WorkRequestMessage,message_id)
    if current_user.role!="admin" and message.user_id!=current_user.id: abort(403)
    if request.method=="POST":
        body=request.form.get("body", "").strip()
        if not body or len(body)>10000:
            flash("A mensagem deve ter entre 1 e 10.000 caracteres.","danger")
        else:
            message.body=body; message.work_request.updated_at=datetime.utcnow()
            audit(current_user,"EDIÇÃO DE MENSAGEM","WorkRequestMessage",message.id)
            db.session.commit(); flash("Mensagem atualizada.","success")
            return redirect(url_for("main.work_request_detail",request_id=message.work_request_id)+"#conversa")
    return render_template("message_edit.html", message=message)


@bp.route("/solicitacoes/mensagens/<int:message_id>/delete", methods=["POST"])
@login_required
def work_request_message_delete(message_id):
    message=db.get_or_404(WorkRequestMessage,message_id); request_id=message.work_request_id
    if current_user.role!="admin" and message.user_id!=current_user.id: abort(403)
    audit(current_user,"EXCLUSÃO DE MENSAGEM","WorkRequestMessage",message.id)
    db.session.delete(message); db.session.commit(); flash("Mensagem excluída.","success")
    return redirect(url_for("main.work_request_detail",request_id=request_id)+"#conversa")


@bp.route("/users/<int:uid>/edit", methods=["GET", "POST"])
@login_required
def user_edit(uid):
    admin_required(); u=db.get_or_404(User,uid)
    if request.method=="POST":
        u.name=request.form.get("name", "").strip(); u.username=request.form.get("username", "").strip()
        u.email=request.form.get("email", "").strip().lower(); role=request.form.get("role",u.role)
        if role not in ["admin","servidor"]: abort(400)
        u.role=role; new_password=request.form.get("password", "")
        if new_password and len(new_password)<10: flash("A senha precisa ter pelo menos 10 caracteres.","danger")
        elif not u.name or not u.username or not u.email: flash("Preencha nome, login e e-mail.","danger")
        else:
            if new_password: u.set_password(new_password)
            try:
                audit(current_user,"EDIÇÃO","User",u.id); db.session.commit(); flash("Usuário atualizado.","success")
                return redirect(url_for("main.users"))
            except IntegrityError: db.session.rollback(); flash("Login ou e-mail já está em uso.","danger")
    fields=[("name","Nome completo","text",u.name),("username","Login","text",u.username),("email","E-mail","email",u.email),("role","Perfil","select",u.role),("password","Nova senha (opcional)","password","")]
    return _record_edit("Usuário",u,fields,"Editar usuário")


@bp.route("/users/<int:uid>/delete", methods=["POST"])
@login_required
def user_delete(uid):
    admin_required(); u=db.get_or_404(User,uid)
    if u.id==current_user.id:
        flash("Não é possível excluir a própria conta.","warning")
    elif u.role=="admin" and User.query.filter_by(role="admin",active=True).count()<=1:
        flash("Não é possível excluir o último Administrador Geral ativo.","warning")
    elif (Loan.query.filter_by(user_id=u.id).first() or Evaluation.query.filter_by(user_id=u.id).first() or Recommendation.query.filter_by(user_id=u.id).first() or Reservation.query.filter_by(user_id=u.id).first() or WorkRequest.query.filter(or_(WorkRequest.created_by_id==u.id,WorkRequest.assigned_to_id==u.id)).first() or WorkRequestMessage.query.filter_by(user_id=u.id).first() or WorkRequestRead.query.filter_by(user_id=u.id).first() or NotificationRead.query.filter_by(user_id=u.id).first() or Notification.query.filter_by(resolved_by_id=u.id).first() or AuditLog.query.filter_by(user_id=u.id).first()):
        u.active=False; audit(current_user,"DESATIVAÇÃO","User",u.id); db.session.commit()
        flash("Usuário possui histórico. Foi desativado para preservar a auditoria.","warning")
    else:
        audit(current_user,"EXCLUSÃO","User",u.id); db.session.delete(u); db.session.commit(); flash("Usuário excluído.","success")
    return redirect(url_for("main.users"))


@bp.route("/notifications/<int:nid>/delete", methods=["POST"])
@login_required
def notification_delete(nid):
    admin_required(); n=db.get_or_404(Notification,nid)
    audit(current_user,"EXCLUSÃO","Notification",n.id,n.message)
    db.session.delete(n); db.session.commit(); flash("Notificação removida.","success")
    return redirect(url_for("main.notifications"))


@bp.route("/audit")
@login_required
def audit_log():
    q=AuditLog.query
    if request.args.get("user"): q=q.filter_by(user_id=int(request.args["user"]))
    if request.args.get("action"): q=q.filter_by(action=request.args["action"])
    rows=q.order_by(AuditLog.created_at.desc()).limit(500).all()
    return render_template("audit.html", rows=rows, users=User.query.all())

@bp.route("/reports")
@login_required
def reports():
    return render_template("reports.html")

@bp.route("/reports/export/<kind>")
@login_required
def export_report(kind):
    rows=Book.query.filter_by(archived=False).order_by(Book.title).all()
    data=[["Código","Título","Autor","Categoria","ISBN","Exemplares","Disponíveis"]]
    for b in rows: data.append([b.internal_code,b.title,b.author,b.category or "",b.isbn or "",len(b.copies),sum(1 for c in b.copies if c.status=="Disponível")])
    if kind=="xlsx":
        wb=Workbook(); ws=wb.active; ws.title="Acervo"
        for row in data: ws.append(row)
        out=BytesIO(); wb.save(out); out.seek(0)
        return send_file(out,as_attachment=True,download_name="relatorio_acervo.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if kind=="pdf":
        out=BytesIO(); p=canvas.Canvas(out,pagesize=landscape(A4)); width,height=landscape(A4)
        y=height-35; p.setFont("Helvetica-Bold",14); p.drawString(30,y,"Relatório do acervo"); y-=28
        p.setFont("Helvetica",8)
        for row in data:
            line=" | ".join(str(x or "") for x in row)
            p.drawString(30,y,line[:150]); y-=14
            if y<35: p.showPage(); y=height-35; p.setFont("Helvetica",8)
        p.save(); out.seek(0)
        return send_file(out,as_attachment=True,download_name="relatorio_acervo.pdf",mimetype="application/pdf")
    abort(404)
