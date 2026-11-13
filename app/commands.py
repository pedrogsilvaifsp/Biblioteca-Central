import click
from flask.cli import with_appcontext
from .extensions import db
from .models import User

def register_commands(app):
    @app.cli.command("create-admin")
    def create_admin():
        """Cria o primeiro usuário administrador de forma interativa."""
        from getpass import getpass
        name = click.prompt("Nome completo")
        username = click.prompt("Login")
        email = click.prompt("E-mail")
        password = getpass("Senha (mínimo 10 caracteres): ")
        if len(password) < 10:
            raise click.ClickException("A senha precisa ter pelo menos 10 caracteres.")
        if User.query.filter((User.username==username)|(User.email==email)).first():
            raise click.ClickException("Login ou e-mail já cadastrado.")
        user=User(name=name,username=username,email=email,role="admin",active=True)
        user.set_password(password)
        db.session.add(user); db.session.commit()
        click.echo("Administrador criado com sucesso.")
