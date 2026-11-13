import pytest
from app import create_app
from app.extensions import db
from app.models import User

class TestConfig:
    TESTING=True
    SECRET_KEY="test-secret"
    SQLALCHEMY_DATABASE_URI="sqlite://"
    SQLALCHEMY_TRACK_MODIFICATIONS=False
    WTF_CSRF_ENABLED=False

@pytest.fixture
def client():
    app=create_app(TestConfig)
    with app.app_context():
        db.create_all()
        u=User(name="Admin Teste",username="admin",email="admin@example.com",role="admin")
        u.set_password("senha-segura-123")
        db.session.add(u); db.session.commit()
        with app.test_client() as client:
            yield client
        db.drop_all()

def test_login_and_dashboard(client):
    response=client.post("/login",data={"username":"admin","password":"senha-segura-123"},follow_redirects=True)
    assert response.status_code==200
    assert b"Vis" in response.data or b"vis" in response.data
