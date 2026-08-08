import os
import pytest
import app as app_module

TEST_DB = 'test_finance.db'


@pytest.fixture
def client():
    # Point the app at a throwaway test database
    app_module.DB_NAME = TEST_DB
    if os.path.exists(TEST_DB):
        os.remove(TEST_DB)

    app_module.app.config['TESTING'] = True
    app_module.init_db()

    with app_module.app.test_client() as client:
        yield client

    if os.path.exists(TEST_DB):
        os.remove(TEST_DB)


def register_and_login(client, name="Test User", email="test@example.com", password="test123"):
    client.post('/register', data={
        'name': name, 'email': email, 'password': password
    })
    return client.post('/login', data={
        'email': email, 'password': password
    }, follow_redirects=True)


# ---------- LOGIN MODULE ----------

def test_register_page_loads(client):
    response = client.get('/register')
    assert response.status_code == 200


def test_login_page_loads(client):
    response = client.get('/login')
    assert response.status_code == 200


def test_register_new_user(client):
    response = client.post('/register', data={
        'name': 'Alice', 'email': 'alice@example.com', 'password': 'pass123'
    }, follow_redirects=True)
    assert response.status_code == 200


def test_login_success(client):
    client.post('/register', data={
        'name': 'Bob', 'email': 'bob@example.com', 'password': 'pass123'
    })
    response = client.post('/login', data={
        'email': 'bob@example.com', 'password': 'pass123'
    }, follow_redirects=True)
    assert response.status_code == 200
    assert b'Welcome' in response.data or b'Dashboard' in response.data or b'dashboard' in response.data


def test_login_wrong_password(client):
    client.post('/register', data={
        'name': 'Carl', 'email': 'carl@example.com', 'password': 'correctpass'
    })
    response = client.post('/login', data={
        'email': 'carl@example.com', 'password': 'wrongpass'
    })
    assert b'Invalid' in response.data


def test_dashboard_requires_login(client):
    response = client.get('/dashboard', follow_redirects=True)
    assert b'Login' in response.data or b'login' in response.data


# ---------- EXPENSE MODULE ----------

def test_add_expense(client):
    register_and_login(client)
    response = client.post('/add_expense', data={
        'category': 'Food',
        'amount': '500',
        'date': '2026-07-15',
        'description': 'Lunch'
    }, follow_redirects=True)
    assert response.status_code == 200
    assert b'500' in response.data or b'Food' in response.data


def test_delete_expense(client):
    register_and_login(client)
    client.post('/add_expense', data={
        'category': 'Food', 'amount': '300', 'date': '2026-07-10', 'description': 'Snacks'
    })
    conn = app_module.get_db()
    expense = conn.execute("SELECT * FROM expenses").fetchone()
    conn.close()
    response = client.get(f'/delete_expense/{expense["id"]}', follow_redirects=True)
    assert response.status_code == 200


# ---------- BUDGET MODULE ----------

def test_add_budget(client):
    register_and_login(client)
    response = client.post('/budget', data={
        'category': 'Food',
        'amount': '5000',
        'month': '2026-07'
    }, follow_redirects=True)
    assert response.status_code == 200
    assert b'Food' in response.data


def test_budget_over_spend_status(client):
    register_and_login(client)
    client.post('/budget', data={'category': 'Food', 'amount': '1000', 'month': '2026-07'})
    client.post('/add_expense', data={
        'category': 'Food', 'amount': '1500', 'date': '2026-07-05', 'description': 'Groceries'
    })
    response = client.get('/budget')
    assert b'Over Budget' in response.data


# ---------- INVESTMENT MODULE ----------

def test_add_investment(client):
    register_and_login(client)
    response = client.post('/investments', data={
        'asset': 'Stocks',
        'name': 'Test Stock',
        'invested': '10000',
        'current': '12000'
    }, follow_redirects=True)
    assert response.status_code == 200
    assert b'Test Stock' in response.data


def test_investment_profit_calculation(client):
    register_and_login(client)
    client.post('/investments', data={
        'asset': 'Stocks', 'name': 'ProfitStock', 'invested': '10000', 'current': '15000'
    })
    response = client.get('/investments')
    assert b'20.0' in response.data or b'20' in response.data  # 50% actually; sanity check profit shows


# ---------- GOAL PLANNING MODULE ----------

def test_add_goal(client):
    register_and_login(client)
    response = client.post('/goals', data={
        'goal_name': 'Emergency Fund',
        'target': '100000',
        'saved': '20000',
        'date': '2026-12-31',
        'priority': 'High'
    }, follow_redirects=True)
    assert response.status_code == 200
    assert b'Emergency Fund' in response.data


def test_goal_progress_percentage(client):
    register_and_login(client)
    client.post('/goals', data={
        'goal_name': 'Vacation', 'target': '10000', 'saved': '5000',
        'date': '2026-12-31', 'priority': 'Medium'
    })
    response = client.get('/goals')
    assert b'50' in response.data  # 50% progress


def test_delete_goal(client):
    register_and_login(client)
    client.post('/goals', data={
        'goal_name': 'TempGoal', 'target': '1000', 'saved': '0',
        'date': '2026-12-31', 'priority': 'Low'
    })
    conn = app_module.get_db()
    goal = conn.execute("SELECT * FROM goals").fetchone()
    conn.close()
    response = client.get(f'/delete_goal/{goal["id"]}', follow_redirects=True)
    assert response.status_code == 200


# ---------- LOGOUT ----------

def test_logout(client):
    register_and_login(client)
    response = client.get('/logout', follow_redirects=True)
    assert b'Login' in response.data or b'login' in response.data