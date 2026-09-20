import os
import sys
import json
import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from web.app import app

@pytest.fixture
def client():
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client

def test_index_page(client):
    response = client.get('/')
    assert response.status_code == 200
    assert b"DocuMint" in response.data

def test_docs_page(client):
    response = client.get('/docs')
    assert response.status_code == 200
    assert b"DocuMint Guide" in response.data

def test_api_status(client):
    response = client.get('/api/status')
    assert response.status_code == 200
    data = response.get_json()
    assert "is_running" in data
    assert "logs" in data

def test_api_stats(client):
    response = client.get('/api/stats')
    assert response.status_code == 200
    data = response.get_json()
    assert "summary" in data
    assert "recent" in data

def test_profile_crud(client):
    profile_name = "test_pytest_profile"
    config = {
        "data_path": "E:/test.xlsx",
        "template_path": "E:/template.docx",
        "subject": "Test Subject"
    }

    # Save profile
    res_save = client.post('/api/profiles', json={"name": profile_name, "config": config})
    assert res_save.status_code == 200
    assert res_save.get_json()["status"] == "success"

    # List profiles
    res_list = client.get('/api/profiles')
    assert res_list.status_code == 200
    profiles = res_list.get_json()["profiles"]
    assert f"{profile_name}.json" in profiles

    # Load profile
    res_load = client.get(f'/api/profiles/{profile_name}.json')
    assert res_load.status_code == 200
    loaded_config = res_load.get_json()
    assert loaded_config["subject"] == "Test Subject"

    # Delete profile
    res_del = client.delete(f'/api/profiles/{profile_name}.json')
    assert res_del.status_code == 200
    assert res_del.get_json()["status"] == "success"
