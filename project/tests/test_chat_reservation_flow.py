import sys
import os
import json
import pytest
from datetime import date
from pathlib import Path

# Add project/src to path so we can import app and excel_ops
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from app import create_app
import excel_ops


@pytest.fixture
def client():
    """Create a test client for the app."""
    app = create_app()
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client


@pytest.fixture
def fake_find_available_device(monkeypatch):
    """
    Monkeypatch excel_ops.find_available_device to always return
    the first candidate device \"FEデモ機1\".
    """
    def mock_find(excel_path, device_type, start_date, end_date):
        # we ignore arguments for this fake
        return "FEデモ機1"

    monkeypatch.setattr(excel_ops, "find_available_device", mock_find)
    return mock_find


@pytest.fixture
def fake_book(monkeypatch):
    """Monkeypatch excel_ops.book to return a fixed booking ID."""
    captured_args = {}
    
    def mock_book(excel_path, device_name, start_date, end_date, user_info):
        captured_args['excel_path'] = excel_path
        captured_args['device_name'] = device_name
        captured_args['start_date'] = start_date
        captured_args['end_date'] = end_date
        captured_args['user_info'] = user_info
        return 'BID1234'
    
    monkeypatch.setattr(excel_ops, 'book', mock_book)
    return captured_args


def test_command_to_device(client):
    """Test transition from command to device type prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_COMMAND',
            'text': '予約',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {}
        }),
        content_type='application/json'
    )
    
    # Assert status code is 200
    assert response.status_code == 200
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert required keys exist
    assert 'reply_text' in response_data
    assert 'next_state' in response_data
    assert 'context' in response_data
    
    # Assert state transition and context update
    assert response_data['next_state'] == 'AWAITING_DEVICE_TYPE'
    assert response_data['context'].get('intent') == 'reserve'


def test_device_to_dates(client):
    """Test transition from device type to dates prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_DEVICE_TYPE',
            # ユーザーは種類のみを入力
            'text': 'FE',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {'intent': 'reserve'}
        }),
        content_type='application/json'
    )
    
    # Assert status code is 200
    assert response.status_code == 200
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert required keys exist
    assert 'reply_text' in response_data
    assert 'next_state' in response_data
    assert 'context' in response_data
    
    # Assert state transition and context update
    assert response_data['next_state'] == 'AWAITING_DATES'
    assert response_data['context'].get('device_type') == 'FE'


def test_dates_to_confirm(client, fake_find_available_device):
    """Test transition from dates to confirmation prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_DATES',
            'text': '2025-09-10,2025-09-12',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {'intent': 'reserve', 'device_type': 'FE'}
        }),
        content_type='application/json'
    )
    
    # Assert status code is 200
    assert response.status_code == 200
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert required keys exist
    assert 'reply_text' in response_data
    assert 'next_state' in response_data
    assert 'context' in response_data
    
    # Assert state transition and context update
    assert response_data['next_state'] == 'CONFIRM_RESERVATION'
    ctx = response_data['context']
    assert ctx['candidate_device'] == 'FEデモ機1'
    # dates should be normalized with dashes
    assert ctx['start_date'] == '2025-09-10'
    assert ctx['end_date'] == '2025-09-12'


def test_confirm_yes_books(client, fake_book, fake_find_available_device):
    """Test confirmation 'yes' triggers booking and returns to command state."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'CONFIRM_RESERVATION',
            'text': 'はい',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {
                'intent': 'reserve',
                'candidate_device': 'FEデモ機1',
                'start_date': '2025-09-10',
                'end_date': '2025-09-12'
            }
        }),
        content_type='application/json'
    )
    
    # Assert status code is 200
    assert response.status_code == 200
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert required keys exist
    assert 'reply_text' in response_data
    assert 'next_state' in response_data
    
    # Assert state transition
    assert response_data['next_state'] == 'AWAITING_COMMAND'
    
    # Assert booking ID is in reply text
    assert '予約完了' in response_data['reply_text']
    assert 'BID1234' in response_data['reply_text']
    
    # Assert excel_ops.book was called with correct arguments
    assert fake_book['device_name'] == 'FEデモ機1'
    assert isinstance(fake_book['start_date'], date)
    assert isinstance(fake_book['end_date'], date)
    assert fake_book['start_date'].isoformat() == '2025-09-10'
    assert fake_book['end_date'].isoformat() == '2025-09-12'
    assert fake_book['user_info'] == {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'}
