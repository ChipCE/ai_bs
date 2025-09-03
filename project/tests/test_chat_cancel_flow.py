import sys
import os
import json
import pytest
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
def fake_cancel(monkeypatch):
    """Monkeypatch excel_ops.cancel to capture booking_id calls."""
    captured_args = {'called': False, 'booking_id': None}
    
    def mock_cancel(excel_path, booking_id):
        captured_args['called'] = True
        captured_args['booking_id'] = booking_id
        # Return nothing (None) like the real cancel function
    
    monkeypatch.setattr(excel_ops, 'cancel', mock_cancel)
    return captured_args


def test_command_to_cancel_id(client):
    """Test transition from command to cancel booking ID prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_COMMAND',
            'text': 'キャンセル',
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
    assert response_data['next_state'] == 'AWAITING_CANCEL_BOOKING_ID'
    assert response_data['context'].get('intent') == 'cancel'


def test_cancel_id_to_confirm(client):
    """Test transition from booking ID input to confirmation prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_CANCEL_BOOKING_ID',
            'text': 'BID1234',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {'intent': 'cancel'}
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
    assert response_data['next_state'] == 'CANCEL_CONFIRM'
    assert response_data['context'].get('booking_id') == 'BID1234'


def test_confirm_yes_cancels(client, fake_cancel):
    """Test confirmation 'yes' triggers cancellation and returns to command state."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'CANCEL_CONFIRM',
            'text': 'はい',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {
                'intent': 'cancel',
                'booking_id': 'BID1234'
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
    
    # Assert cancellation message in reply text
    assert 'キャンセル完了' in response_data['reply_text']
    
    # Assert excel_ops.cancel was called with correct booking ID
    assert fake_cancel['called'] is True
    assert fake_cancel['booking_id'] == 'BID1234'


def test_confirm_no_aborts(client, fake_cancel):
    """Test confirmation 'no' aborts cancellation and returns to command state."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'CANCEL_CONFIRM',
            'text': 'いいえ',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {
                'intent': 'cancel',
                'booking_id': 'BID1234'
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
    
    # Assert abort message in reply text
    assert '中止' in response_data['reply_text']
    
    # Assert excel_ops.cancel was NOT called
    assert fake_cancel['called'] is False
