import sys
import os
import json
import pytest
from pathlib import Path

# Add project/src to path so we can import app
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from app import create_app


@pytest.fixture
def client():
    """Create a test client for the app."""
    app = create_app()
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client


def test_name_to_extension(client):
    """Test transition from name input to extension prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_USER_INFO_NAME',
            'text': '山田太郎',
            'user_info': {}
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
    assert 'user_info' in response_data
    
    # Assert state transition and user info update
    assert response_data['next_state'] == 'AWAITING_USER_INFO_EXTENSION'
    assert response_data['user_info'].get('name') == '山田太郎'


def test_extension_to_employee_id(client):
    """Test transition from extension input to employee ID prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_USER_INFO_EXTENSION',
            'text': '1234',
            'user_info': {'name': '山田太郎'}
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
    assert 'user_info' in response_data
    
    # Assert state transition and user info update
    assert response_data['next_state'] == 'AWAITING_USER_INFO_EMPLOYEE_ID'
    assert response_data['user_info'].get('name') == '山田太郎'
    assert response_data['user_info'].get('extension') == '1234'


def test_employee_id_to_command(client):
    """Test transition from employee ID input to command prompt."""
    response = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_USER_INFO_EMPLOYEE_ID',
            'text': '56789',
            'user_info': {
                'name': '山田太郎',
                'extension': '1234'
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
    assert 'user_info' in response_data
    
    # Assert state transition and user info update
    assert response_data['next_state'] == 'AWAITING_COMMAND'
    assert response_data['user_info'].get('name') == '山田太郎'
    assert response_data['user_info'].get('extension') == '1234'
    assert response_data['user_info'].get('employee_id') == '56789'
