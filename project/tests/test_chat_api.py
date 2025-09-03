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


def test_initial_chat_request(client):
    """Test the initial chat request with null state."""
    # Prepare test data
    data = {
        "text": None,
        "state": None,
        "user_info": {},
        "context": {}
    }
    
    # Send POST request to /api/chat
    response = client.post(
        '/api/chat',
        data=json.dumps(data),
        content_type='application/json'
    )
    
    # Assert status code is 200
    assert response.status_code == 200
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert required keys exist
    assert 'reply_text' in response_data
    assert 'next_state' in response_data
    
    # Assert next_state is AWAITING_USER_INFO_NAME when state is null
    assert response_data['next_state'] == 'AWAITING_USER_INFO_NAME'
    
    # Additional assertions to verify the expected behavior
    assert 'こんにちは！デモ機予約Botです。お名前を教えてください。' in response_data['reply_text']
    assert 'user_info' in response_data
    assert 'context' in response_data
