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


def test_missing_json_returns_400(client):
    """Test that sending a request without JSON body returns 400 Bad Request."""
    # Send POST request to /api/chat without any data
    response = client.post(
        '/api/chat',
        content_type='application/json'  # Set content type but don't include data
    )
    
    # Assert status code is 400
    assert response.status_code == 400
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert error message exists
    assert 'error' in response_data


def test_non_object_json_returns_400(client):
    """Test that sending a JSON array instead of object returns 400 Bad Request."""
    # Send POST request to /api/chat with array instead of object
    response = client.post(
        '/api/chat',
        data=json.dumps([]),  # Send empty array instead of object
        content_type='application/json'
    )
    
    # Assert status code is 400
    assert response.status_code == 400
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert error message exists
    assert 'error' in response_data


def test_empty_json_uses_defaults(client):
    """Test that sending an empty JSON object uses default values."""
    # Send POST request to /api/chat with empty object
    response = client.post(
        '/api/chat',
        data=json.dumps({}),  # Send empty object
        content_type='application/json'
    )
    
    # Assert status code is 200
    assert response.status_code == 200
    
    # Parse response JSON
    response_data = json.loads(response.data)
    
    # Assert required keys exist
    assert 'reply_text' in response_data
    assert 'next_state' in response_data
    
    # Assert next_state is AWAITING_USER_INFO_NAME when state is null/not provided
    assert response_data['next_state'] == 'AWAITING_USER_INFO_NAME'
    
    # Assert user_info and context exist and are empty objects
    assert 'user_info' in response_data
    assert 'context' in response_data
