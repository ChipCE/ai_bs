import sys
import json
import pytest
from pathlib import Path

# Add project/src to path so we can import app and excel_ops
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from app import create_app
import excel_ops


@pytest.fixture
def client():
    app = create_app()
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client


def test_dates_invalid_format_stays_in_awaiting_dates(client, monkeypatch):
    # Make find_available_device not used by forcing parse error first
    resp = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_DATES',
            'text': '2025-09-10,09-12-2025',  # invalid second date format
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {'intent': 'reserve', 'device_type': 'FE'}
        }),
        content_type='application/json'
    )
    assert resp.status_code == 200
    data = json.loads(resp.data)
    assert data['next_state'] == 'AWAITING_DATES'
    assert '日付形式' in data['reply_text']


def test_no_available_device_prompts_retry(client, monkeypatch):
    # Monkeypatch to return None (no device)
    monkeypatch.setattr(excel_ops, 'find_available_device', lambda *args, **kwargs: None)
    resp = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'AWAITING_DATES',
            'text': '2025-09-10,2025-09-12',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {'intent': 'reserve', 'device_type': 'FE'}
        }),
        content_type='application/json'
    )
    assert resp.status_code == 200
    data = json.loads(resp.data)
    assert data['next_state'] == 'AWAITING_DATES'
    assert '空いているデモ機が見つかりません' in data['reply_text']


def test_cancel_invalid_answer_stays_confirm(client):
    resp = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'CANCEL_CONFIRM',
            'text': 'うん',  # not はい/いいえ
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {'intent': 'cancel', 'booking_id': 'BID1234'}
        }),
        content_type='application/json'
    )
    assert resp.status_code == 200
    data = json.loads(resp.data)
    assert data['next_state'] == 'CANCEL_CONFIRM'
    assert 'はい' in data['reply_text'] and 'いいえ' in data['reply_text']


def test_cancel_invalid_booking_id_error_message(client, monkeypatch):
    def _raise(_excel_path, _booking_id):
        raise ValueError('予約IDが見つかりません: BID404')
    monkeypatch.setattr(excel_ops, 'cancel', _raise)
    resp = client.post(
        '/api/chat',
        data=json.dumps({
            'state': 'CANCEL_CONFIRM',
            'text': 'はい',
            'user_info': {'name': '山田太郎', 'extension': '1234', 'employee_id': '56789'},
            'context': {'intent': 'cancel', 'booking_id': 'BID404'}
        }),
        content_type='application/json'
    )
    assert resp.status_code == 200
    data = json.loads(resp.data)
    assert data['next_state'] == 'AWAITING_COMMAND'
    assert 'エラー' in data['reply_text'] or 'エラーが発生' in data['reply_text']
