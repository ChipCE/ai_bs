from flask import Flask, request, jsonify, render_template
import os
from datetime import datetime, date

# excel_ops resides in the same src package
import excel_ops

# --------------------------------------------------------------------------- #
# Helpers                                                                    #
# --------------------------------------------------------------------------- #

# Excel ファイルの絶対パスを取得
excel_path = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "data", "デモ機予約表.xlsx")
)


def _parse_date_any(s: str) -> date:
    """
    Accept `YYYY-MM-DD` or `YYYY/MM/DD` and return datetime.date.
    """
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError("日付形式が正しくありません (YYYY-MM-DD または YYYY/MM/DD)")

def create_app():
    app = Flask(__name__)
    
    @app.route('/api/chat', methods=['POST'])
    def chat():
        data = request.get_json(silent=True)
        # Treat `None` (no body / invalid JSON) or non-dict as bad request.
        # An empty dict `{}` is allowed and will fall through to defaults.
        if data is None or not isinstance(data, dict):
            return jsonify({"error": "invalid JSON body"}), 400
        
        text = data.get('text')
        state = data.get('state')
        user_info = data.get('user_info', {})
        context = data.get('context') or {}
        
        if not state:
            reply_text = "こんにちは！デモ機予約Botです。お名前を教えてください。"
            next_state = "AWAITING_USER_INFO_NAME"
        elif state == "AWAITING_USER_INFO_NAME":
            # ユーザー名を取得し、次に内線番号を尋ねる
            if text:  # 簡易チェック、空入力はそのまま同じ質問を繰り返す
                user_info["name"] = text
                reply_text = f"{text} さんですね。内線番号を教えてください。"
                next_state = "AWAITING_USER_INFO_EXTENSION"
            else:
                reply_text = "お名前を入力してください。"
                next_state = "AWAITING_USER_INFO_NAME"
        elif state == "AWAITING_USER_INFO_EXTENSION":
            if text:
                user_info["extension"] = text
                reply_text = "社員番号（職番）を教えてください。"
                next_state = "AWAITING_USER_INFO_EMPLOYEE_ID"
            else:
                reply_text = "内線番号を入力してください。"
                next_state = "AWAITING_USER_INFO_EXTENSION"
        elif state == "AWAITING_USER_INFO_EMPLOYEE_ID":
            if text:
                user_info["employee_id"] = text
                reply_text = "ありがとうございます。ご用件をどうぞ。（予約 / キャンセル など）"
                next_state = "AWAITING_COMMAND"
            else:
                reply_text = "社員番号（職番）を入力してください。"
                next_state = "AWAITING_USER_INFO_EMPLOYEE_ID"
        # ------------------------------------------------------------------ #
        # 予約フロー                                                       #
        # ------------------------------------------------------------------ #
        elif state == "AWAITING_COMMAND" and text == "予約":
            # initiate reservation intent
            context["intent"] = "reserve"
            reply_text = "ご希望のデモ機の種類を入力してください。（例: FE / RT / PC）"
            next_state = "AWAITING_DEVICE_TYPE"
        elif state == "AWAITING_DEVICE_TYPE" and context.get("intent") == "reserve":
            device_type = (text or "").strip()
            if not device_type:
                reply_text = "デモ機の種類を入力してください。"
                next_state = "AWAITING_DEVICE_TYPE"
            else:
                context["device_type"] = device_type
                reply_text = (
                    "予約期間を入力してください（開始日,終了日）。"
                    "例: 2025-09-10,2025/09/12"
                )
                next_state = "AWAITING_DATES"
        elif state == "AWAITING_DATES" and context.get("intent") == "reserve":
            try:
                start_str, end_str = [p.strip() for p in (text or "").split(",")]
                start_date = _parse_date_any(start_str)
                end_date = _parse_date_any(end_str)
            except Exception:
                reply_text = (
                    "日付形式が正しくありません。再度 'YYYY-MM-DD,YYYY/MM/DD' 形式で入力してください。"
                )
                next_state = "AWAITING_DATES"
            else:
                # 検索
                dev_type = context.get("device_type")
                try:
                    candidate = excel_ops.find_available_device(
                        excel_path, dev_type, start_date, end_date
                    )
                except Exception as e:
                    # 例: 月シートが存在しない等
                    reply_text = f"期間の確認中にエラーが発生しました: {e}"
                    next_state = "AWAITING_DATES"
                else:
                    if candidate:
                        context.update(
                            {
                                "start_date": start_date.isoformat(),
                                "end_date": end_date.isoformat(),
                                "candidate_device": candidate,
                            }
                        )
                        reply_text = (
                            f"{candidate} を {start_date:%Y-%m-%d} から "
                            f"{end_date:%Y-%m-%d} で予約します。よろしいですか？（はい / いいえ）"
                        )
                        next_state = "CONFIRM_RESERVATION"
                    else:
                        reply_text = (
                            "指定期間で空いているデモ機が見つかりません。"
                            "別の日付を入力してください。"
                        )
                        next_state = "AWAITING_DATES"
        elif state == "CONFIRM_RESERVATION" and context.get("intent") == "reserve":
            if text == "はい":
                try:
                    start_d = _parse_date_any(context["start_date"])
                    end_d = _parse_date_any(context["end_date"])
                    booking_id = excel_ops.book(
                        excel_path,
                        context["candidate_device"],
                        start_d,
                        end_d,
                        user_info,
                    )
                    reply_text = f"予約完了しました！ 予約ID: {booking_id}"
                except Exception as e:
                    reply_text = f"予約処理でエラーが発生しました: {e}"
                next_state = "AWAITING_COMMAND"
                # clear intent/context
                context = {}
            elif text == "いいえ":
                reply_text = "予約を中止しました。ご用件をどうぞ。"
                next_state = "AWAITING_COMMAND"
                context = {}
            else:
                reply_text = "『はい』または『いいえ』で回答してください。"
                next_state = "CONFIRM_RESERVATION"

        # ------------------------------------------------------------------ #
        # キャンセルフロー                                                 #
        # ------------------------------------------------------------------ #
        elif state == "AWAITING_COMMAND" and text == "キャンセル":
            context["intent"] = "cancel"
            reply_text = "予約IDを入力してください。"
            next_state = "AWAITING_CANCEL_BOOKING_ID"
        elif state == "AWAITING_CANCEL_BOOKING_ID" and context.get("intent") == "cancel":
            booking_id = (text or "").strip()
            if booking_id:
                context["booking_id"] = booking_id
                reply_text = "この予約をキャンセルしますか？（はい / いいえ）"
                next_state = "CANCEL_CONFIRM"
            else:
                reply_text = "予約IDを入力してください。"
                next_state = "AWAITING_CANCEL_BOOKING_ID"
        elif state == "CANCEL_CONFIRM" and context.get("intent") == "cancel":
            if text == "はい":
                try:
                    b_id = context.get("booking_id")
                    excel_ops.cancel(excel_path, b_id)
                    reply_text = "キャンセル完了しました。ご用件をどうぞ。"
                except Exception as e:
                    reply_text = f"キャンセル処理でエラーが発生しました: {e}"
                next_state = "AWAITING_COMMAND"
                context = {}
            elif text == "いいえ":
                reply_text = "キャンセルを中止しました。ご用件をどうぞ。"
                next_state = "AWAITING_COMMAND"
                context = {}
            else:
                reply_text = "『はい』または『いいえ』で回答してください。"
                next_state = "CANCEL_CONFIRM"

        else:
            reply_text = f"入力を受け付けました: {text}"
            next_state = "AWAITING_COMMAND"
        
        response = {
            "reply_text": reply_text,
            "next_state": next_state,
            "user_info": user_info,
            "context": context
        }
        
        return jsonify(response)
    
    @app.route('/', methods=['GET'])
    def index():
        return render_template('index.html')
    
    return app

app = create_app()

if __name__ == '__main__':
    app.run(debug=True)
