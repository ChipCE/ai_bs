from flask import Flask, request, jsonify, render_template
import os
import json
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

# --------------------------------------------------------------------------- #
# LLM (LM Studio) configuration                                               #
# --------------------------------------------------------------------------- #

LM_BASE = os.getenv("LMSTUDIO_BASE_URL", "http://172.17.200.13:1234")
LM_MODEL = os.getenv("LMSTUDIO_MODEL", "meta-llama-llama-3.1-8b-instruct")
LM_ENABLED = os.getenv("LMSTUDIO_ENABLED", "false").lower() in {"1", "true", "yes", "on"}


def _ai_natural_reply(
    history,
    base_reply: str,
    state: str,
    user_info: dict,
    context: dict,
) -> tuple[str, bool]:
    """
    Ask LM Studio to paraphrase `base_reply` into natural Japanese while
    preserving its factual content (IDs, dates, bullet layout).

    Parameters
    ----------
    history : list[dict]
        Recent message history items like {'role': 'user'|'assistant', 'content': str}
        Only last ~8 turns are sent.
    base_reply : str
        Deterministic reply produced by state-machine.
    state : str
        Next state (for possible future prompt conditioning, not used now).
    user_info / context : dict
        Extra info if needed (currently unused).
    """
    if not LM_ENABLED:
        return base_reply, False
    try:
        import requests  # local import to avoid hard dependency in tests
    except Exception:
        return base_reply, False

    # Build prompt messages
    messages = [
        {
            "role": "system",
            "content": (
                "あなたは社内デモ機予約Botです。常に日本語で、丁寧かつ簡潔に自然な口調で返答してください。"
                "これから示す『基準応答』の意味内容・語順に含まれるID/日付/数値/箇条書きは決して改変しないで下さい。"
                "必要なら短い前置きや補足を加えて構いませんが、重要情報は保持してください。"
            ),
        }
    ]

    # Add up to 8 recent turns
    for m in (history or [])[-8:]:
        role = "assistant" if m.get("role") == "assistant" else "user"
        content = str(m.get("content", ""))
        if content:
            messages.append({"role": role, "content": content})

    # Provide base reply to be rewritten
    messages.append({"role": "system", "content": f"基準応答:\n{base_reply}"})
    messages.append(
        {
            "role": "user",
            "content": (
                "上の基準応答の意味を絶対に変えずに、自然な日本語に言い換えてください。"
                "ID,日付,箇条書きはそのまま維持してください。"
            ),
        }
    )

    try:
        resp = requests.post(
            f"{LM_BASE}/v1/chat/completions",
            headers={"Content-Type": "application/json"},
            data=json.dumps(
                {
                    "model": LM_MODEL,
                    "messages": messages,
                    "temperature": 0.3,
                    "max_tokens": 512,
                }
            ),
            timeout=8,
        )
        resp.raise_for_status()
        data = resp.json()
        content = (
            data.get("choices", [{}])[0]
            .get("message", {})
            .get("content", "")
            .strip()
        )
        return (content or base_reply), True
    except Exception:
        # fallback silently
        return base_reply, False


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
        history = data.get('history') or []
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
            # fetch cancellable bookings for this user
            try:
                items = excel_ops.list_cancellable_bookings(excel_path, user_info)
            except Exception:
                items = []
            if items:
                # show up to 10 entries
                lines = [
                    f"- {it['booking_id']} ({it['device_name']}: {it['start_date']}→{it['end_date']})"
                    for it in items[:10]
                ]
                extra = "\n" + "\n".join(lines)
            else:
                extra = "\n(キャンセル可能な予約は見つかりませんでした)"
            reply_text = "予約IDを入力してください。" + extra
            next_state = "AWAITING_CANCEL_BOOKING_ID"
        elif state == "AWAITING_CANCEL_BOOKING_ID" and context.get("intent") == "cancel":
            booking_id = (text or "").strip()
            if booking_id:
                context["booking_id"] = booking_id
                reply_text = "この予約をキャンセルしますか？（はい / いいえ）"
                next_state = "CANCEL_CONFIRM"
            else:
                try:
                    items = excel_ops.list_cancellable_bookings(excel_path, user_info)
                except Exception:
                    items = []
                if items:
                    lines = [
                        f"- {it['booking_id']} ({it['device_name']}: {it['start_date']}→{it['end_date']})"
                        for it in items[:10]
                    ]
                    extra = "\n" + "\n".join(lines)
                else:
                    extra = "\n(キャンセル可能な予約は見つかりませんでした)"
                reply_text = "予約IDを入力してください。" + extra
                next_state = "AWAITING_CANCEL_BOOKING_ID"

        # ------------------------------------------------------------------ #
        # 予約確認フロー                                                 #
        # ------------------------------------------------------------------ #
        elif state == "AWAITING_COMMAND" and text in ("確認", "予約確認", "予約状況"):
            try:
                items = excel_ops.list_user_bookings(excel_path, user_info)
            except Exception:
                items = []

            if items:
                lines = [
                    f"- {it['booking_id']} [{it['status']}] {it['device_name']} "
                    f"{it['start_date']}→{it['end_date']}"
                    for it in items[:10]
                ]
                reply_text = "あなたの予約一覧:\n" + "\n".join(lines)
            else:
                reply_text = "あなたの予約は見つかりませんでした。"
            next_state = "AWAITING_COMMAND"
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
        
        # --- AI natural rephrase (optional) -------------------------------- #
        reply_text, ai_used = _ai_natural_reply(
            history, reply_text, next_state, user_info, context
        )

        response = {
            "reply_text": reply_text,
            "next_state": next_state,
            "user_info": user_info,
            "context": context,
            "ai_used": ai_used,
        }
        
        return jsonify(response)
    
    @app.route('/', methods=['GET'])
    def index():
        return render_template('index.html')
    
    return app

app = create_app()

if __name__ == '__main__':
    app.run(debug=True)
