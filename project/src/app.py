from flask import Flask, request, jsonify, render_template

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
        context = data.get('context', {})
        
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
