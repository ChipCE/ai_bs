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
