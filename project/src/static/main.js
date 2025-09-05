document.addEventListener('DOMContentLoaded', () => {
    // Always start a fresh session on page load
    sessionStorage.setItem('state', 'null');
    sessionStorage.setItem('user_info', JSON.stringify({}));
    sessionStorage.setItem('context', JSON.stringify({}));
    sessionStorage.setItem('history', JSON.stringify([])); // conversation history

    const messagesContainer = document.getElementById('messages');
    const inputField = document.getElementById('input');
    const sendButton = document.getElementById('send');

    function appendMessage(sender, text) {
        const messageDiv = document.createElement('div');
        messageDiv.innerHTML = `<strong>${sender}:</strong> ${text}`;
        messagesContainer.appendChild(messageDiv);
        messagesContainer.scrollTop = messagesContainer.scrollHeight;

        // --- persist to history ---------------------------------------- //
        try {
            const hist = JSON.parse(sessionStorage.getItem('history') || '[]');
            const role = sender === 'Bot' ? 'assistant' : 'user';
            hist.push({ role, content: text });
            // keep only last 16 turns
            const trimmed = hist.slice(-16);
            sessionStorage.setItem('history', JSON.stringify(trimmed));
        } catch (e) {
            console.error('history update failed', e);
        }
    }

    async function sendMessage(text = null) {
        const state = sessionStorage.getItem('state');
        const user_info = JSON.parse(sessionStorage.getItem('user_info'));
        const context = JSON.parse(sessionStorage.getItem('context'));
        const history = JSON.parse(sessionStorage.getItem('history') || '[]');

        const requestData = {
            text: text,
            state: state === 'null' ? null : state,
            user_info: user_info,
            context: context,
            history: history.slice(-8) // only send last 8
        };

        try {
            const response = await fetch('/api/chat', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestData)
            });

            const data = await response.json();
            
            if (text) {
                // user message already appended before fetch
            }
            
            appendMessage('Bot', data.reply_text);
            
            sessionStorage.setItem('state', data.next_state);
            sessionStorage.setItem('user_info', JSON.stringify(data.user_info));
            sessionStorage.setItem('context', JSON.stringify(data.context));
            
            return data;
        } catch (error) {
            appendMessage('System', 'エラーが発生しました。もう一度お試しください。');
            console.error('Error:', error);
        }
    }

    // Send initial message
    sendMessage();

    // Handle send button click
    sendButton.addEventListener('click', () => {
        const text = inputField.value.trim();
        if (text) {
            appendMessage('あなた', text); // record user turn first
            sendMessage(text);
            inputField.value = '';
        }
    });

    // Handle Enter key press
    inputField.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            const text = inputField.value.trim();
            if (text) {
                appendMessage('あなた', text);
                sendMessage(text);
                inputField.value = '';
            }
        }
    });
});
