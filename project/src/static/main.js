document.addEventListener('DOMContentLoaded', () => {
    // Always start a fresh session on page load
    sessionStorage.setItem('state', 'null');
    sessionStorage.setItem('user_info', JSON.stringify({}));
    sessionStorage.setItem('context', JSON.stringify({}));

    const messagesContainer = document.getElementById('messages');
    const inputField = document.getElementById('input');
    const sendButton = document.getElementById('send');

    function appendMessage(sender, text) {
        const messageDiv = document.createElement('div');
        messageDiv.innerHTML = `<strong>${sender}:</strong> ${text}`;
        messagesContainer.appendChild(messageDiv);
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
    }

    async function sendMessage(text = null) {
        const state = sessionStorage.getItem('state');
        const user_info = JSON.parse(sessionStorage.getItem('user_info'));
        const context = JSON.parse(sessionStorage.getItem('context'));

        const requestData = {
            text: text,
            state: state === 'null' ? null : state,
            user_info: user_info,
            context: context
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
                appendMessage('あなた', text);
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
            sendMessage(text);
            inputField.value = '';
        }
    });

    // Handle Enter key press
    inputField.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            const text = inputField.value.trim();
            if (text) {
                sendMessage(text);
                inputField.value = '';
            }
        }
    });
});
