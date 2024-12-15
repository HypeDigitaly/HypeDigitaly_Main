(function(d, t) {
    // Vytvoření a vložení stylů
    const styles = `
        .vf-stream-widget {
            position: fixed;
            bottom: 20px;
            right: 20px;
            width: 350px;
            height: 500px;
            border: 1px solid #ccc;
            border-radius: 10px;
            background: white;
            display: flex;
            flex-direction: column;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
        }
        .vf-stream-messages {
            flex: 1;
            overflow-y: auto;
            padding: 15px;
        }
        .vf-stream-message {
            margin: 10px 0;
            padding: 10px;
            border-radius: 10px;
            max-width: 80%;
            font-size: 14px;
            line-height: 1.4;
        }
        .vf-stream-user-message {
            background: #007bff;
            color: white;
            margin-left: auto;
        }
        .vf-stream-bot-message {
            background: #f1f1f1;
            margin-right: auto;
        }
        .vf-stream-input-container {
            padding: 15px;
            border-top: 1px solid #ccc;
            display: flex;
            gap: 10px;
        }
        .vf-stream-input {
            flex: 1;
            padding: 8px;
            border: 1px solid #ccc;
            border-radius: 5px;
            font-size: 14px;
        }
        .vf-stream-button {
            padding: 8px 15px;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }
    `;

    const styleSheet = d.createElement('style');
    styleSheet.textContent = styles;
    d.head.appendChild(styleSheet);

    // Hlavní namespace pro widget
    window.VoiceflowStreamWidget = {
        load: function({ projectID, apiKey = null, versionID = 'production' }) {
            // Vytvoření HTML struktury
            const widget = d.createElement('div');
            widget.className = 'vf-stream-widget';
            widget.innerHTML = `
                <div class="vf-stream-messages"></div>
                <div class="vf-stream-input-container">
                    <input type="text" class="vf-stream-input" placeholder="Napište zprávu...">
                    <button class="vf-stream-button">Odeslat</button>
                </div>
            `;
            d.body.appendChild(widget);

            const messageContainer = widget.querySelector('.vf-stream-messages');
            const input = widget.querySelector('.vf-stream-input');
            const button = widget.querySelector('.vf-stream-button');

            const userID = 'user_' + Math.random().toString(36).substr(2, 9);
            let buffer = '';

            // Přidání zprávy do chatu
            const addMessage = (message, isUser = false) => {
                const messageDiv = d.createElement('div');
                messageDiv.className = `vf-stream-message ${isUser ? 'vf-stream-user-message' : 'vf-stream-bot-message'}`;
                messageDiv.textContent = message;
                messageContainer.appendChild(messageDiv);
                messageContainer.scrollTop = messageContainer.scrollHeight;
            };

            // Zpracování zprávy
            const sendMessage = async (message) => {
                if (!message.trim()) return;

                addMessage(message, true);
                input.value = '';

                try {
                    const response = await fetch(
                        `https://general-runtime.voiceflow.com/v2/project/${projectID}/user/${userID}/interact/stream?completion_events=true`,
                        {
                            method: 'POST',
                            headers: {
                                'Accept': 'text/event-stream',
                                'Authorization': apiKey || projectID,
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify({
                                action: {
                                    type: 'text',
                                    payload: message
                                }
                            })
                        }
                    );

                    const reader = response.body.getReader();
                    const decoder = new TextDecoder();

                    while (true) {
                        const {value, done} = await reader.read();
                        if (done) break;
                        
                        buffer += decoder.decode(value, {stream: true});
                        const lines = buffer.split('\n');
                        buffer = lines.pop();

                        for (const line of lines) {
                            if (line.startsWith('data:')) {
                                try {
                                    const data = JSON.parse(line.slice(5));
                                    if (data.type === 'text') {
                                        addMessage(data.payload.message);
                                    } else if (data.type === 'completion' && data.payload.state === 'content') {
                                        addMessage(data.payload.content);
                                    }
                                } catch (e) {
                                    console.error('Chyba při parsování dat:', e);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.error('Chyba při komunikaci s Voiceflow:', error);
                    addMessage('Omlouváme se, došlo k chybě při komunikaci.');
                }
            };

            // Event listeners
            button.addEventListener('click', () => sendMessage(input.value));
            input.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') sendMessage(input.value);
            });
        }
    };
})(document, 'script'); 