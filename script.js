// M365 Copilot Application JavaScript
class M365Copilot {
    constructor() {
        this.currentTab = 'web';
        this.currentChat = null;
        this.chats = [];
        this.settings = {
            apiUrl: '',
            apiKey: '',
            maxTokens: 1000
        };
        
        this.init();
    }
    
    init() {
        this.loadSettings();
        this.loadChats();
        this.bindEvents();
        this.updateConversationsList();
        this.bindWelcomeScreenInput();
        this.initSpeechRecognition();
        this.initTextToSpeech();
    }
    
    // Load settings from localStorage
    loadSettings() {
        const savedSettings = localStorage.getItem('chatAppSettings');
        if (savedSettings) {
            this.settings = { ...this.settings, ...JSON.parse(savedSettings) };
        }
    }
    
    // Save settings to localStorage
    saveSettings() {
        localStorage.setItem('chatAppSettings', JSON.stringify(this.settings));
    }
    
    // Load chats from localStorage
    loadChats() {
        const savedChats = localStorage.getItem('m365CopilotChats');
        if (savedChats) {
            const parsed = JSON.parse(savedChats);
            // Handle both old object format and new array format
            this.chats = Array.isArray(parsed) ? parsed : Object.values(parsed);
        }
        
        const currentChatId = localStorage.getItem('currentChatId');
        if (currentChatId) {
            this.currentChat = this.chats.find(chat => chat.id === currentChatId) || null;
        }
    }
    
    // Save chats to localStorage
    saveChats() {
        localStorage.setItem('m365CopilotChats', JSON.stringify(this.chats));
        if (this.currentChat) {
            localStorage.setItem('currentChatId', this.currentChat.id);
        }
    }
    
    // Bind event listeners
    bindEvents() {
        // Tab switching
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this.switchTab(btn.dataset.tab);
            });
        });
        
        // New Chat button
        const newChatBtn = document.getElementById('newChatBtn');
        if (newChatBtn) {
            newChatBtn.addEventListener('click', () => {
                this.createNewChat();
            });
        }
        
        // Navigation items
        document.querySelectorAll('.nav-item').forEach(item => {
            item.addEventListener('click', () => {
                this.handleNavigation(item);
            });
        });
        
        // Settings modal buttons
        const closeSettingsBtn = document.getElementById('closeSettingsBtn');
        const saveSettingsBtn = document.getElementById('saveSettingsBtn');
        const resetSettingsBtn = document.getElementById('resetSettingsBtn');
        
        if (closeSettingsBtn) {
            closeSettingsBtn.addEventListener('click', () => this.closeSettings());
        }
        if (saveSettingsBtn) {
            saveSettingsBtn.addEventListener('click', () => this.saveSettingsFromModal());
        }
        if (resetSettingsBtn) {
            resetSettingsBtn.addEventListener('click', () => this.resetSettings());
        }
        
        // Close modal when clicking overlay
        const settingsModal = document.getElementById('settingsModal');
        if (settingsModal) {
            settingsModal.addEventListener('click', (e) => {
                if (e.target.id === 'settingsModal') {
                    this.closeSettings();
                }
            });
        }
    }
    
    // Switch tabs
    switchTab(tabName) {
        // Update active tab
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
        
        this.currentTab = tabName;
        console.log(`Switched to ${tabName} tab`);
    }
    
    
    // Toggle microphone
    toggleMicrophone() {
        console.log('Microphone toggled');
        
        if (!this.speechRecognitionSupported) {
            alert('Speech recognition is not supported in this browser. Please use Chrome, Edge, or Safari.');
            return;
        }
        
        if (this.recognition) {
            // Check if recognition is currently active
            const micBtn = document.getElementById('micBtn');
            const isListening = micBtn && micBtn.classList.contains('listening');
            
            if (isListening) {
                // Stop recognition
                this.recognition.stop();
                console.log('Stopping voice recognition');
            } else {
                // Start recognition
                try {
                    this.recognition.start();
                    console.log('Starting voice recognition');
                } catch (error) {
                    console.error('Error starting voice recognition:', error);
                    alert('Error starting voice recognition. Please try again.');
                }
            }
        }
    }
    
    // Toggle audio - now with text-to-speech functionality
    toggleAudio() {
        console.log('Audio toggled');
        
        // Get the last assistant message to read aloud
        const lastAssistantMessage = this.getLastAssistantMessage();
        if (lastAssistantMessage) {
            this.speakText(lastAssistantMessage);
        } else {
            // If no assistant message, speak a welcome message
            this.speakText("Welcome! I am WindowsFlowAgent. I can help you with automating tasks such as playing videos on streaming platforms, opening applications like Outlook, and clicking icons to automate workflows. Just tell me what you'd like me to do!");
        }
    }
    
    // Get last assistant message content
    getLastAssistantMessage() {
        if (!this.currentChat || !this.currentChat.messages) {
            return null;
        }
        
        // Find the last assistant message
        for (let i = this.currentChat.messages.length - 1; i >= 0; i--) {
            const message = this.currentChat.messages[i];
            if (message.role === 'assistant' && !message.isError) {
                return this.stripHtml(message.content);
            }
        }
        return null;
    }
    
    // Text-to-speech functionality
    speakText(text) {
        // Stop any currently playing speech
        if (window.speechSynthesis.speaking) {
            window.speechSynthesis.cancel();
            this.updateAudioButton(false);
            return;
        }
        
        if (!text || text.trim() === '') {
            console.log('No text to speak');
            return;
        }
        
        // Check if speech synthesis is supported
        if (!this.ttsSupported) {
            alert('Text-to-speech is not supported in this browser. Please use Chrome, Firefox, Safari, or Edge.');
            return;
        }
        
        const utterance = new SpeechSynthesisUtterance(text);
        
        // Configure speech settings
        utterance.rate = 0.9;  // Slightly slower than normal
        utterance.pitch = 1.0;
        utterance.volume = 0.8;
        
        // Try to use a more natural voice if available
        const voices = this.voices || window.speechSynthesis.getVoices();
        const preferredVoice = voices.find(voice => 
            voice.name.includes('Natural') || 
            voice.name.includes('Enhanced') || 
            voice.name.includes('Premium') ||
            voice.name.includes('Neural') ||
            voice.name.includes('Google') ||
            (voice.localService === false && voice.lang.startsWith('en'))
        );
        if (preferredVoice) {
            utterance.voice = preferredVoice;
            console.log('Using voice:', preferredVoice.name);
        }
        
        // Event listeners
        utterance.onstart = () => {
            console.log('Speech started');
            this.updateAudioButton(true);
        };
        
        utterance.onend = () => {
            console.log('Speech ended');
            this.updateAudioButton(false);
        };
        
        utterance.onerror = (event) => {
            console.error('Speech error:', event.error);
            this.updateAudioButton(false);
        };
        
        // Start speaking
        window.speechSynthesis.speak(utterance);
    }
    
    // Update audio button state
    updateAudioButton(isSpeaking) {
        const audioBtn = document.getElementById('audioBtn');
        if (audioBtn) {
            const icon = audioBtn.querySelector('i');
            if (isSpeaking) {
                audioBtn.classList.add('speaking');
                icon.className = 'fas fa-stop';
                audioBtn.title = 'Stop speaking';
            } else {
                audioBtn.classList.remove('speaking');
                icon.className = 'fas fa-volume-up';
                audioBtn.title = 'Read response aloud';
            }
        }
    }
    
    // Strip HTML tags from text for speech
    stripHtml(html) {
        const temp = document.createElement('div');
        temp.innerHTML = html;
        return temp.textContent || temp.innerText || '';
    }
    
    // Escape text for use in HTML attributes
    escapeForAttribute(text) {
        return text
            .replace(/'/g, '&apos;')
            .replace(/"/g, '&quot;')
            .replace(/\n/g, ' ')
            .replace(/\r/g, ' ')
            .replace(/\\/g, '\\\\');
    }
    
    // Handle navigation
    handleNavigation(item) {
        // Remove active class from all nav items
        document.querySelectorAll('.nav-item').forEach(nav => {
            nav.classList.remove('active');
        });
        
        // Add active class to clicked item
        item.classList.add('active');
        
        const navText = item.querySelector('span').textContent;
        console.log(`Navigation to: ${navText}`);
        
        // Handle navigation
        if (navText === 'Conversations') {
            this.showConversations();
        }
    }
    
    
    // Navigation methods
    showConversations() {
        console.log('Showing conversations');
        // Add conversations functionality here
    }
    
    // Create new chat
    createNewChat() {
        // Reset to welcome screen
        const chatContent = document.getElementById('chatContent');
        chatContent.innerHTML = `
            <div class="welcome-section">
                <h1 class="welcome-title">Welcome. I am WindowsFlowAgent.</h1>
                <div class="welcome-description">
                    <p>I can help you with automating tasks such as:</p>
                    <ul class="capability-list">
                        <li>Playing videos on Amazon Prime, YouTube, or Netflix.</li>
                        <li>Opening Outlook and navigating buttons.</li>
                        <li>Clicking icons and automating workflows.</li>
                    </ul>
                    <p class="welcome-cta">Just tell me what you'd like me to do!</p>
                </div>
                
                <div class="input-section">
                    <div class="input-container">
                        <input 
                            type="text" 
                            id="messageInput" 
                            placeholder="Message Copilot"
                            class="message-input"
                        >
                        <div class="input-actions">
                            <button class="input-action-btn" id="micBtn" title="Voice to text">
                                <i class="fas fa-microphone"></i>
                            </button>
                            <button class="input-action-btn" id="audioBtn">
                                <i class="fas fa-volume-up"></i>
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        // Reset current chat
        this.currentChat = null;
        this.updateConversationsList();
        
        // Re-bind events
        this.bindWelcomeScreenInput();
    }
    
    // Initialize text-to-speech
    initTextToSpeech() {
        if ('speechSynthesis' in window) {
            // Load voices when they become available
            if (speechSynthesis.onvoiceschanged !== undefined) {
                speechSynthesis.onvoiceschanged = () => {
                    this.voices = speechSynthesis.getVoices();
                    console.log('Voices loaded:', this.voices.length);
                };
            }
            this.voices = speechSynthesis.getVoices();
            this.ttsSupported = true;
        } else {
            this.ttsSupported = false;
            console.log('Text-to-speech not supported');
        }
    }
    
    // Initialize speech recognition
    initSpeechRecognition() {
        if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
            const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
            this.recognition = new SpeechRecognition();
            this.recognition.continuous = false;
            this.recognition.interimResults = false;
            this.recognition.lang = 'en-US';
            
            this.recognition.onstart = () => {
                console.log('Voice recognition started');
                this.updateMicButton(true);
            };
            
            this.recognition.onresult = (event) => {
                const transcript = event.results[0][0].transcript;
                console.log('Voice recognition result:', transcript);
                const messageInput = document.getElementById('messageInput');
                if (messageInput) {
                    messageInput.value = transcript;
                    messageInput.focus();
                }
            };
            
            this.recognition.onerror = (event) => {
                console.error('Voice recognition error:', event.error);
                this.updateMicButton(false);
                if (event.error === 'not-allowed') {
                    alert('Microphone access denied. Please allow microphone access and try again.');
                } else if (event.error === 'no-speech') {
                    alert('No speech detected. Please try again.');
                } else {
                    alert('Voice recognition error: ' + event.error);
                }
            };
            
            this.recognition.onend = () => {
                console.log('Voice recognition ended');
                this.updateMicButton(false);
            };
            
            this.speechRecognitionSupported = true;
        } else {
            this.speechRecognitionSupported = false;
            console.log('Speech recognition not supported');
        }
    }
    
    // Update microphone button state
    updateMicButton(isListening) {
        const micBtn = document.getElementById('micBtn');
        if (micBtn) {
            const icon = micBtn.querySelector('i');
            if (isListening) {
                micBtn.classList.add('listening');
                icon.className = 'fas fa-stop';
                micBtn.title = 'Stop listening';
            } else {
                micBtn.classList.remove('listening');
                icon.className = 'fas fa-microphone';
                micBtn.title = 'Voice to text';
            }
        }
    }
    
    // Bind welcome screen input events
    bindWelcomeScreenInput() {
        console.log('Binding welcome screen input events'); // Debug log
        const messageInput = document.getElementById('messageInput');
        const micBtn = document.getElementById('micBtn');
        const audioBtn = document.getElementById('audioBtn');
        
        console.log('Message input found:', !!messageInput); // Debug log
        
        if (messageInput) {
            // Remove any existing listeners to avoid duplicates
            const newInput = messageInput.cloneNode(true);
            messageInput.parentNode.replaceChild(newInput, messageInput);
            
            console.log('Binding Enter key event to message input'); // Debug log
            newInput.addEventListener('keydown', (e) => {
                console.log('Key pressed:', e.key); // Debug log
                if (e.key === 'Enter') {
                    e.preventDefault();
                    console.log('Enter key pressed, calling sendMessage'); // Debug log
                    this.sendMessage();
                }
            });
        }
        
        if (micBtn) {
            micBtn.addEventListener('click', () => this.toggleMicrophone());
        }
        
        if (audioBtn) {
            audioBtn.addEventListener('click', () => this.toggleAudio());
        }
    }
    
    // Update conversations list in sidebar
    updateConversationsList() {
        const conversationsList = document.getElementById('conversationsList');
        conversationsList.innerHTML = '';
        
        if (this.chats.length === 0) {
            conversationsList.innerHTML = '<div style="padding: 16px; color: #8a8886; font-size: 12px; text-align: center;">No conversations yet</div>';
            return;
        }
        
        // Sort chats by creation time (newest first)
        const sortedChats = [...this.chats].sort((a, b) => 
            new Date(b.createdAt) - new Date(a.createdAt)
        );
        
        sortedChats.forEach(chat => {
            const conversationEl = document.createElement('div');
            conversationEl.className = `conversation-item ${chat === this.currentChat ? 'active' : ''}`;
            conversationEl.dataset.chatId = chat.id;
            
            const title = this.getChatTitle(chat);
            const preview = this.getChatPreview(chat);
            const time = this.formatDate(chat.createdAt);
            
            conversationEl.innerHTML = `
                <div class="conversation-title">${title}</div>
                <div class="conversation-preview">${preview}</div>
                <div class="conversation-time">${time}</div>
            `;
            
            conversationEl.addEventListener('click', () => {
                this.switchToConversation(chat);
            });
            
            conversationsList.appendChild(conversationEl);
        });
    }
    
    // Get chat title
    getChatTitle(chat) {
        if (chat.messages.length > 0) {
            const firstMessage = chat.messages[0];
            return firstMessage.content.substring(0, 30) + (firstMessage.content.length > 30 ? '...' : '');
        }
        return 'New Chat';
    }
    
    // Get chat preview
    getChatPreview(chat) {
        if (chat.messages.length > 1) {
            const lastMessage = chat.messages[chat.messages.length - 1];
            const prefix = lastMessage.role === 'user' ? 'You: ' : 'Copilot: ';
            return prefix + lastMessage.content.substring(0, 40) + (lastMessage.content.length > 40 ? '...' : '');
        } else if (chat.messages.length === 1) {
            return 'Started conversation';
        }
        return 'New chat';
    }
    
    // Switch to a conversation
    switchToConversation(chat) {
        this.currentChat = chat;
        this.updateConversationsList();
        this.displayConversation(chat);
    }
    
    // Display a conversation
    displayConversation(chat) {
        const chatContent = document.getElementById('chatContent');
        chatContent.innerHTML = `
            <div class="chat-messages" id="chatMessages">
                ${chat.messages.map(message => this.renderMessage(message)).join('')}
            </div>
            <div class="input-section">
                <div class="input-container">
                    <input 
                        type="text" 
                        id="messageInput" 
                        placeholder="Message Copilot"
                        class="message-input"
                    >
                    <div class="input-actions">
                        <button class="input-action-btn" id="micBtn" title="Voice to text">
                            <i class="fas fa-microphone"></i>
                        </button>
                        <button class="input-action-btn" id="audioBtn" title="Read response aloud">
                            <i class="fas fa-volume-up"></i>
                        </button>
                    </div>
                </div>
            </div>
        `;
        
        // Re-bind events for the input
        this.bindMessageInputEvents();
        
        // Scroll to bottom
        const chatMessages = document.getElementById('chatMessages');
        if (chatMessages) {
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
    }
    
    // Render a message
    renderMessage(message) {
        if (message.role === 'user') {
            return `
                <div class="message user-message">
                    <div class="message-content">
                        <div class="message-text">${message.content}</div>
                        <div class="message-time">${this.formatTime(message.timestamp)}</div>
                    </div>
                    <div class="message-avatar">
                        <div class="user-avatar">
                            <img src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle fill='%234285f4' cx='16' cy='16' r='16'/%3E%3Ctext x='16' y='21' font-family='Arial' font-size='14' fill='white' text-anchor='middle'%3EJ%3C/text%3E%3C/svg%3E" alt="User">
                        </div>
                    </div>
                </div>
            `;
        } else {
            return `
                <div class="message assistant-message ${message.isError ? 'error' : ''}">
                    <div class="message-avatar">
                        <div class="copilot-logo">
                            <div class="logo-squares">
                                <div class="square red"></div>
                                <div class="square yellow"></div>
                                <div class="square green"></div>
                                <div class="square blue"></div>
                            </div>
                        </div>
                    </div>
                    <div class="message-content">
                        <div class="message-text">${this.formatMessageContent(message.content)}</div>
                        <div class="message-actions">
                            <button class="message-speak-btn" title="Read this message aloud" onclick="window.copilot.speakText('${this.escapeForAttribute(this.stripHtml(message.content))}')">
                                <i class="fas fa-volume-up"></i>
                            </button>
                            <div class="message-time">${this.formatTime(message.timestamp)}</div>
                        </div>
                    </div>
                </div>
            `;
        }
    }
    
    // Switch to a specific chat
    switchToChat(chatId) {
        if (this.chats[chatId]) {
            this.currentChatId = chatId;
            this.saveChats();
            this.updateUI();
        }
    }
    
    // Clear current chat
    clearCurrentChat() {
        if (this.currentChatId && this.chats[this.currentChatId]) {
            if (confirm('Are you sure you want to clear this chat? This action cannot be undone.')) {
                this.chats[this.currentChatId].messages = [];
                this.chats[this.currentChatId].title = 'New Chat';
                this.chats[this.currentChatId].updatedAt = new Date().toISOString();
                this.saveChats();
                this.updateUI();
            }
        }
    }
    
    // Delete a chat
    deleteChat(chatId, event) {
        if (event) {
            event.stopPropagation();
        }
        
        // Remove the chat from the chats object
        delete this.chats[chatId];
        
        // Handle current chat deletion
        if (this.currentChatId === chatId) {
            // Switch to another chat or reset to no chat
            const remainingChats = Object.keys(this.chats);
            if (remainingChats.length > 0) {
                // Switch to the most recently updated remaining chat
                const sortedChats = Object.values(this.chats).sort((a, b) => 
                    new Date(b.updatedAt) - new Date(a.updatedAt)
                );
                this.currentChatId = sortedChats[0].id;
            } else {
                // No chats remaining, reset current chat
                this.currentChatId = null;
            }
        }
        
        // Save changes and update UI
        this.saveChats();
        this.updateUI();
        
        // Clear message input if we're on a new/empty state
        if (!this.currentChatId) {
            const messageInput = document.getElementById('messageInput');
            messageInput.value = '';
            this.updateCharacterCount();
            this.updateSendButton();
        }
    }
    
    // Send message
    async sendMessage() {
        console.log('sendMessage called'); // Debug log
        const input = document.getElementById('messageInput');
        if (!input) {
            console.log('Message input not found');
            return;
        }
        
        const message = input.value.trim();
        console.log('Message:', message); // Debug log
        
        if (!message) {
            console.log('Empty message, returning');
            return;
        }
        
        // Start a new conversation
        this.startNewConversation(message);
        
        // Clear input
        input.value = '';
        
        // Show typing indicator and call API
        this.showTypingIndicator();
        await this.callAPI(message);
    }
    
    // Start a new conversation
    startNewConversation(firstMessage) {
        // Hide welcome section and show chat interface
        const chatContent = document.getElementById('chatContent');
        chatContent.innerHTML = `
            <div class="chat-messages" id="chatMessages">
                <div class="message user-message">
                    <div class="message-content">
                        <div class="message-text">${firstMessage}</div>
                        <div class="message-time">${this.formatTime(new Date().toISOString())}</div>
                    </div>
                    <div class="message-avatar">
                        <div class="user-avatar">
                            <img src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle fill='%234285f4' cx='16' cy='16' r='16'/%3E%3Ctext x='16' y='21' font-family='Arial' font-size='14' fill='white' text-anchor='middle'%3EJ%3C/text%3E%3C/svg%3E" alt="User">
                        </div>
                    </div>
                </div>
            </div>
            <div class="input-section">
                <div class="input-container">
                    <input 
                        type="text" 
                        id="messageInput" 
                        placeholder="Message Copilot"
                        class="message-input"
                    >
                    <div class="input-actions">
                        <button class="input-action-btn" id="micBtn" title="Voice to text">
                            <i class="fas fa-microphone"></i>
                        </button>
                        <button class="input-action-btn" id="audioBtn" title="Read response aloud">
                            <i class="fas fa-volume-up"></i>
                        </button>
                    </div>
                </div>
            </div>
        `;
        
        // Re-bind events for the new input
        this.bindMessageInputEvents();
        
        // Store the conversation
        this.currentChat = {
            id: 'chat_' + Date.now(),
            messages: [{
                role: 'user',
                content: firstMessage,
                timestamp: new Date().toISOString()
            }],
            createdAt: new Date().toISOString()
        };
        
        this.chats.push(this.currentChat);
        this.saveChats();
        this.updateConversationsList();
    }
    
    // Bind events for message input
    bindMessageInputEvents() {
        const messageInput = document.getElementById('messageInput');
        const micBtn = document.getElementById('micBtn');
        const audioBtn = document.getElementById('audioBtn');
        
        if (messageInput) {
            messageInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    this.sendFollowUpMessage();
                }
            });
        }
        
        if (micBtn) {
            micBtn.addEventListener('click', () => this.toggleMicrophone());
        }
        
        if (audioBtn) {
            audioBtn.addEventListener('click', () => this.toggleAudio());
        }
    }
    
    // Send follow-up message
    async sendFollowUpMessage() {
        const input = document.getElementById('messageInput');
        const message = input.value.trim();
        
        if (!message) return;
        
        // Add user message to chat
        const chatMessages = document.getElementById('chatMessages');
        const userMessageEl = document.createElement('div');
        userMessageEl.className = 'message user-message';
        userMessageEl.innerHTML = `
            <div class="message-content">
                <div class="message-text">${message}</div>
                <div class="message-time">${this.formatTime(new Date().toISOString())}</div>
            </div>
            <div class="message-avatar">
                <div class="user-avatar">
                    <img src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Ccircle fill='%234285f4' cx='16' cy='16' r='16'/%3E%3Ctext x='16' y='21' font-family='Arial' font-size='14' fill='white' text-anchor='middle'%3EJ%3C/text%3E%3C/svg%3E" alt="User">
                </div>
            </div>
        `;
        
        chatMessages.appendChild(userMessageEl);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        
        // Store message
        if (this.currentChat) {
            this.currentChat.messages.push({
                role: 'user',
                content: message,
                timestamp: new Date().toISOString()
            });
        }
        
        // Clear input
        input.value = '';
        
        // Update conversations list
        this.updateConversationsList();
        
        // Show typing indicator and call API
        this.showTypingIndicator();
        await this.callAPI(message);
    }
    
    // Show typing indicator
    showTypingIndicator() {
        const chatMessages = document.getElementById('chatMessages');
        if (chatMessages) {
            const typingIndicator = document.createElement('div');
            typingIndicator.className = 'message assistant-message typing';
            typingIndicator.innerHTML = `
                <div class="message-avatar">
                    <div class="copilot-logo">
                        <div class="logo-squares">
                            <div class="square red"></div>
                            <div class="square yellow"></div>
                            <div class="square green"></div>
                            <div class="square blue"></div>
                        </div>
                    </div>
                </div>
                <div class="message-content">
                    <div class="typing-indicator">
                        <span>Copilot is thinking</span>
                        <div class="typing-dots">
                            <div class="typing-dot"></div>
                            <div class="typing-dot"></div>
                            <div class="typing-dot"></div>
                        </div>
                    </div>
                </div>
            `;
            
            chatMessages.appendChild(typingIndicator);
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
    }
    
    // Hide typing indicator
    hideTypingIndicator() {
        const typingIndicator = document.querySelector('.message.typing');
        if (typingIndicator) {
            typingIndicator.remove();
        }
    }


    // ...existing code...
// Call REST API
async callAPI(userMessage) {
    try {
        // Build the URL with the user prompt
        const encodedPrompt = encodeURIComponent(userMessage);
        const apiUrl = `http://127.0.0.1:8000/assistant/${encodedPrompt}`;

        const response = await fetch(apiUrl, {
            method: 'GET'
        });

        if (!response.ok) {
            throw new Error(`API request failed: ${response.status} ${response.statusText}`);
        }

        const data = await response.json();

        // Extract response text (adapt this based on your API response format)
        let assistantResponse = '';
        if (data.response) {
            assistantResponse = data.response;
        } else if (data.message) {
            assistantResponse = data.message;
        } else if (data.content) {
            assistantResponse = data.content;
        } else if (typeof data === 'string') {
            assistantResponse = data;
        } else {
            assistantResponse = 'I received a response, but I\'m having trouble displaying it properly.';
        }

        // Add assistant message to the chat interface
        this.addAssistantMessage(assistantResponse);

        // Store the message
        if (this.currentChat) {
            this.currentChat.messages.push({
                role: 'assistant',
                content: assistantResponse,
                timestamp: new Date().toISOString()
            });
        }

    } catch (error) {
        console.error('API Error:', error);

        const errorResponse = `I apologize, but I encountered an error: ${error.message}. Please check your API settings and try again.`;
        this.addAssistantMessage(errorResponse, true);

        // Store error message
        if (this.currentChat) {
            this.currentChat.messages.push({
                role: 'assistant',
                content: errorResponse,
                timestamp: new Date().toISOString(),
                isError: true
            });
        }
    } finally {
        this.hideTypingIndicator();
        this.saveChats();
        this.updateConversationsList();
    }
}
// ...existing code...
    
    // Call REST API
    // async callAPI(userMessage) {
    //     try {
    //         if (!this.settings.apiUrl) {
    //             throw new Error('API endpoint not configured. Please configure it in settings.');
    //         }
            
    //         // Prepare request
    //         const headers = {
    //             'Content-Type': 'application/json'
    //         };
            
    //         if (this.settings.apiKey) {
    //             headers['Authorization'] = `Bearer ${this.settings.apiKey}`;
    //         }
            
    //         const requestBody = {
    //             message: userMessage,
    //             max_tokens: this.settings.maxTokens,
    //             conversation_history: this.currentChat ? this.currentChat.messages.slice(-10) : [] // Last 10 messages for context
    //         };
            
    //         const response = await fetch(this.settings.apiUrl, {
    //             method: 'POST',
    //             headers: headers,
    //             body: JSON.stringify(requestBody)
    //         });
            
    //         if (!response.ok) {
    //             throw new Error(`API request failed: ${response.status} ${response.statusText}`);
    //         }
            
    //         const data = await response.json();
            
    //         // Extract response text (adapt this based on your API response format)
    //         let assistantResponse = '';
    //         if (data.response) {
    //             assistantResponse = data.response;
    //         } else if (data.message) {
    //             assistantResponse = data.message;
    //         } else if (data.content) {
    //             assistantResponse = data.content;
    //         } else if (typeof data === 'string') {
    //             assistantResponse = data;
    //         } else {
    //             assistantResponse = 'I received a response, but I\'m having trouble displaying it properly.';
    //         }
            
    //         // Add assistant message to the chat interface
    //         this.addAssistantMessage(assistantResponse);
            
    //         // Store the message
    //         if (this.currentChat) {
    //             this.currentChat.messages.push({
    //                 role: 'assistant',
    //                 content: assistantResponse,
    //                 timestamp: new Date().toISOString()
    //             });
    //         }
            
    //     } catch (error) {
    //         console.error('API Error:', error);
            
    //         const errorResponse = `I apologize, but I encountered an error: ${error.message}. Please check your API settings and try again.`;
    //         this.addAssistantMessage(errorResponse, true);
            
    //         // Store error message
    //         if (this.currentChat) {
    //             this.currentChat.messages.push({
    //                 role: 'assistant',
    //                 content: errorResponse,
    //                 timestamp: new Date().toISOString(),
    //                 isError: true
    //             });
    //         }
    //     } finally {
    //         this.hideTypingIndicator();
    //         this.saveChats();
    //         this.updateConversationsList();
    //     }
    // }
    
    // Add assistant message to chat interface
    addAssistantMessage(message, isError = false) {
        const chatMessages = document.getElementById('chatMessages');
        if (chatMessages) {
            const assistantMessageEl = document.createElement('div');
            assistantMessageEl.className = `message assistant-message ${isError ? 'error' : ''}`;
            assistantMessageEl.innerHTML = `
                <div class="message-avatar">
                    <div class="copilot-logo">
                        <div class="logo-squares">
                            <div class="square red"></div>
                            <div class="square yellow"></div>
                            <div class="square green"></div>
                            <div class="square blue"></div>
                        </div>
                    </div>
                </div>
                <div class="message-content">
                    <div class="message-text">${this.formatMessageContent(message)}</div>
                    <div class="message-actions">
                        <button class="message-speak-btn" title="Read this message aloud" onclick="window.copilot.speakText('${this.escapeForAttribute(this.stripHtml(message))}')">
                            <i class="fas fa-volume-up"></i>
                        </button>
                        <div class="message-time">${this.formatTime(new Date().toISOString())}</div>
                    </div>
                </div>
            `;
            
            chatMessages.appendChild(assistantMessageEl);
            chatMessages.scrollTop = chatMessages.scrollHeight;
        }
    }
    
    // Update UI
    updateUI() {
        this.updateChatList();
        this.updateMessages();
        this.updateChatTitle();
    }
    
    // Update chat list in sidebar
    updateChatList() {
        const chatList = document.getElementById('chatList');
        chatList.innerHTML = '';
        
        const sortedChats = Object.values(this.chats).sort((a, b) => 
            new Date(b.updatedAt) - new Date(a.updatedAt)
        );
        
        if (sortedChats.length === 0) {
            chatList.innerHTML = '<p style="color: #8a8886; font-size: 14px; text-align: center; padding: 20px;">No chats yet. Start a new conversation!</p>';
            return;
        }
        
        sortedChats.forEach(chat => {
            const chatItem = document.createElement('div');
            chatItem.className = `chat-item ${chat.id === this.currentChatId ? 'active' : ''}`;
            
            const lastMessage = chat.messages[chat.messages.length - 1];
            const preview = lastMessage ? 
                (lastMessage.role === 'user' ? `You: ${lastMessage.content}` : lastMessage.content) : 
                'No messages yet';
            
            chatItem.innerHTML = `
                <div class="chat-item-content">
                    <div class="chat-item-title">${chat.title}</div>
                    <div class="chat-item-preview">${preview.substring(0, 60)}${preview.length > 60 ? '...' : ''}</div>
                    <div class="chat-item-date">${this.formatDate(chat.updatedAt)}</div>
                </div>
                <button class="chat-item-delete" title="Delete chat" data-chat-id="${chat.id}">
                    <i class="fas fa-trash"></i>
                </button>
            `;
            
            // Add click handler for the main chat area (excluding delete button)
            const chatContent = chatItem.querySelector('.chat-item-content');
            chatContent.addEventListener('click', () => this.switchToChat(chat.id));
            
            // Add delete button handler
            const deleteBtn = chatItem.querySelector('.chat-item-delete');
            deleteBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (confirm('Are you sure you want to delete this chat? This action cannot be undone.')) {
                    this.deleteChat(chat.id, e);
                }
            });
            
            // Keep right-click context menu as alternative
            chatItem.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                if (confirm('Delete this chat?')) {
                    this.deleteChat(chat.id, e);
                }
            });
            
            chatList.appendChild(chatItem);
        });
    }
    
    // Update messages in chat area
    updateMessages() {
        const messagesContainer = document.getElementById('messagesContainer');
        
        if (!this.currentChatId || !this.chats[this.currentChatId] || this.chats[this.currentChatId].messages.length === 0) {
            messagesContainer.innerHTML = `
                <div class="welcome-message">
                    <div class="welcome-icon">
                        <i class="fas fa-robot"></i>
                    </div>
                    <h2>How can I help you today?</h2>
                    <p>Ask me anything and I'll do my best to assist you. I can help with various tasks and answer your questions.</p>
                    
                    <div class="suggested-prompts">
                        <button class="prompt-suggestion" data-prompt="What's the weather like today?">
                            <i class="fas fa-cloud-sun"></i>
                            What's the weather like today?
                        </button>
                        <button class="prompt-suggestion" data-prompt="Help me write an email">
                            <i class="fas fa-envelope"></i>
                            Help me write an email
                        </button>
                        <button class="prompt-suggestion" data-prompt="Explain a complex topic">
                            <i class="fas fa-book"></i>
                            Explain a complex topic
                        </button>
                        <button class="prompt-suggestion" data-prompt="Generate code examples">
                            <i class="fas fa-code"></i>
                            Generate code examples
                        </button>
                    </div>
                </div>
            `;
            
            // Re-bind prompt suggestion events
            document.querySelectorAll('.prompt-suggestion').forEach(btn => {
                btn.addEventListener('click', () => {
                    const prompt = btn.dataset.prompt;
                    const messageInput = document.getElementById('messageInput');
                    messageInput.value = prompt;
                    this.updateCharacterCount();
                    this.updateSendButton();
                    messageInput.focus();
                });
            });
            
            return;
        }
        
        messagesContainer.innerHTML = '';
        
        this.chats[this.currentChatId].messages.forEach(message => {
            const messageEl = document.createElement('div');
            messageEl.className = `message ${message.role}`;
            
            const avatar = message.role === 'user' ? 
                '<i class="fas fa-user"></i>' : 
                '<i class="fas fa-robot"></i>';
            
            const contentClass = message.isError ? 'message-content error' : 'message-content';
            
            messageEl.innerHTML = `
                <div class="message-avatar">${avatar}</div>
                <div class="${contentClass}">
                    ${this.formatMessageContent(message.content)}
                    <div class="message-time">${this.formatTime(message.timestamp)}</div>
                </div>
            `;
            
            messagesContainer.appendChild(messageEl);
        });
        
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
    }
    
    // Update chat title
    updateChatTitle() {
        const chatTitle = document.getElementById('chatTitle');
        if (this.currentChatId && this.chats[this.currentChatId]) {
            chatTitle.textContent = this.chats[this.currentChatId].title;
        } else {
            chatTitle.textContent = 'New Chat';
        }
    }
    
    // Format message content (basic markdown-like formatting)
    formatMessageContent(content) {
        // Simple formatting - you can extend this
        return content
            .replace(/\n/g, '<br>')
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\*(.*?)\*/g, '<em>$1</em>')
            .replace(/`(.*?)`/g, '<code>$1</code>');
    }
    
    // Format date for chat list
    formatDate(dateString) {
        const date = new Date(dateString);
        const now = new Date();
        const diff = now - date;
        const days = Math.floor(diff / (1000 * 60 * 60 * 24));
        
        if (days === 0) return 'Today';
        if (days === 1) return 'Yesterday';
        if (days < 7) return `${days} days ago`;
        
        return date.toLocaleDateString();
    }
    
    // Format time for messages
    formatTime(dateString) {
        return new Date(dateString).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    }
    
    // Settings modal functions
    openSettings() {
        const modal = document.getElementById('settingsModal');
        document.getElementById('apiUrl').value = this.settings.apiUrl;
        document.getElementById('apiKey').value = this.settings.apiKey;
        document.getElementById('maxTokens').value = this.settings.maxTokens;
        modal.classList.add('active');
    }
    
    closeSettings() {
        const modal = document.getElementById('settingsModal');
        modal.classList.remove('active');
    }
    
    saveSettingsFromModal() {
        this.settings.apiUrl = document.getElementById('apiUrl').value.trim();
        this.settings.apiKey = document.getElementById('apiKey').value.trim();
        this.settings.maxTokens = parseInt(document.getElementById('maxTokens').value) || 1000;
        
        this.saveSettings();
        this.closeSettings();
        
        // Update API status
        const apiStatus = document.getElementById('apiStatus');
        if (this.settings.apiUrl) {
            apiStatus.textContent = 'Ready';
            apiStatus.className = 'api-status connected';
        } else {
            apiStatus.textContent = 'Not configured';
            apiStatus.className = 'api-status error';
        }
    }
    
    resetSettings() {
        if (confirm('Are you sure you want to reset all settings? This will clear your API configuration.')) {
            this.settings = {
                apiUrl: '',
                apiKey: '',
                maxTokens: 1000
                
            };
            this.saveSettings();
            
            document.getElementById('apiUrl').value = '';
            document.getElementById('apiKey').value = '';
            document.getElementById('maxTokens').value = '1000';
            
            const apiStatus = document.getElementById('apiStatus');
            apiStatus.textContent = 'Not configured';
            apiStatus.className = 'api-status error';
        }
    }
}

// Initialize the app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.copilot = new M365Copilot();
});

// Handle page visibility change to save state
document.addEventListener('visibilitychange', () => {
    if (window.copilot && document.visibilityState === 'hidden') {
        window.copilot.saveChats();
    }
});

// Auto-save periodically
setInterval(() => {
    if (window.copilot) {
        window.copilot.saveChats();
    }
}, 30000); // Save every 30 seconds
