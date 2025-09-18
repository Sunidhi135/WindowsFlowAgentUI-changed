# Copilot Chat - Microsoft 365 Copilot-inspired UI

A modern, responsive chat interface inspired by Microsoft 365 Copilot, featuring chat history management and REST API integration.

## Features

- **Microsoft 365 Copilot-inspired Design**: Clean, modern interface with familiar styling
- **Chat History Management**: Left sidebar showing recent conversations with persistence
- **REST API Integration**: Configurable endpoint for connecting to your chat API
- **Real-time Chat**: Send messages and receive responses with typing indicators
- **Local Storage**: Automatic saving of chat history and settings
- **Responsive Design**: Works on desktop, tablet, and mobile devices
- **Settings Panel**: Configure API endpoint, authentication, and response limits
- **Message Formatting**: Basic markdown-like formatting support
- **Suggested Prompts**: Quick-start prompts for new conversations

## Getting Started

1. **Open the Application**
   - Simply open `index.html` in your web browser
   - No server setup required for the basic interface

2. **Configure API Settings**
   - Click the settings icon (⚙️) in the top-right corner
   - Enter your REST API endpoint URL
   - Optionally add an API key for authentication
   - Set the maximum response length
   - Click "Save" to store your settings

3. **Start Chatting**
   - Type your message in the input field at the bottom
   - Press Enter or click the send button to send your message
   - The app will call your configured API and display the response

## API Configuration

### Expected API Request Format

The app sends POST requests to your configured endpoint with the following JSON structure:

```json
{
  "message": "User's message text",
  "max_tokens": 1000,
  "conversation_history": [
    {
      "id": "msg_123",
      "role": "user",
      "content": "Previous message",
      "timestamp": "2024-01-01T12:00:00.000Z"
    }
  ]
}
```

### Expected API Response Format

Your API should respond with JSON in one of these formats:

```json
// Option 1
{
  "response": "Assistant's response text"
}

// Option 2
{
  "message": "Assistant's response text"
}

// Option 3
{
  "content": "Assistant's response text"
}

// Option 4 - Simple string
"Assistant's response text"
```

### Authentication

If you provide an API key in settings, it will be sent as:
```
Authorization: Bearer YOUR_API_KEY
```

## File Structure

```
mshack/
├── index.html      # Main HTML structure
├── styles.css      # Microsoft 365 Copilot-inspired styling
├── script.js       # JavaScript functionality
└── README.md       # This file
```

## Browser Compatibility

- Chrome 70+
- Firefox 65+
- Safari 12+
- Edge 79+

## Features Overview

### Left Sidebar
- **New Chat Button**: Start fresh conversations
- **Chat History**: View and switch between recent chats
- **User Profile**: Simple user indicator

### Main Chat Area
- **Welcome Screen**: Suggested prompts for new users
- **Message Display**: Clean message bubbles with timestamps
- **Typing Indicator**: Shows when the assistant is processing
- **Input Field**: Auto-resizing text area with character counter

### Settings Panel
- **API Configuration**: Set endpoint URL and authentication
- **Response Limits**: Control maximum response length
- **Reset Option**: Clear all settings

## Local Storage

The app automatically saves:
- Chat conversations and messages
- API configuration settings
- Current active chat selection

Data persists across browser sessions and is automatically saved every 30 seconds.

## Customization

### Styling
Modify `styles.css` to change colors, fonts, or layout. The design uses CSS custom properties for easy theming.

### API Integration
Update the `callAPI` method in `script.js` to match your specific API requirements.

### Message Formatting
Extend the `formatMessageContent` method to add more markdown or HTML formatting support.

## Troubleshooting

### API Not Working
1. Check that your API URL is correct in settings
2. Verify your API accepts POST requests with JSON
3. Check browser console for error messages
4. Ensure CORS is properly configured on your API

### Chat History Not Saving
1. Check if local storage is enabled in your browser
2. Try clearing browser cache and cookies
3. Ensure you're not in private/incognito mode

### Mobile Issues
1. The interface is responsive, but some features may be limited on very small screens
2. Try rotating your device to landscape mode for better chat history visibility

## Development

To modify or extend the application:

1. **HTML Structure**: Edit `index.html` for layout changes
2. **Styling**: Modify `styles.css` for visual customization
3. **Functionality**: Update `script.js` for feature additions

The code is modular and well-commented for easy understanding and modification.

## License

This project is open source and available under the MIT License.
