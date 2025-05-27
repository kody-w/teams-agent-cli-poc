import { App } from '@microsoft/teams.apps';
import { DevtoolsPlugin } from '@microsoft/teams.dev';

// Configuration - matching the index.html setup
const AZURE_FUNCTION_URL = process.env.AZURE_FUNCTION_URL || 'https://azfbusinessbot.azurewebsites.net/api/businessinsightbot_function';
const FUNCTION_KEY = process.env.FUNCTION_KEY || '';

// Initialize Teams app
const app = new App({
  plugins: [new DevtoolsPlugin()],
});

// Conversation history and user GUID management
let conversationHistory: any[] = [];
let userGuid: string | null = null;

// Helper function to send message to Azure Function using native fetch
async function sendToAzureFunction(userInput: string, retries = 3): Promise<any> {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      console.log(`Attempt ${attempt}: Sending request to Azure Function...`);
      
      // Build conversation history in the expected format
      const payload = {
        user_input: userInput,
        conversation_history: conversationHistory,
        user_guid: userGuid
      };

      console.log('Payload:', JSON.stringify(payload, null, 2));

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), attempt === 1 ? 60000 : 120000);

      const response = await fetch(AZURE_FUNCTION_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-functions-key': FUNCTION_KEY,
        },
        body: JSON.stringify(payload),
        signal: controller.signal,
      });

      clearTimeout(timeoutId);
      
      console.log('Response received:', response.status);

      if (!response.ok) {
        if (response.status === 401) {
          throw new Error('Invalid function key. Please check your Azure Function key configuration.');
        }
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      return data;
    } catch (error: any) {
      console.error(`Attempt ${attempt} failed:`, error.message);
      
      if (error.message.includes('Invalid function key')) {
        throw error;
      }
      
      if (attempt === retries) {
        console.error('All retry attempts failed');
        throw error;
      }
      
      // Wait before retry (exponential backoff)
      const waitTime = attempt * 2000;
      console.log(`Waiting ${waitTime}ms before retry...`);
      await new Promise(resolve => setTimeout(resolve, waitTime));
    }
  }
}

// Main message handler
app.on('message', async ({ send, activity }) => {
  console.log('Received activity:', JSON.stringify(activity, null, 2));
  
  try {
    // Send typing indicator
    await send({ type: 'typing' });
    
    // Extract user message
    const userMessage = activity.text || '';
    const userId = activity.from?.id || 'unknown';
    
    // Skip empty messages
    if (!userMessage.trim()) {
      console.log('Empty message received, skipping...');
      return;
    }
    
    // If no user GUID exists, check if the message is a GUID
    const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!userGuid && guidPattern.test(userMessage.trim())) {
      userGuid = userMessage.trim();
      console.log('User GUID set to:', userGuid);
      
      // Send confirmation
      await send({
        type: 'message',
        text: `GUID set to: ${userGuid}`,
      });
    }
    
    // If still no GUID, generate one
    if (!userGuid) {
      userGuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        const r = (Math.random() * 16) | 0;
        const v = c === 'x' ? r : (r & 0x3) | 0x8;
        return v.toString(16);
      });
      console.log('Generated new GUID:', userGuid);
    }
    
    // Add user message to conversation history
    conversationHistory.push({
      role: 'user',
      content: userMessage
    });
    
    // Send initial message to indicate processing
    await send({
      type: 'message',
      text: '🤔 Processing your request...',
    });
    
    // Call Azure Function with retry logic
    const functionResponse = await sendToAzureFunction(userMessage);
    
    // Handle the response
    const assistantResponse = functionResponse.assistant_response || functionResponse.text || 'I received your message but got an empty response.';
    
    // Add assistant response to history
    conversationHistory.push({
      role: 'assistant',
      content: assistantResponse
    });
    
    // If agent logs exist, add them to history
    if (functionResponse.agent_logs) {
      conversationHistory.push({
        role: 'system',
        content: functionResponse.agent_logs
      });
    }
    
    // Update user GUID if returned
    if (functionResponse.user_guid && functionResponse.user_guid !== userGuid) {
      userGuid = functionResponse.user_guid;
      console.log('User GUID updated to:', userGuid);
    }
    
    // Send the response back to Teams
    await send({
      type: 'message',
      text: assistantResponse,
    });
    
    // If there are agent logs, send them as a separate message (optional)
    if (functionResponse.agent_logs) {
      // Format agent logs for better readability
      const formattedLogs = functionResponse.agent_logs
        .replace(/Performed (\w+)/, '🔧 **Agent: $1**')
        .replace(/and got result:/, '\n📊 **Result:**');
      
      await send({
        type: 'message',
        text: formattedLogs,
      });
    }
    
  } catch (error: any) {
    console.error('Error processing message:', error);
    
    let errorMessage = 'Sorry, I encountered an error processing your message.';
    
    if (error.name === 'AbortError' || error.message.includes('aborted')) {
      errorMessage = '⏱️ The request took too long to process. This might be due to:\n\n' +
        '• The Azure Function is warming up (first request)\n' +
        '• Complex processing is taking place\n' +
        '• Network connectivity issues\n\n' +
        'Please try again in a moment.';
    } else if (error.message.includes('Invalid function key')) {
      errorMessage = '🔑 Authentication failed. Please check your Azure Function key.';
    } else if (error.message.includes('500')) {
      errorMessage = '❌ The Azure Function encountered an internal error. Please try again later.';
    }
    
    // Send error message to user
    await send({
      type: 'message',
      text: errorMessage,
    });
  }
});

// Start the app
(async () => {
  const port = +(process.env.PORT || 3000);
  
  // Check if function key is configured
  if (!FUNCTION_KEY) {
    console.warn('⚠️  WARNING: No Azure Function key configured. Set FUNCTION_KEY environment variable.');
    console.warn('The app will not be able to communicate with the Azure Function without it.');
  } else {
    console.log('✅ Function key is configured');
  }
  
  // Warm up the Azure Function with a test request
  console.log('🔥 Warming up Azure Function...');
  try {
    await sendToAzureFunction('warm-up', 1);
    console.log('✅ Azure Function is responsive');
  } catch (error) {
    console.warn('⚠️  Azure Function warm-up failed. It might be cold starting.');
    console.warn('First user request might take longer than usual.');
  }
  
  await app.start(port);
  console.log(`\n🚀 Teams app started on port ${port}`);
  console.log(`📡 Azure Function URL: ${AZURE_FUNCTION_URL}`);
  console.log(`🔑 Function key configured: ${FUNCTION_KEY ? 'Yes' : 'No'}`);
  console.log(`\n💡 Tip: If you're experiencing timeouts, try sending your message again.`);
  console.log(`   Azure Functions can take 30-60 seconds to warm up after being idle.\n`);
})();