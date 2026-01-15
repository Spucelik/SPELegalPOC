import { SHAREPOINT_CONFIG, GRAPH_ENDPOINT, IChatEmbeddedApiAuthProvider, ChatLaunchConfig } from "@/config/sharepoint";

export type { ChatLaunchConfig };

export interface CopilotMessage {
  role: "user" | "assistant";
  content: string;
  timestamp: Date;
}

// Default launch configuration following SDK patterns
export const DEFAULT_CHAT_CONFIG: ChatLaunchConfig = {
  header: "Case Assistant",
  zeroQueryPrompts: {
    headerText: "How can I help you with this case?",
    promptSuggestionList: [
      { suggestionText: "Summarize the key facts of this case" },
      { suggestionText: "Who are the parties involved?" },
      { suggestionText: "What are the important dates?" },
      { suggestionText: "List the key documents" },
    ],
  },
  suggestedPrompts: [
    "What are the main legal issues?",
    "Summarize the evidence",
    "What is the current status?",
  ],
  instruction: "You are a legal case assistant. Provide clear, professional responses based on the case documents.",
  locale: "en",
};

// Clean up text from Copilot API responses
function cleanCopilotText(text: string): string {
  let cleaned = text;
  // Remove page markers
  cleaned = cleaned.replace(/<page_\d+>/g, '').replace(/<\/page_\d+>/g, '');
  // Remove escaped markdown characters
  cleaned = cleaned.replace(/\\_/g, '_').replace(/\\-/g, '-');
  cleaned = cleaned.replace(/\\\[/g, '[').replace(/\\\]/g, ']');
  cleaned = cleaned.replace(/\\\(/g, '(').replace(/\\\)/g, ')');
  cleaned = cleaned.replace(/\\\*/g, '*');
  // Remove standalone asterisks used as separators
  cleaned = cleaned.replace(/(\s*\*\s*){2,}/g, ' ');
  // Remove single asterisks at word boundaries
  cleaned = cleaned.replace(/\*+/g, '');
  // Remove backslashes before common characters
  cleaned = cleaned.replace(/\\([^\\])/g, '$1');
  // Clean up whitespace
  cleaned = cleaned.replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
  return cleaned;
}

// Create auth provider following SDK's IChatEmbeddedApiAuthProvider interface
export function createChatAuthProvider(
  getToken: (scopes: string[]) => Promise<string | null>
): IChatEmbeddedApiAuthProvider {
  return {
    hostname: SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME,
    getToken: async () => {
      // Request token with Container.Selected scope as per SDK documentation
      const scope = `${SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME}/Container.Selected`;
      const token = await getToken([scope]);
      if (!token) {
        throw new Error("Failed to acquire token for Copilot chat");
      }
      return token;
    },
  };
}

export async function sendCopilotMessage(
  authProvider: IChatEmbeddedApiAuthProvider,
  containerId: string,
  containerName: string,
  userMessage: string,
  conversationHistory: CopilotMessage[],
  config: ChatLaunchConfig = DEFAULT_CHAT_CONFIG
): Promise<string> {
  // Get token using the auth provider
  const accessToken = await authProvider.getToken();

  // Build context from conversation history
  const contextMessages = conversationHistory
    .slice(-6) // Keep last 6 messages for context
    .map((m) => `${m.role === "user" ? "User" : "Assistant"}: ${m.content}`)
    .join("\n");

  // Build system instruction from config
  const systemInstruction = config.instruction || DEFAULT_CHAT_CONFIG.instruction;

  // Create a search query that includes the user's question and context
  const queryString = contextMessages 
    ? `${systemInstruction}\n\nConversation context:\n${contextMessages}\n\nUser's question: ${userMessage}`
    : `${systemInstruction}\n\nUser's question: ${userMessage}`;

  const searchUrl = `${GRAPH_ENDPOINT}/search/query`;

  const requestBody = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: {
          queryString: queryString,
        },
        sharePointOneDriveOptions: {
          includeHiddenContent: false,
        },
        enableTopResults: true,
        from: 0,
        size: 10,
        queryAlterationOptions: {
          enableSuggestion: true,
          enableModification: true,
        },
        // Filter to the specific container
        contentSources: [`/drives/${containerId}`],
        fields: [
          "name",
          "webUrl",
          "lastModifiedDateTime",
          "createdBy",
        ],
      },
    ],
  };

  try {
    // First, try to get relevant documents using Graph API
    const searchResponse = await fetch(searchUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });

    if (!searchResponse.ok) {
      console.error("Search failed:", await searchResponse.text());
      // Fall back to direct query
      return await queryCopilotDirectly(accessToken, containerName, userMessage, config);
    }

    const searchData = await searchResponse.json();
    const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];

    if (hits.length === 0) {
      // No results from search, try direct query
      return await queryCopilotDirectly(accessToken, containerName, userMessage, config);
    }

    // Build response from search results
    const responses: string[] = [];
    
    for (const hit of hits.slice(0, 5)) {
      if (hit.extracts && hit.extracts.length > 0) {
        const extractText = cleanCopilotText(hit.extracts[0].text);
        if (extractText) {
          responses.push(`From "${hit.resource?.name || 'document'}":\n${extractText}`);
        }
      } else if (hit.summary) {
        responses.push(`From "${hit.resource?.name || 'document'}":\n${cleanCopilotText(hit.summary)}`);
      }
    }

    if (responses.length > 0) {
      return responses.join("\n\n");
    }

    // If no extracts, try direct query
    return await queryCopilotDirectly(accessToken, containerName, userMessage, config);
  } catch (error) {
    console.error("Chat error:", error);
    return await queryCopilotDirectly(accessToken, containerName, userMessage, config);
  }
}

async function queryCopilotDirectly(
  accessToken: string,
  containerName: string,
  userMessage: string,
  config: ChatLaunchConfig
): Promise<string> {
  const searchUrl = `${GRAPH_ENDPOINT}/search/query`;

  const requestBody = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: {
          queryString: `${userMessage} containerTypeId:${SHAREPOINT_CONFIG.CONTAINER_TYPE_ID}`,
        },
        sharePointOneDriveOptions: {
          includeHiddenContent: false,
        },
        enableTopResults: true,
        from: 0,
        size: 5,
      },
    ],
  };

  try {
    const response = await fetch(searchUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      throw new Error(`Search failed: ${response.status}`);
    }

    const data = await response.json();
    const hits = data.value?.[0]?.hitsContainers?.[0]?.hits || [];

    if (hits.length === 0) {
      return `I couldn't find specific information about "${userMessage}" in the ${containerName} case documents. Try rephrasing your question or asking about a different topic.`;
    }

    const responses: string[] = [];
    
    for (const hit of hits.slice(0, 3)) {
      if (hit.extracts && hit.extracts.length > 0) {
        const extractText = cleanCopilotText(hit.extracts[0].text);
        if (extractText) {
          responses.push(`From "${hit.resource?.name || 'document'}":\n${extractText}`);
        }
      }
    }

    if (responses.length > 0) {
      return responses.join("\n\n");
    }

    return `I found some documents that might be relevant, but couldn't extract specific information. Try asking a more specific question about the ${containerName} case.`;
  } catch (error) {
    console.error("Copilot query error:", error);
    return "I'm having trouble accessing the case documents right now. Please try again in a moment.";
  }
}