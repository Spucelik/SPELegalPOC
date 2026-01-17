import { 
  SHAREPOINT_CONFIG, 
  GRAPH_ENDPOINT, 
  GRAPH_BETA_ENDPOINT,
  COPILOT_SCOPES,
  GRAPH_SEARCH_SCOPES,
  IChatEmbeddedApiAuthProvider, 
  ChatLaunchConfig 
} from "@/config/sharepoint";

export type { ChatLaunchConfig };

export interface CopilotMessage {
  role: "user" | "assistant";
  content: string;
  timestamp: Date;
}

// Response type with optional citations
export interface CopilotResponse {
  content: string;
  citations?: Array<{
    documentName: string;
    webUrl: string;
    snippet?: string;
  }>;
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

/**
 * Create auth provider following SDK's IChatEmbeddedApiAuthProvider interface.
 * Uses SharePoint Container.Selected scope as per Microsoft documentation.
 * Falls back to Graph scopes if Container.Selected fails.
 */
export function createChatAuthProvider(
  getToken: (scopes: string[]) => Promise<string | null>
): IChatEmbeddedApiAuthProvider {
  return {
    hostname: SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME,
    getToken: async () => {
      // First try Container.Selected scope (official SDK requirement)
      try {
        const token = await getToken(COPILOT_SCOPES);
        if (token) {
          return token;
        }
      } catch (error) {
        console.warn("Container.Selected scope not available, falling back to Graph scopes:", error);
      }
      
      // Fallback to Graph scopes for search-based functionality
      const fallbackToken = await getToken(GRAPH_SEARCH_SCOPES);
      if (!fallbackToken) {
        throw new Error("Failed to acquire token for Copilot chat");
      }
      return fallbackToken;
    },
  };
}

/**
 * Send a message to Copilot using the beta Copilot retrieval API.
 * Falls back to Graph Search if the beta API is not available.
 */
export async function sendCopilotMessage(
  authProvider: IChatEmbeddedApiAuthProvider,
  containerId: string,
  containerName: string,
  userMessage: string,
  conversationHistory: CopilotMessage[],
  config: ChatLaunchConfig = DEFAULT_CHAT_CONFIG
): Promise<string> {
  const accessToken = await authProvider.getToken();

  // Build conversation context for the Copilot API
  const contextMessages = conversationHistory
    .slice(-6)
    .map((m) => ({
      role: m.role,
      content: m.content,
    }));

  const systemInstruction = config.instruction || DEFAULT_CHAT_CONFIG.instruction;

  // Try the beta Copilot retrieval API first
  try {
    const copilotResponse = await callCopilotRetrievalAPI(
      accessToken,
      containerId,
      containerName,
      userMessage,
      contextMessages,
      systemInstruction
    );
    
    if (copilotResponse) {
      return copilotResponse;
    }
  } catch (error) {
    console.warn("Beta Copilot API not available, falling back to Graph Search:", error);
  }

  // Fallback to Graph Search API
  return await searchBasedResponse(accessToken, containerId, containerName, userMessage, config);
}

/**
 * Call the Microsoft Graph beta Copilot retrieval API.
 * This provides AI-generated responses grounded in container documents.
 */
async function callCopilotRetrievalAPI(
  accessToken: string,
  containerId: string,
  containerName: string,
  userMessage: string,
  contextMessages: Array<{ role: string; content: string }>,
  systemInstruction: string
): Promise<string | null> {
  const copilotUrl = `${GRAPH_BETA_ENDPOINT}/copilot/retrieval`;

  const requestBody = {
    requests: [
      {
        entityTypes: ["driveItem"],
        contentSources: [`/drives/${containerId}`],
        query: {
          queryString: userMessage,
        },
        groundingOptions: {
          systemPrompt: systemInstruction,
          conversationContext: contextMessages,
        },
      },
    ],
  };

  const response = await fetch(copilotUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error("Copilot retrieval API error:", response.status, errorText);
    return null;
  }

  const data = await response.json();
  
  // Extract response content from the Copilot API response
  const responseContent = data.value?.response || data.value?.[0]?.response;
  
  if (responseContent) {
    return cleanCopilotText(responseContent);
  }

  // Check for any other response format
  if (data.value?.content) {
    return cleanCopilotText(data.value.content);
  }

  return null;
}

/**
 * Search-based response using Graph Search API.
 * Scoped to the specific container (drive) to only search documents in the selected case.
 */
async function searchBasedResponse(
  accessToken: string,
  containerId: string,
  containerName: string,
  userMessage: string,
  config: ChatLaunchConfig
): Promise<string> {
  const searchUrl = `${GRAPH_ENDPOINT}/search/query`;

  // Use the container ID as the drive ID to scope search to this specific case
  // SharePoint Embedded containers ARE Graph drives, so we can filter by driveId
  const requestBody = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: {
          // Scope to specific container/drive using the container ID
          queryString: userMessage,
        },
        // Restrict search to the specific drive (container)
        sharePointOneDriveOptions: {
          includeContent: "privateContent",
        },
        // Filter to only this container's drive
        contentSources: [`/drives/${containerId}`],
        from: 0,
        size: 10,
      },
    ],
  };

  try {
    const searchResponse = await fetch(searchUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });

    if (!searchResponse.ok) {
      const errorText = await searchResponse.text();
      console.error("Search failed:", searchResponse.status, errorText);
      
      // Try alternative approach: search with drive filter in query
      return await searchWithDriveFilter(accessToken, containerId, containerName, userMessage);
    }

    const searchData = await searchResponse.json();
    const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];

    if (hits.length === 0) {
      // Try alternative approach before giving up
      return await searchWithDriveFilter(accessToken, containerId, containerName, userMessage);
    }

    // Build response from search results with extracts
    return formatSearchResults(hits, containerName);
  } catch (error) {
    console.error("Search error:", error);
    return "I'm having trouble accessing the case documents right now. Please try again in a moment.";
  }
}

/**
 * Alternative search approach: directly query the container's drive for content.
 */
async function searchWithDriveFilter(
  accessToken: string,
  containerId: string,
  containerName: string,
  userMessage: string
): Promise<string> {
  // Use the drive's search endpoint directly to scope to this container
  const driveSearchUrl = `${GRAPH_ENDPOINT}/drives/${containerId}/root/search(q='${encodeURIComponent(userMessage)}')`;
  
  try {
    const response = await fetch(driveSearchUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      console.error("Drive search failed:", await response.text());
      return getNoResultsMessage(containerName, userMessage);
    }

    const data = await response.json();
    const items = data.value || [];

    if (items.length === 0) {
      return getNoResultsMessage(containerName, userMessage);
    }

    // Format results from drive search
    const responses: string[] = items.slice(0, 5).map((item: any) => {
      const name = item.name || 'Document';
      const description = item.description || item.webUrl || '';
      return `**${name}**${description ? `\n${description}` : ''}`;
    });

    if (responses.length > 0) {
      return `Found ${items.length} document(s) in the ${containerName} case matching your query:\n\n${responses.join("\n\n")}\n\nWould you like more details about any of these documents?`;
    }

    return getNoResultsMessage(containerName, userMessage);
  } catch (error) {
    console.error("Drive search error:", error);
    return getNoResultsMessage(containerName, userMessage);
  }
}

/**
 * Format search results into a readable response.
 */
function formatSearchResults(hits: any[], containerName: string): string {
  const responses: string[] = [];

  for (const hit of hits.slice(0, 5)) {
    if (hit.extracts && hit.extracts.length > 0) {
      const extractText = cleanCopilotText(hit.extracts[0].text);
      if (extractText) {
        responses.push(`**${hit.resource?.name || 'Document'}:**\n${extractText}`);
      }
    } else if (hit.summary) {
      responses.push(`**${hit.resource?.name || 'Document'}:**\n${cleanCopilotText(hit.summary)}`);
    }
  }

  if (responses.length > 0) {
    return `Based on documents in the ${containerName} case:\n\n${responses.join("\n\n")}`;
  }

  return getNoResultsMessage(containerName, "your query");
}

/**
 * Generate a helpful message when no results are found.
 */
function getNoResultsMessage(containerName: string, userMessage: string): string {
  return `I couldn't find specific information about "${userMessage}" in the ${containerName} case documents. Try:
• Rephrasing your question with different keywords
• Asking about specific document names or topics
• Checking if the documents have been uploaded to this case`;
}