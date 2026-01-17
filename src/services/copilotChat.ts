import { 
  SHAREPOINT_CONFIG, 
  COPILOT_SCOPES,
  IChatEmbeddedApiAuthProvider, 
  ChatLaunchConfig 
} from "@/config/sharepoint";

export type { ChatLaunchConfig };

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

/**
 * Create auth provider for SharePoint Embedded Copilot SDK.
 * 
 * Uses Container.Selected scope as required by the official SDK.
 * This scope provides access to Copilot features scoped to the specific container.
 */
export function createChatAuthProvider(
  getToken: (scopes: string[]) => Promise<string | null>
): IChatEmbeddedApiAuthProvider {
  return {
    hostname: SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME,
    getToken: async () => {
      // SDK requires Container.Selected scope for SharePoint Embedded
      const token = await getToken(COPILOT_SCOPES);
      if (!token) {
        throw new Error("Failed to acquire Container.Selected token for Copilot");
      }
      console.log("Acquired Container.Selected token for Copilot SDK");
      return token;
    },
  };
}
