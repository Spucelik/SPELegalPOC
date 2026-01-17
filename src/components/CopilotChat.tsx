import { useState, useCallback, useMemo, useRef, useEffect } from "react";
import { MessageCircle, X } from "lucide-react";
import { cn } from "@/lib/utils";
import { useAuth } from "@/contexts/AuthContext";
import { ChatEmbedded, ChatEmbeddedAPI, type ChatLaunchConfig, type IChatEmbeddedApiAuthProvider } from "@microsoft/sharepointembedded-copilotchat-react";
import { createTheme, ThemeProvider } from "@fluentui/react";
import { SHAREPOINT_CONFIG, COPILOT_SCOPES } from "@/config/sharepoint";

interface CopilotChatProps {
  containerId: string;
  containerName: string;
  config?: ChatLaunchConfig;
}

// Default chat configuration
const DEFAULT_CHAT_CONFIG: ChatLaunchConfig = {
  header: "Case Assistant",
  zeroQueryPrompts: {
    headerText: "How can I help you with this case?",
    promptSuggestionList: [
      { suggestionText: "Summarize the key facts of this case" },
      { suggestionText: "What are the important deadlines?" },
      { suggestionText: "List all parties involved" },
    ],
  },
  suggestedPrompts: [
    "Find relevant documents",
    "Analyze case timeline",
  ],
  chatInputPlaceholder: "Ask about this case...",
};

// Create a Fluent UI theme that matches our app's design
const fluentTheme = createTheme({
  palette: {
    themePrimary: '#0f172a',
    themeLighterAlt: '#f3f4f6',
    themeLighter: '#e5e7eb',
    themeLight: '#d1d5db',
    themeTertiary: '#9ca3af',
    themeSecondary: '#6b7280',
    themeDarkAlt: '#1e293b',
    themeDark: '#334155',
    themeDarker: '#475569',
    neutralLighterAlt: '#fafafa',
    neutralLighter: '#f5f5f5',
    neutralLight: '#eaeaea',
    neutralQuaternaryAlt: '#d4d4d4',
    neutralQuaternary: '#c8c8c8',
    neutralTertiaryAlt: '#a3a3a3',
    neutralTertiary: '#737373',
    neutralSecondary: '#525252',
    neutralPrimaryAlt: '#404040',
    neutralPrimary: '#171717',
    neutralDark: '#0a0a0a',
    black: '#000000',
    white: '#ffffff',
  },
});

export default function CopilotChat({ 
  containerId, 
  containerName, 
  config = DEFAULT_CHAT_CONFIG 
}: CopilotChatProps) {
  const { getAccessToken } = useAuth();
  const [isOpen, setIsOpen] = useState(false);
  const [chatApi, setChatApi] = useState<ChatEmbeddedAPI | null>(null);
  const hasOpenedChat = useRef(false);

  // Create auth provider following SDK's IChatEmbeddedApiAuthProvider interface
  const authProvider: IChatEmbeddedApiAuthProvider = useMemo(() => ({
    hostname: SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME,
    getToken: async (): Promise<string> => {
      const token = await getAccessToken(COPILOT_SCOPES);
      if (!token) {
        throw new Error("Failed to acquire access token for Copilot");
      }
      return token;
    },
  }), [getAccessToken]);

  // Merged config with defaults
  const chatConfig = useMemo(() => ({
    ...DEFAULT_CHAT_CONFIG,
    ...config,
    header: config?.header || containerName,
  }), [config, containerName]);

  // Handle API ready callback
  const handleApiReady = useCallback((api: ChatEmbeddedAPI) => {
    console.log("ChatEmbedded API ready");
    setChatApi(api);
  }, []);

  // Open chat when API is ready and panel is opened
  useEffect(() => {
    if (chatApi && isOpen && !hasOpenedChat.current) {
      hasOpenedChat.current = true;
      chatApi.openChat(chatConfig);
    }
  }, [chatApi, isOpen, chatConfig]);

  // Reset state when container changes
  useEffect(() => {
    setChatApi(null);
    hasOpenedChat.current = false;
    setIsOpen(false);
  }, [containerId]);

  // Handle chat close
  const handleChatClose = useCallback((data: object) => {
    console.log("Chat closed", data);
    setIsOpen(false);
    hasOpenedChat.current = false;
  }, []);

  // Handle notifications
  const handleNotification = useCallback((data: object) => {
    console.log("Chat notification", data);
  }, []);

  // Don't render if no container is selected
  if (!containerId) return null;

  return (
    <>
      {/* Chat Bubble Button */}
      <button
        onClick={() => setIsOpen(!isOpen)}
        className={cn(
          "fixed bottom-6 right-6 z-50 flex items-center justify-center",
          "w-14 h-14 rounded-full shadow-lg transition-all duration-300",
          "bg-primary hover:bg-primary/90 text-primary-foreground",
          "hover:scale-105 active:scale-95",
          isOpen && "rotate-90"
        )}
        aria-label={isOpen ? "Close chat" : "Open Copilot chat"}
      >
        {isOpen ? (
          <X className="w-6 h-6" />
        ) : (
          <MessageCircle className="w-6 h-6" />
        )}
      </button>

      {/* Chat Flyout Panel */}
      <div
        className={cn(
          "fixed bottom-24 right-6 z-40 w-[400px] max-w-[calc(100vw-3rem)]",
          "bg-card border border-border rounded-xl shadow-2xl",
          "flex flex-col overflow-hidden transition-all duration-300",
          isOpen
            ? "opacity-100 translate-y-0 pointer-events-auto"
            : "opacity-0 translate-y-4 pointer-events-none"
        )}
        style={{ height: "550px", maxHeight: "calc(100vh - 150px)" }}
      >
        <ThemeProvider theme={fluentTheme}>
          <ChatEmbedded
            authProvider={authProvider}
            containerId={containerId}
            onApiReady={handleApiReady}
            onChatClose={handleChatClose}
            onNotification={handleNotification}
            themeV8={fluentTheme}
            style={{ 
              width: "100%", 
              height: "100%",
              border: "none",
            }}
          />
        </ThemeProvider>
      </div>
    </>
  );
}
