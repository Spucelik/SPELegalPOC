import { useMemo, useState, useEffect, useCallback } from "react";
import { Loader2, AlertTriangle, RefreshCw, ExternalLink } from "lucide-react";
import { useAuth } from "@/contexts/AuthContext";
import { Button } from "@/components/ui/button";
import { CopilotAuthProvider } from "@/components/copilot/CopilotAuthProvider";
import { CopilotErrorBoundary } from "@/components/copilot/CopilotErrorBoundary";
import { ChatEmbedded, ChatEmbeddedAPI } from "@microsoft/sharepointembedded-copilotchat-react";

interface CopilotPanelProps {
  containerId: string;
  containerName: string;
}

export default function CopilotPanel({ containerId, containerName }: CopilotPanelProps) {
  const { getAccessToken } = useAuth();
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [chatApi, setChatApi] = useState<ChatEmbeddedAPI | null>(null);

  const authProvider = useMemo(
    () => new CopilotAuthProvider(getAccessToken),
    [getAccessToken]
  );

  useEffect(() => {
    const initAuth = async () => {
      setIsLoading(true);
      setError(null);

      try {
        await authProvider.initialize();
        console.log("CopilotPanel: Auth provider initialized");
        setIsLoading(false);
      } catch (err) {
        console.error("CopilotPanel: Auth initialization failed", err);
        setError("Authentication failed. Please try again.");
        setIsLoading(false);
      }
    };

    const timeout = setTimeout(() => {
      if (isLoading) {
        setError(
          "The SharePoint Embedded Copilot chat is not responding.\n\n" +
          "Possible causes:\n" +
          "• CopilotEmbeddedChatHosts not configured\n" +
          "• DiscoverabilityDisabled is true\n" +
          "• Copilot not enabled for your tenant"
        );
        setIsLoading(false);
      }
    }, 15000);

    initAuth();

    return () => clearTimeout(timeout);
  }, [containerId, authProvider]);

  useEffect(() => {
    if (!chatApi) return;

    const openChat = async () => {
      try {
        await chatApi.openChat({
          header: `Case Assistant - ${containerName}`,
          zeroQueryPrompts: {
            headerText: "How can I help you with this case?",
            promptSuggestionList: [
              { suggestionText: "Summarize the key facts of this case" },
              { suggestionText: "Who are the parties involved?" },
              { suggestionText: "What are the important dates?" },
              { suggestionText: "List the key documents" },
            ],
          },
          instruction:
            "You are a legal case assistant. Provide clear, professional responses based on the case documents.",
          locale: "en",
        });
        console.log("CopilotPanel: Chat opened successfully");
      } catch (err) {
        console.error("CopilotPanel: Failed to open chat", err);
        setError("Failed to open chat interface");
      }
    };

    openChat();
  }, [chatApi, containerName]);

  const handleApiReady = useCallback((api: ChatEmbeddedAPI) => {
    console.log("CopilotPanel: API ready");
    setChatApi(api);
    setIsLoading(false);
    setError(null);
  }, []);

  const handleRetry = useCallback(() => {
    setError(null);
    setIsLoading(true);
    setChatApi(null);
  }, []);

  if (error) {
    return (
      <div className="flex flex-col items-center justify-center h-full p-6 text-center">
        <div className="p-3 rounded-full bg-destructive/10 mb-3">
          <AlertTriangle className="w-6 h-6 text-destructive" />
        </div>
        <h4 className="font-semibold text-sm mb-2">Copilot Unavailable</h4>
        <p className="text-xs text-muted-foreground whitespace-pre-wrap mb-4 max-w-sm">
          {error}
        </p>
        <Button variant="outline" size="sm" onClick={handleRetry}>
          <RefreshCw className="w-4 h-4 mr-2" />
          Retry
        </Button>
        <a 
          href="https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/declarative-agent/spe-da-adv"
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center gap-1 mt-4 text-xs text-primary hover:underline"
        >
          View documentation
          <ExternalLink className="w-3 h-3" />
        </a>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="flex flex-col items-center justify-center h-full">
        <Loader2 className="w-6 h-6 animate-spin text-primary mb-3" />
        <p className="text-sm text-muted-foreground">Initializing Copilot...</p>
        <p className="text-xs text-muted-foreground mt-1">Connecting to Microsoft services</p>
      </div>
    );
  }

  return (
    <CopilotErrorBoundary onRetry={handleRetry}>
      <div className="w-full h-full min-h-[400px]">
        <ChatEmbedded
          onApiReady={handleApiReady}
          authProvider={authProvider}
          containerId={containerId}
          style={{ width: '100%', height: '100%', minHeight: '400px' }}
        />
      </div>
    </CopilotErrorBoundary>
  );
}
