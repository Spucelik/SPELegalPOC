import React, { useMemo, useState, useEffect, useCallback, useRef } from "react";
import { Loader2, AlertTriangle, RefreshCw, ExternalLink, Send, Database, CheckCircle2 } from "lucide-react";
import { useAuth } from "@/contexts/AuthContext";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { cn } from "@/lib/utils";
import { CopilotAuthProvider, CopilotErrorBoundary, CopilotDesktopView } from "@/components/copilot";
import { ChatEmbeddedAPI, ChatLaunchConfig } from "@microsoft/sharepointembedded-copilotchat-react";
import { 
  sendCopilotMessage, 
  createChatAuthProvider, 
  CopilotMessage,
  DEFAULT_CHAT_CONFIG 
} from "@/services/copilotChat";
import { SHAREPOINT_CONFIG } from "@/config/sharepoint";

interface CopilotPanelProps {
  containerId: string;
  containerName: string;
}

/**
 * Verify container metadata is accessible before mounting SDK.
 */
async function verifyContainerMetadata(
  containerId: string,
  getAccessToken: (scopes: string[]) => Promise<string | null>
): Promise<{ valid: boolean; driveId?: string; error?: string }> {
  try {
    const token = await getAccessToken([
      "https://graph.microsoft.com/Files.Read.All",
      "https://graph.microsoft.com/Sites.Read.All",
    ]);
    
    if (!token) {
      return { valid: false, error: "Unable to acquire Graph token" };
    }

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${containerId}`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
      }
    );

    if (!response.ok) {
      return { valid: false, error: `Container not accessible: ${response.status}` };
    }

    const data = await response.json();
    console.log("CopilotPanel: Container metadata verified", data.displayName);
    
    return { valid: true, driveId: data.id };
  } catch (err) {
    console.error("CopilotPanel: Container verification failed", err);
    return { valid: false, error: "Failed to verify container access" };
  }
}

/**
 * Fallback chat component using Graph API when SDK fails
 */
const FallbackChat = React.forwardRef<HTMLDivElement, CopilotPanelProps>(
  function FallbackChat({ containerId, containerName }, ref) {
  const { getAccessToken } = useAuth();
  const [messages, setMessages] = useState<CopilotMessage[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isConnected, setIsConnected] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);

  const authProvider = useMemo(() => 
    createChatAuthProvider(getAccessToken), 
    [getAccessToken]
  );

  useEffect(() => {
    const testConnection = async () => {
      try {
        await authProvider.getToken();
        setIsConnected(true);
      } catch {
        setIsConnected(false);
      }
    };
    testConnection();
  }, [authProvider]);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [messages]);

  const handleSendMessage = useCallback(async (text: string) => {
    if (!text.trim() || isLoading) return;

    const userMessage: CopilotMessage = {
      role: "user",
      content: text.trim(),
      timestamp: new Date(),
    };

    setMessages(prev => [...prev, userMessage]);
    setInputValue("");
    setIsLoading(true);

    try {
      const response = await sendCopilotMessage(
        authProvider,
        containerId,
        containerName,
        text.trim(),
        messages,
        DEFAULT_CHAT_CONFIG
      );

      const assistantMessage: CopilotMessage = {
        role: "assistant",
        content: response,
        timestamp: new Date(),
      };

      setMessages(prev => [...prev, assistantMessage]);
    } catch (error) {
      console.error("Chat error:", error);
      const errorMessage: CopilotMessage = {
        role: "assistant",
        content: "I'm sorry, I encountered an error processing your request. Please try again.",
        timestamp: new Date(),
      };
      setMessages(prev => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  }, [authProvider, containerId, containerName, messages, isLoading]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    handleSendMessage(inputValue);
  };

  const zeroQueryPrompts = DEFAULT_CHAT_CONFIG.zeroQueryPrompts;

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="flex items-center gap-2 pb-3 border-b border-border mb-3">
        <Database className="w-4 h-4 text-muted-foreground" />
        <span className="text-sm text-muted-foreground">{containerName}</span>
        {isConnected && (
          <span className="flex items-center gap-1 text-xs text-primary">
            <CheckCircle2 className="w-3 h-3" />
            Connected
          </span>
        )}
      </div>

      {/* Messages */}
      <ScrollArea className="flex-1 pr-2" ref={scrollRef}>
        {messages.length === 0 ? (
          <div className="space-y-4">
            <p className="text-sm text-muted-foreground text-center">
              {zeroQueryPrompts?.headerText}
            </p>
            <div className="space-y-2">
              {zeroQueryPrompts?.promptSuggestionList?.map((prompt, index) => (
                <button
                  key={index}
                  onClick={() => handleSendMessage(prompt.suggestionText)}
                  className="w-full text-left px-3 py-2 text-sm rounded-lg 
                           bg-muted hover:bg-muted/80 transition-colors
                           border border-border hover:border-primary/50"
                >
                  {prompt.suggestionText}
                </button>
              ))}
            </div>
          </div>
        ) : (
          <div className="space-y-3">
            {messages.map((message, index) => (
              <div
                key={index}
                className={cn(
                  "flex",
                  message.role === "user" ? "justify-end" : "justify-start"
                )}
              >
                <div
                  className={cn(
                    "max-w-[85%] px-3 py-2 rounded-lg text-sm",
                    message.role === "user"
                      ? "bg-primary text-primary-foreground"
                      : "bg-muted text-foreground"
                  )}
                >
                  <p className="whitespace-pre-wrap">{message.content}</p>
                </div>
              </div>
            ))}
            {isLoading && (
              <div className="flex justify-start">
                <div className="bg-muted px-3 py-2 rounded-lg">
                  <Loader2 className="w-4 h-4 animate-spin text-muted-foreground" />
                </div>
              </div>
            )}
          </div>
        )}
      </ScrollArea>

      {/* Input */}
      <form onSubmit={handleSubmit} className="pt-3 border-t border-border mt-3">
        <div className="flex gap-2">
          <Input
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            placeholder="Ask about this case..."
            disabled={isLoading}
            className="flex-1"
          />
          <Button type="submit" size="icon" disabled={!inputValue.trim() || isLoading}>
            <Send className="w-4 h-4" />
          </Button>
        </div>
      </form>
    </div>
  );
});

export default function CopilotPanel({ containerId, containerName }: CopilotPanelProps) {
  const { getAccessToken, isAuthenticated } = useAuth();
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [chatApi, setChatApi] = useState<ChatEmbeddedAPI | null>(null);
  const [isContainerVerified, setIsContainerVerified] = useState(false);
  const [chatKey, setChatKey] = useState(0);
  const [useFallback, setUseFallback] = useState(false);
  const [sdkFailed, setSdkFailed] = useState(false);

  const authProvider = useMemo(
    () => new CopilotAuthProvider(getAccessToken),
    [getAccessToken]
  );

  // Chat configuration matching the working implementation
  const chatConfig: ChatLaunchConfig = useMemo(() => ({
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
  }), [containerName]);

  useEffect(() => {
    let cancelled = false;

    const verify = async () => {
      setIsLoading(true);
      setError(null);
      setIsContainerVerified(false);

      const verification = await verifyContainerMetadata(containerId, getAccessToken);
      if (cancelled) return;

      if (!verification.valid) {
        setError(verification.error || "Container verification failed");
        setIsLoading(false);
        return;
      }

      try {
        await authProvider.initialize();
        console.log("CopilotPanel: Auth provider initialized");
        if (cancelled) return;
        setIsContainerVerified(true);
        setIsLoading(false);
      } catch (err) {
        console.error("CopilotPanel: Auth initialization failed", err);
        if (cancelled) return;
        setError("Authentication failed. Please try again.");
        setIsLoading(false);
      }
    };

    const timeout = setTimeout(() => {
      if (isLoading && !isContainerVerified) {
        console.log("CopilotPanel: Timeout, switching to fallback");
        setUseFallback(true);
        setIsLoading(false);
      }
    }, 20000);

    verify();

    return () => {
      cancelled = true;
      clearTimeout(timeout);
    };
  }, [containerId, authProvider, getAccessToken, chatKey]);

  const handleApiReady = useCallback((api: ChatEmbeddedAPI) => {
    console.log("CopilotPanel: API ready");
    setChatApi(api);
    setIsLoading(false);
    setError(null);
  }, []);

  const handleError = useCallback((errorMessage: string) => {
    console.error("CopilotPanel: Error -", errorMessage);
    setError(errorMessage);
  }, []);

  const handleRetry = useCallback(() => {
    setError(null);
    setIsLoading(true);
    setChatApi(null);
    setIsContainerVerified(false);
    setUseFallback(false);
    setSdkFailed(false);
    setChatKey(k => k + 1);
  }, []);

  const handleSdkError = useCallback(() => {
    console.log("CopilotPanel: SDK error caught, switching to fallback");
    setSdkFailed(true);
    setUseFallback(true);
  }, []);

  const handleSwitchToFallback = useCallback(() => {
    setUseFallback(true);
  }, []);

  // Use fallback chat
  if (useFallback) {
    return (
      <div className="flex flex-col h-full">
        {sdkFailed && (
          <div className="mb-3 p-2 bg-amber-500/10 rounded-lg text-xs text-amber-700 dark:text-amber-400">
            <p className="font-medium">Using Graph API fallback</p>
            <p className="text-muted-foreground mt-1">
              The native Copilot SDK is unavailable. Using search-based chat instead.
            </p>
            <Button variant="link" size="sm" className="h-auto p-0 mt-1" onClick={handleRetry}>
              Try SDK again
            </Button>
          </div>
        )}
        <FallbackChat containerId={containerId} containerName={containerName} />
      </div>
    );
  }

  if (error && !isContainerVerified) {
    return (
      <div className="flex flex-col items-center justify-center h-full p-6 text-center">
        <div className="p-3 rounded-full bg-destructive/10 mb-3">
          <AlertTriangle className="w-6 h-6 text-destructive" />
        </div>
        <h4 className="font-semibold text-sm mb-2">Copilot Unavailable</h4>
        <p className="text-xs text-muted-foreground whitespace-pre-wrap mb-4 max-w-sm">
          {error}
        </p>
        <div className="flex flex-col gap-2 w-full max-w-xs">
          <Button variant="outline" size="sm" onClick={handleRetry}>
            <RefreshCw className="w-4 h-4 mr-2" />
            Retry SDK
          </Button>
          <Button variant="secondary" size="sm" onClick={handleSwitchToFallback}>
            Use Fallback Chat
          </Button>
        </div>
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

  if (isLoading || !isContainerVerified) {
    return (
      <div className="flex flex-col items-center justify-center h-full">
        <Loader2 className="w-6 h-6 animate-spin text-primary mb-3" />
        <p className="text-sm text-muted-foreground">
          {isContainerVerified ? "Starting Copilot..." : "Verifying container access..."}
        </p>
        <p className="text-xs text-muted-foreground mt-1">Connecting to Microsoft services</p>
        <Button 
          variant="link" 
          size="sm" 
          className="mt-4 text-xs"
          onClick={handleSwitchToFallback}
        >
          Skip and use fallback chat
        </Button>
      </div>
    );
  }

  // Use the new CopilotDesktopView component
  return (
    <CopilotErrorBoundary onRetry={handleRetry} onClose={handleSdkError}>
      <CopilotDesktopView
        isOpen={true}
        setIsOpen={() => {}}
        siteName={containerName}
        siteUrl={null}
        isLoading={isLoading}
        error={error}
        containerId={containerId}
        onError={handleError}
        onSdkFailed={handleSdkError}
        chatConfig={chatConfig}
        authProvider={authProvider}
        onApiReady={handleApiReady}
        chatKey={chatKey}
        onResetChat={handleRetry}
        isAuthenticated={isAuthenticated}
        chatApi={chatApi}
      />
    </CopilotErrorBoundary>
  );
}
