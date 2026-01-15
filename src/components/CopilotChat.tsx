import { useState, useRef, useEffect, useCallback, useMemo } from "react";
import { MessageCircle, X, Send, Loader2, Bot, User, Sparkles } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { cn } from "@/lib/utils";
import { useAuth } from "@/contexts/AuthContext";
import { 
  sendCopilotMessage, 
  CopilotMessage, 
  createChatAuthProvider,
  DEFAULT_CHAT_CONFIG,
  type ChatLaunchConfig 
} from "@/services/copilotChat";
import { IChatEmbeddedApiAuthProvider } from "@/config/sharepoint";

interface CopilotChatProps {
  containerId: string;
  containerName: string;
  config?: ChatLaunchConfig;
}

export default function CopilotChat({ 
  containerId, 
  containerName, 
  config = DEFAULT_CHAT_CONFIG 
}: CopilotChatProps) {
  const { getAccessToken } = useAuth();
  const [isOpen, setIsOpen] = useState(false);
  const [messages, setMessages] = useState<CopilotMessage[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Create auth provider following SDK's IChatEmbeddedApiAuthProvider interface
  const authProvider: IChatEmbeddedApiAuthProvider = useMemo(() => 
    createChatAuthProvider(getAccessToken),
    [getAccessToken]
  );

  // Merged config with defaults
  const chatConfig = useMemo(() => ({
    ...DEFAULT_CHAT_CONFIG,
    ...config,
    header: config?.header || containerName,
  }), [config, containerName]);

  // Reset chat when container changes
  useEffect(() => {
    setMessages([]);
    setInputValue("");
    setIsLoading(false);
  }, [containerId]);

  // Close chat when navigating away (container changes to null)
  useEffect(() => {
    if (!containerId) {
      setIsOpen(false);
    }
  }, [containerId]);

  // Scroll to bottom when messages update
  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [messages]);

  // Focus input when chat opens
  useEffect(() => {
    if (isOpen && inputRef.current) {
      inputRef.current.focus();
    }
  }, [isOpen]);

  const handleSendMessage = useCallback(async (messageText?: string) => {
    const text = messageText || inputValue.trim();
    if (!text || isLoading) return;

    const userMessage: CopilotMessage = {
      role: "user",
      content: text,
      timestamp: new Date(),
    };

    setMessages((prev) => [...prev, userMessage]);
    setInputValue("");
    setIsLoading(true);

    try {
      const response = await sendCopilotMessage(
        authProvider,
        containerId,
        containerName,
        userMessage.content,
        messages,
        chatConfig
      );

      const assistantMessage: CopilotMessage = {
        role: "assistant",
        content: response,
        timestamp: new Date(),
      };

      setMessages((prev) => [...prev, assistantMessage]);
    } catch (error) {
      console.error("Chat error:", error);
      const errorMessage: CopilotMessage = {
        role: "assistant",
        content: "I'm sorry, I encountered an error processing your request. Please try again.",
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  }, [inputValue, isLoading, authProvider, containerId, containerName, messages, chatConfig]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const handleSuggestedPrompt = (prompt: string) => {
    handleSendMessage(prompt);
  };

  // Don't render if no container is selected
  if (!containerId) return null;

  const zeroQueryPrompts = chatConfig.zeroQueryPrompts;
  const suggestedPrompts = chatConfig.suggestedPrompts;

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
          "fixed bottom-24 right-6 z-40 w-96 max-w-[calc(100vw-3rem)]",
          "bg-card border border-border rounded-xl shadow-2xl",
          "flex flex-col overflow-hidden transition-all duration-300",
          isOpen
            ? "opacity-100 translate-y-0 pointer-events-auto"
            : "opacity-0 translate-y-4 pointer-events-none"
        )}
        style={{ height: "500px", maxHeight: "calc(100vh - 150px)" }}
      >
        {/* Header */}
        <div className="flex items-center gap-3 p-4 border-b border-border bg-muted/50">
          <div className="flex items-center justify-center w-10 h-10 rounded-full bg-primary/10">
            <Bot className="w-5 h-5 text-primary" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="font-semibold text-foreground truncate">
              {chatConfig.header}
            </h3>
            <p className="text-xs text-muted-foreground truncate">{containerName}</p>
          </div>
          <Button
            variant="ghost"
            size="icon"
            className="h-8 w-8"
            onClick={() => setIsOpen(false)}
          >
            <X className="w-4 h-4" />
          </Button>
        </div>

        {/* Messages Area */}
        <ScrollArea className="flex-1 p-4" ref={scrollRef}>
          {messages.length === 0 ? (
            <div className="flex flex-col h-full py-4">
              {/* Zero Query State with Starter Prompts */}
              <div className="flex flex-col items-center text-center mb-6">
                <div className="w-12 h-12 rounded-full bg-primary/10 flex items-center justify-center mb-3">
                  <Sparkles className="w-6 h-6 text-primary" />
                </div>
                <p className="text-sm font-medium text-foreground">
                  {zeroQueryPrompts?.headerText || "How can I help you?"}
                </p>
              </div>

              {/* Zero Query Prompt Suggestions */}
              {zeroQueryPrompts?.promptSuggestionList && (
                <div className="space-y-2 mb-4">
                  {zeroQueryPrompts.promptSuggestionList.map((prompt, index) => (
                    <button
                      key={index}
                      onClick={() => handleSuggestedPrompt(prompt.suggestionText)}
                      className={cn(
                        "w-full text-left px-4 py-3 rounded-lg",
                        "bg-muted/50 hover:bg-muted transition-colors",
                        "text-sm text-foreground",
                        "border border-border/50 hover:border-border"
                      )}
                    >
                      {prompt.suggestionText}
                    </button>
                  ))}
                </div>
              )}

              {/* Additional Suggested Prompts */}
              {suggestedPrompts && suggestedPrompts.length > 0 && (
                <div className="mt-auto">
                  <p className="text-xs text-muted-foreground mb-2">Or try:</p>
                  <div className="flex flex-wrap gap-2">
                    {suggestedPrompts.map((prompt, index) => (
                      <button
                        key={index}
                        onClick={() => handleSuggestedPrompt(prompt)}
                        className={cn(
                          "px-3 py-1.5 rounded-full text-xs",
                          "bg-primary/10 hover:bg-primary/20 text-primary",
                          "transition-colors"
                        )}
                      >
                        {prompt}
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          ) : (
            <div className="space-y-4">
              {messages.map((message, index) => (
                <div
                  key={index}
                  className={cn(
                    "flex gap-3",
                    message.role === "user" ? "flex-row-reverse" : ""
                  )}
                >
                  <div
                    className={cn(
                      "flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center",
                      message.role === "user"
                        ? "bg-primary text-primary-foreground"
                        : "bg-muted"
                    )}
                  >
                    {message.role === "user" ? (
                      <User className="w-4 h-4" />
                    ) : (
                      <Bot className="w-4 h-4" />
                    )}
                  </div>
                  <div
                    className={cn(
                      "flex-1 px-4 py-2 rounded-xl text-sm max-w-[75%]",
                      message.role === "user"
                        ? "bg-primary text-primary-foreground ml-auto"
                        : "bg-muted text-foreground"
                    )}
                  >
                    <p className="whitespace-pre-wrap">{message.content}</p>
                  </div>
                </div>
              ))}
              {isLoading && (
                <div className="flex gap-3">
                  <div className="flex-shrink-0 w-8 h-8 rounded-full bg-muted flex items-center justify-center">
                    <Bot className="w-4 h-4" />
                  </div>
                  <div className="bg-muted px-4 py-2 rounded-xl">
                    <Loader2 className="w-4 h-4 animate-spin" />
                  </div>
                </div>
              )}
            </div>
          )}
        </ScrollArea>

        {/* Input Area */}
        <div className="p-4 border-t border-border bg-card">
          <div className="flex gap-2">
            <Input
              ref={inputRef}
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder="Ask about this case..."
              disabled={isLoading}
              className="flex-1"
            />
            <Button
              onClick={() => handleSendMessage()}
              disabled={!inputValue.trim() || isLoading}
              size="icon"
            >
              {isLoading ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Send className="w-4 h-4" />
              )}
            </Button>
          </div>
        </div>
      </div>
    </>
  );
}