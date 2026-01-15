import { useState, useRef, useEffect, useCallback } from "react";
import { MessageCircle, X, Send, Loader2, Bot, User } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { cn } from "@/lib/utils";
import { useAuth } from "@/contexts/AuthContext";
import { sendCopilotMessage, CopilotMessage } from "@/services/copilotChat";

interface CopilotChatProps {
  containerId: string;
  containerName: string;
}

export default function CopilotChat({ containerId, containerName }: CopilotChatProps) {
  const { getAccessToken } = useAuth();
  const [isOpen, setIsOpen] = useState(false);
  const [messages, setMessages] = useState<CopilotMessage[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

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

  const handleSendMessage = useCallback(async () => {
    if (!inputValue.trim() || isLoading) return;

    const userMessage: CopilotMessage = {
      role: "user",
      content: inputValue.trim(),
      timestamp: new Date(),
    };

    setMessages((prev) => [...prev, userMessage]);
    setInputValue("");
    setIsLoading(true);

    try {
      const token = await getAccessToken([
        "https://graph.microsoft.com/Files.Read.All",
        "https://graph.microsoft.com/Files.ReadWrite.All",
      ]);

      if (!token) {
        throw new Error("Failed to get access token");
      }

      const response = await sendCopilotMessage(
        token,
        containerId,
        containerName,
        userMessage.content,
        messages
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
  }, [inputValue, isLoading, getAccessToken, containerId, containerName, messages]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

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
            <h3 className="font-semibold text-foreground truncate">Copilot Assistant</h3>
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
            <div className="flex flex-col items-center justify-center h-full text-center py-8">
              <Bot className="w-12 h-12 text-muted-foreground/30 mb-3" />
              <p className="text-sm text-muted-foreground">
                Ask me anything about the documents in this case.
              </p>
              <p className="text-xs text-muted-foreground/70 mt-1">
                I can help you search, summarize, and analyze case files.
              </p>
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
              onClick={handleSendMessage}
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
