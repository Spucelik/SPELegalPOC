import { useRef, useMemo, useCallback, useEffect } from "react";
// NOTE: Import from official SDK after installing via:
// npm install "https://download.microsoft.com/download/970802a5-2a7e-44ed-b17d-ad7dc99be312/microsoft-sharepointembedded-copilotchat-react-1.0.9.tgz"
import { ChatEmbedded, ChatEmbeddedAPI } from "microsoft-sharepointembedded-copilotchat-react";
import { useAuth } from "@/contexts/AuthContext";
import { createChatAuthProvider, DEFAULT_CHAT_CONFIG } from "@/services/copilotChat";
import type { ChatLaunchConfig } from "@/config/sharepoint";

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
  const chatApiRef = useRef<ChatEmbeddedAPI | null>(null);

  // Create auth provider using Container.Selected scope (required by SDK)
  const authProvider = useMemo(() => 
    createChatAuthProvider(getAccessToken),
    [getAccessToken]
  );

  // Merged config with container-specific header
  const chatConfig = useMemo(() => ({
    ...DEFAULT_CHAT_CONFIG,
    ...config,
    header: config?.header || containerName,
  }), [config, containerName]);

  // Handle API ready - auto-open the chat with configuration
  const handleApiReady = useCallback((api: ChatEmbeddedAPI) => {
    chatApiRef.current = api;
    
    // Open chat with full configuration
    api.openChat({
      header: chatConfig.header,
      zeroQueryPrompts: chatConfig.zeroQueryPrompts,
      suggestedPrompts: chatConfig.suggestedPrompts,
      instruction: chatConfig.instruction,
      locale: chatConfig.locale,
    });
  }, [chatConfig]);

  // Reset chat when container changes
  useEffect(() => {
    if (chatApiRef.current && containerId) {
      // Re-open with new container context
      chatApiRef.current.openChat({
        header: containerName,
        zeroQueryPrompts: chatConfig.zeroQueryPrompts,
        suggestedPrompts: chatConfig.suggestedPrompts,
        instruction: chatConfig.instruction,
        locale: chatConfig.locale,
      });
    }
  }, [containerId, containerName, chatConfig]);

  // Don't render if no container is selected
  if (!containerId) return null;

  return (
    <ChatEmbedded
      authProvider={authProvider}
      containerId={containerId}
      onApiReady={handleApiReady}
    />
  );
}
