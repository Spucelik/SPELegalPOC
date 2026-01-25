import React, { useState, useCallback, useMemo } from 'react';
import { useCopilotSite } from '@/hooks/useCopilotSite';
import CopilotDesktopView from './CopilotDesktopView';
import { toast } from '@/hooks/use-toast';
import { appConfig } from '@/config/appConfig';
import { useAuth } from '@/context/AuthContext';
import { 
  IChatEmbeddedApiAuthProvider, 
  ChatEmbeddedAPI, 
  ChatLaunchConfig 
} from '@microsoft/sharepointembedded-copilotchat-react';

interface CopilotChatContainerProps {
  containerId: string;
  containerName?: string;
}

const CopilotChatContainer: React.FC<CopilotChatContainerProps> = ({ 
  containerId,
  containerName: propContainerName 
}) => {
  const [isOpen, setIsOpen] = useState(true);
  const { getSharePointToken, isAuthenticated } = useAuth();
  const [chatApi, setChatApi] = useState<ChatEmbeddedAPI | null>(null);
  const [chatKey, setChatKey] = useState(0);
  
  // Validate and normalize containerId
  const normalizedContainerId = useMemo(() => {
    if (!containerId || typeof containerId !== 'string') return '';
    return containerId.startsWith('b!') ? containerId : `b!${containerId}`;
  }, [containerId]);
  
  const {
    isLoading,
    error,
    webUrl: siteUrl,
    containerName: hookSiteName,
    sharePointHostname,
  } = useCopilotSite(normalizedContainerId);
  
  // Use prop name or hook name
  const siteName = propContainerName || hookSiteName || 'SharePoint Site';
  
  // Ensure we have valid hostnames and site names with proper normalization
  const rawHostname = sharePointHostname || appConfig.sharePointHostname;
  const safeSharePointHostname = appConfig.normalizeSharePointUrl(rawHostname);
  
  console.log('🏠 SharePoint hostname details:', {
    original: rawHostname,
    normalized: safeSharePointHostname,
    fromConfig: appConfig.sharePointHostname,
    fromHook: sharePointHostname
  });
  
  const handleError = useCallback((errorMessage: string) => {
    console.error('Copilot chat error:', errorMessage);
    
    // Add delay to allow auto-recovery mechanism to work first
    // This prevents showing error notifications when the chat recovers automatically
    setTimeout(() => {
      // Check if chat has recovered by looking for successful iframe loading
      const chatContainer = document.querySelector('[data-testid="copilot-chat-wrapper"]');
      const hasIframe = chatContainer?.querySelector('iframe');
      
      if (!hasIframe) {
        toast({
          title: "Copilot error",
          description: `${errorMessage} The system will attempt to recover automatically.`,
          variant: "destructive",
        });
      } else {
        console.log('🔄 Copilot chat recovered automatically, skipping error notification');
      }
    }, 2000); // 2 second delay to allow auto-recovery
  }, []);
  
  // Create auth provider for Copilot chat with enhanced URL handling
  const authProvider = useMemo((): IChatEmbeddedApiAuthProvider => {
    // Use the actual container webUrl as the siteUrl for the SDK
    const containerWebUrl = siteUrl || safeSharePointHostname;
    
    console.log('🔧 Creating auth provider with URLs:', {
      hostname: safeSharePointHostname,
      siteUrl: containerWebUrl,
      originalSiteUrl: siteUrl,
      fallbackHostname: safeSharePointHostname
    });

    const provider: IChatEmbeddedApiAuthProvider = {
      hostname: safeSharePointHostname,
      getToken: async () => {
        try {
          if (!isAuthenticated) {
            console.error('User not authenticated, cannot get token');
            return '';
          }
          
          console.log('🔑 Getting SharePoint token for hostname:', safeSharePointHostname);
          
          // getSharePointToken uses the configured scopes from appConfig
          const token = await getSharePointToken();
          console.log('🔑 SharePoint auth token retrieved:', token ? 'successfully' : 'failed');
          
          if (!token) {
            handleError('Failed to get authentication token for SharePoint.');
            return '';
          }
          
          return token;
        } catch (err) {
          console.error('❌ Error getting token for Copilot chat:', err);
          handleError('Failed to authenticate with SharePoint. Please try again.');
          return '';
        }
      }
    };

    // The SDK requires siteUrl to be available on the auth provider
    // Use the container's actual webUrl to ensure proper site context
    (provider as any).siteUrl = containerWebUrl;
    
    console.log('🔧 Auth provider created with siteUrl:', (provider as any).siteUrl);
    
    return provider;
  }, [safeSharePointHostname, siteUrl, getSharePointToken, handleError, isAuthenticated]);
  
  // Create chat configuration following Microsoft documentation
  const chatConfig = useMemo((): ChatLaunchConfig => {
    const config: ChatLaunchConfig = {
      header: `Copilot Chat - ${siteName}`,
      instruction: "You are a helpful AI assistant. Help users find information, answer questions, and work with their SharePoint files and documents.",
      locale: "en",
      suggestedPrompts: ["What are my files?", "Help me find documents", "Show me recent changes"]
    };
    
    console.log('📋 Created chat config:', {
      header: config.header,
      hasInstruction: !!config.instruction,
      locale: config.locale,
      suggestedPrompts: config.suggestedPrompts
    });
    
    return config;
  }, [siteName]);
  
  // Reset chat when there's an issue
  const handleResetChat = useCallback(() => {
    console.log('🔄 Resetting Copilot chat container');
    setChatKey(prev => prev + 1);
    setChatApi(null);
    setIsOpen(false);
    setTimeout(() => {
      setIsOpen(true);
    }, 500);
  }, []);
  
  // Handles API ready event from ChatEmbedded component
  const handleApiReady = useCallback((api: ChatEmbeddedAPI) => {
    if (!api) {
      console.error('❌ Chat API is undefined');
      handleError('Chat API initialization failed');
      return;
    }
    
    console.log('✅ Copilot chat API is ready');
    setChatApi(api);
  }, [handleError]);

  // Early return after all hooks are called
  if (!normalizedContainerId) {
    console.error('CopilotChatContainer: Invalid containerId provided:', containerId);
    return null;
  }

  return (
    <CopilotDesktopView
      isOpen={isOpen}
      setIsOpen={setIsOpen}
      siteName={siteName}
      siteUrl={siteUrl}
      isLoading={isLoading}
      error={error}
      containerId={normalizedContainerId}
      onError={handleError}
      chatConfig={chatConfig}
      authProvider={authProvider}
      onApiReady={handleApiReady}
      chatKey={chatKey}
      onResetChat={handleResetChat}
      isAuthenticated={isAuthenticated}
      chatApi={chatApi}
    />
  );
};

export default CopilotChatContainer;
