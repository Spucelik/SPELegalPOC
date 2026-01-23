import React, { useState, useCallback } from 'react';
import { useCopilotSite } from '@/hooks/useCopilotSite';
import { CopilotDesktopView } from '@/components/copilot';
import { toast } from '@/hooks/use-toast';
import { SHAREPOINT_CONFIG } from '@/config/sharepoint';
import { useAuth } from '@/contexts/AuthContext';
import { 
  IChatEmbeddedApiAuthProvider, 
  ChatEmbeddedAPI, 
  ChatLaunchConfig 
} from '@microsoft/sharepointembedded-copilotchat-react';

interface CopilotChatContainerProps {
  containerId: string;
  containerName?: string;
}

/**
 * Normalize SharePoint URL to ensure proper format
 */
function normalizeSharePointUrl(url: string): string {
  if (!url) return '';
  
  // Remove trailing slashes
  let normalized = url.replace(/\/+$/, '');
  
  // Ensure https:// prefix
  if (!normalized.startsWith('https://') && !normalized.startsWith('http://')) {
    normalized = `https://${normalized}`;
  }
  
  return normalized;
}

const CopilotChatContainer: React.FC<CopilotChatContainerProps> = ({ 
  containerId,
  containerName: propContainerName 
}) => {
  const [isOpen, setIsOpen] = useState(true);
  const { getAccessToken, isAuthenticated } = useAuth();
  const [chatApi, setChatApi] = useState<ChatEmbeddedAPI | null>(null);
  const [chatKey, setChatKey] = useState(0);
  
  // Validate and normalize containerId
  const normalizedContainerId = containerId && typeof containerId === 'string' 
    ? (containerId.startsWith('b!') ? containerId : containerId)
    : '';
  
  const {
    isLoading,
    error,
    siteUrl,
    siteName: hookSiteName,
    sharePointHostname,
  } = useCopilotSite(normalizedContainerId);
  
  // Use prop name or hook name
  const siteName = propContainerName || hookSiteName || 'SharePoint Site';
  
  // Ensure we have valid hostnames with proper normalization
  const rawHostname = sharePointHostname || SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME;
  const safeSharePointHostname = normalizeSharePointUrl(rawHostname);
  
  console.log('🏠 SharePoint hostname details:', {
    original: rawHostname,
    normalized: safeSharePointHostname,
    fromConfig: SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME,
    fromHook: sharePointHostname
  });
  
  const handleError = useCallback((errorMessage: string) => {
    console.error('Copilot chat error:', errorMessage);
    
    // Add delay to allow auto-recovery mechanism to work first
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
    }, 2000);
  }, []);
  
  // Create auth provider for Copilot chat with enhanced URL handling
  const authProvider = React.useMemo((): IChatEmbeddedApiAuthProvider => {
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
          
          // Use Container.Selected scope as required by SDK
          const scope = `${safeSharePointHostname}/Container.Selected`;
          const token = await getAccessToken([scope]);
          
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

    // The SDK may require siteUrl to be available on the auth provider
    (provider as any).siteUrl = containerWebUrl;
    
    console.log('🔧 Auth provider created with siteUrl:', (provider as any).siteUrl);
    
    return provider;
  }, [safeSharePointHostname, siteUrl, getAccessToken, handleError, isAuthenticated]);
  
  // Create chat configuration following Microsoft documentation
  const chatConfig = React.useMemo((): ChatLaunchConfig => {
    const config: ChatLaunchConfig = {
      header: `Case Assistant - ${siteName}`,
      instruction: "You are a legal case assistant. Provide clear, professional responses based on the case documents. Help users find information, summarize documents, and answer questions about the case files.",
      locale: "en",
      zeroQueryPrompts: {
        headerText: "How can I help you with this case?",
        promptSuggestionList: [
          { suggestionText: "Summarize the key facts of this case" },
          { suggestionText: "Who are the parties involved?" },
          { suggestionText: "What are the important dates?" },
          { suggestionText: "List the key documents" },
        ],
      },
    };
    
    console.log('📋 Created chat config:', {
      header: config.header,
      hasInstruction: !!config.instruction,
      locale: config.locale,
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
