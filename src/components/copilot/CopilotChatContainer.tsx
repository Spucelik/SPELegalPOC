import React, { useState, useCallback, useEffect } from 'react';
import { useCopilotSite } from '@/hooks/useCopilotSite';
import { CopilotDesktopView } from '@/components/copilot';
import { toast } from '@/hooks/use-toast';
import { APP_CONFIG } from '@/config/appConfig';
import { useAuth } from '@/context/AuthContext';
import { 
  IChatEmbeddedApiAuthProvider, 
  ChatEmbeddedAPI, 
  ChatLaunchConfig 
} from '@microsoft/sharepointembedded-copilotchat-react';
import InlineCopilotChat from '@/components/copilot/InlineCopilotChat';

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
  const [useFallback, setUseFallback] = useState(false);
  const [sdkCheckComplete, setSdkCheckComplete] = useState(false);
  
  // Validate and normalize containerId
  const normalizedContainerId = containerId && typeof containerId === 'string' 
    ? (containerId.startsWith('b!') ? containerId : containerId)
    : '';
  
  const {
    isLoading,
    error,
    webUrl: siteUrl,
    containerName: hookSiteName,
    sharePointHostname,
  } = useCopilotSite(normalizedContainerId);
  
  // Use prop name or hook name
  const siteName = propContainerName || hookSiteName || 'SharePoint Site';
  
  // Ensure we have valid hostnames with proper normalization
  const rawHostname = sharePointHostname || APP_CONFIG.sharePointHostname;
  const safeSharePointHostname = normalizeSharePointUrl(rawHostname);
  
  console.log('🏠 SharePoint hostname details:', {
    original: rawHostname,
    normalized: safeSharePointHostname,
    fromConfig: APP_CONFIG.sharePointHostname,
    fromHook: sharePointHostname
  });

  // Check if SDK iframe has content after a delay - if not, use fallback
  useEffect(() => {
    if (!isOpen || useFallback || !isAuthenticated) return;

    const checkIframeContent = () => {
      const chatWrapper = document.querySelector('[data-testid="copilot-chat-wrapper"]');
      const iframe = chatWrapper?.querySelector('iframe') as HTMLIFrameElement | null;
      
      if (iframe) {
        // Check if iframe has visible content
        const rect = iframe.getBoundingClientRect();
        const hasVisibleSize = rect.width > 0 && rect.height > 0;
        
        // Try to check iframe body content (will fail for cross-origin but that's expected)
        let hasContent = false;
        try {
          const iframeBody = iframe.contentDocument?.body;
          hasContent = !!(iframeBody && iframeBody.innerHTML.trim().length > 100);
        } catch {
          // Cross-origin - assume it might have content if visible
          hasContent = hasVisibleSize;
        }
        
        console.log('🔍 SDK iframe check:', { hasVisibleSize, hasContent, width: rect.width, height: rect.height });
        
        // If iframe exists but has no visible content after 8 seconds, use fallback
        if (!hasContent || rect.height < 50) {
          console.log('⚠️ SDK iframe appears empty, switching to fallback chat');
          setUseFallback(true);
        }
      } else {
        // No iframe found after timeout - use fallback
        console.log('⚠️ No SDK iframe found, switching to fallback chat');
        setUseFallback(true);
      }
      setSdkCheckComplete(true);
    };

    // Wait 8 seconds for SDK to load, then check
    const timer = setTimeout(checkIframeContent, 8000);
    return () => clearTimeout(timer);
  }, [isOpen, useFallback, isAuthenticated, chatKey]);
  
  const handleError = useCallback((errorMessage: string) => {
    console.error('Copilot chat error:', errorMessage);
    
    // Check for CSP or SDK errors that indicate we should use fallback
    const shouldUseFallback = 
      errorMessage.includes('Content Security Policy') ||
      errorMessage.includes('frame-ancestors') ||
      errorMessage.includes('CSP') ||
      errorMessage.includes('Cannot read properties of undefined');
    
    if (shouldUseFallback) {
      console.log('🔄 Switching to fallback chat due to SDK error');
      setUseFallback(true);
      return;
    }
    
    // Add delay to allow auto-recovery mechanism to work first
    setTimeout(() => {
      const chatContainer = document.querySelector('[data-testid="copilot-chat-wrapper"]');
      const hasIframe = chatContainer?.querySelector('iframe');
      
      if (!hasIframe) {
        toast({
          title: "Copilot error",
          description: `${errorMessage} Switching to alternative chat.`,
          variant: "destructive",
        });
        setUseFallback(true);
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
    setUseFallback(false);
    setSdkCheckComplete(false);
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

  // Use fallback chat if SDK fails
  if (useFallback) {
    console.log('📱 Rendering fallback CustomCopilotChat');
    return (
      <div className="h-full w-full flex flex-col">
        <div className="flex items-center justify-between p-3 border-b border-border bg-muted/30">
          <div className="flex flex-col">
            <span className="text-sm font-semibold text-foreground">Case Assistant</span>
            <span className="text-xs text-muted-foreground">{siteName}</span>
          </div>
          <button
            onClick={handleResetChat}
            className="text-xs text-primary hover:underline"
          >
            Try SDK Again
          </button>
        </div>
        <div className="flex-1 overflow-hidden">
          <InlineCopilotChat
            containerId={normalizedContainerId}
            containerName={siteName}
            config={chatConfig}
          />
        </div>
      </div>
    );
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
