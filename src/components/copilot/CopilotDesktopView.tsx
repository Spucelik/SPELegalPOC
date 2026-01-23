import React, { useEffect, useRef, useState, useCallback } from 'react';
import { Button } from '@/components/ui/button';
import { RefreshCw, AlertTriangle, ExternalLink, Loader2 } from 'lucide-react';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { ChatEmbedded, ChatEmbeddedAPI, IChatEmbeddedApiAuthProvider, ChatLaunchConfig } from '@microsoft/sharepointembedded-copilotchat-react';

interface CopilotDesktopViewProps {
  isOpen: boolean;
  setIsOpen: (value: boolean) => void;
  siteName: string;
  siteUrl: string | null;
  isLoading: boolean;
  error: string | null;
  containerId: string;
  onError: (errorMessage: string) => void;
  chatConfig: ChatLaunchConfig;
  authProvider: IChatEmbeddedApiAuthProvider;
  onApiReady: (api: ChatEmbeddedAPI) => void;
  chatKey: number;
  onResetChat?: () => void;
  isAuthenticated?: boolean;
  chatApi: ChatEmbeddedAPI | null;
}

const CopilotDesktopView: React.FC<CopilotDesktopViewProps> = ({
  isOpen,
  setIsOpen,
  siteName,
  siteUrl,
  isLoading,
  error,
  containerId,
  onError,
  chatConfig,
  authProvider,
  onApiReady,
  chatKey,
  onResetChat,
  isAuthenticated = true,
  chatApi
}) => {
  const chatInitializedRef = useRef(false);
  const containerRef = useRef<HTMLDivElement>(null);
  const [componentReady, setComponentReady] = useState(false);
  const [cspError, setCspError] = useState(false);
  const [debugInfo, setDebugInfo] = useState<{hostname: string, origin: string} | null>(null);
  
  // Enhanced CSP error detection to catch all variations
  const isCSPError = useCallback((errorMessage: string) => {
    return errorMessage.includes('Content Security Policy') || 
           errorMessage.includes('frame-ancestors') ||
           errorMessage.includes('Refused to frame') ||
           errorMessage.includes('Refused to display') ||
           errorMessage.includes('because an ancestor violates');
  }, []);
  
  // Debug initial component state
  useEffect(() => {
    console.log('🟢 CopilotDesktopView mounted with props:', {
      isOpen,
      containerId,
      siteName,
      isLoading,
      error,
      isAuthenticated,
      authHostname: authProvider.hostname
    });
  }, [isOpen, containerId, siteName, isLoading, error, isAuthenticated, authProvider.hostname]);
  
  // Capture debug information for CSP troubleshooting
  useEffect(() => {
    setDebugInfo({
      hostname: authProvider.hostname,
      origin: window.location.origin
    });
  }, [authProvider.hostname]);
  
  // Check if error is the known SDK internal error
  const isSDKInternalError = (errorMessage: string) => {
    return errorMessage.includes("Cannot read properties of undefined (reading 'name')") ||
           errorMessage.includes("chatodsp") ||
           errorMessage.includes("odsp.react.lib");
  };

  // Handle CSP errors and SDK internal errors with enhanced detection
  useEffect(() => {
    const handleError = (event: ErrorEvent) => {
      const errorMessage = event.message || event.error?.message || '';
      
      // Check for SDK internal error first
      if (isSDKInternalError(errorMessage)) {
        console.error('🚨 SDK internal error detected, switching to fallback:', errorMessage);
        onError('SDK internal error - switching to fallback');
        event.preventDefault();
        return;
      }
      
      // Then check for CSP errors
      if (isCSPError(errorMessage)) {
        console.error('🚨 CSP Error detected:', errorMessage);
        setCspError(true);
        onError(`SharePoint Content Security Policy Error: ${errorMessage}`);
        event.preventDefault();
      }
    };

    const handleUnhandledRejection = (event: PromiseRejectionEvent) => {
      const reason = event.reason?.message || event.reason || '';
      const reasonStr = typeof reason === 'string' ? reason : String(reason);
      
      // Check for SDK internal error first
      if (isSDKInternalError(reasonStr)) {
        console.error('🚨 SDK internal error in promise, switching to fallback:', reasonStr);
        onError('SDK internal error - switching to fallback');
        event.preventDefault();
        return;
      }
      
      // Then check for CSP errors
      if (isCSPError(reasonStr)) {
        console.error('🚨 CSP Promise rejection:', reasonStr);
        setCspError(true);
        onError(`SharePoint CSP Rejection: ${reasonStr}`);
        event.preventDefault();
      }
    };

    // Also listen for console errors that might indicate issues
    const originalConsoleError = console.error;
    console.error = (...args) => {
      const message = args.join(' ');
      
      if (isSDKInternalError(message)) {
        // Don't recurse, just trigger fallback
        originalConsoleError.apply(console, args);
        onError('SDK internal error - switching to fallback');
        return;
      }
      
      if (isCSPError(message)) {
        setCspError(true);
        onError(`Console CSP Error: ${message}`);
      }
      originalConsoleError.apply(console, args);
    };

    window.addEventListener('error', handleError);
    window.addEventListener('unhandledrejection', handleUnhandledRejection);
    
    return () => {
      window.removeEventListener('error', handleError);
      window.removeEventListener('unhandledrejection', handleUnhandledRejection);
      console.error = originalConsoleError;
    };
  }, [onError, isCSPError]);
  
  // Handle API ready with better initialization
  const handleApiReady = useCallback((api: ChatEmbeddedAPI) => {
    console.log('🚀 Copilot API ready, initializing chat...');
    onApiReady(api);
    setComponentReady(true);
    setCspError(false); // Reset CSP error if API is ready
  }, [onApiReady]);
  
  // Debug container contents periodically
  const debugContainerContents = useCallback(() => {
    if (containerRef.current) {
      const container = containerRef.current;
      console.log('🔍 Container debugging:', {
        hasChildren: container.children.length > 0,
        childCount: container.children.length,
        innerHTML: container.innerHTML.substring(0, 500) + '...',
        iframes: container.querySelectorAll('iframe').length,
        inputs: container.querySelectorAll('input').length,
        textareas: container.querySelectorAll('textarea').length,
        buttons: container.querySelectorAll('button').length,
        divs: container.querySelectorAll('div').length
      });
      
      // Look for any chat-related elements
      const chatElements = container.querySelectorAll('[class*="chat"], [id*="chat"], [data-*="chat"]');
      console.log('🔍 Chat-related elements found:', chatElements.length);
      
      // Check for any hidden elements
      const hiddenElements = container.querySelectorAll('[style*="display: none"], [style*="visibility: hidden"], [hidden]');
      console.log('🔍 Hidden elements found:', hiddenElements.length);
      
      // Force show any hidden input elements as a last resort
      const allInputs = container.querySelectorAll('input, textarea');
      allInputs.forEach((input, index) => {
        const styles = window.getComputedStyle(input);
        console.log(`🔍 Input ${index}:`, {
          display: styles.display,
          visibility: styles.visibility,
          opacity: styles.opacity,
          zIndex: styles.zIndex,
          position: styles.position
        });
      });
    }
  }, []);
  
  // Open chat following Microsoft documentation pattern
  const initializeCopilotChat = useCallback(async (api: ChatEmbeddedAPI) => {
    try {
      console.log('📋 Opening copilot chat with config:', {
        header: chatConfig.header,
        locale: chatConfig.locale,
        containerId: containerId,
        hasInstruction: !!chatConfig.instruction,
        authHostname: authProvider.hostname,
        currentOrigin: window.location.origin
      });
      
      // Wait for component to be fully mounted (following MS docs pattern)
      await new Promise(resolve => setTimeout(resolve, 100));
      
      // Use the exact pattern from Microsoft documentation
      await api.openChat(chatConfig);
      console.log('✅ Copilot chat opened successfully');
      
      // Debug container contents immediately after opening
      setTimeout(() => {
        debugContainerContents();
      }, 1000);
      
      // Continue debugging every few seconds to track changes
      const debugInterval = setInterval(() => {
        debugContainerContents();
      }, 3000);
      
      // Stop debugging after 15 seconds
      setTimeout(() => {
        clearInterval(debugInterval);
      }, 15000);
      
      // Monitor for CSP errors and UI issues after chat opens
      setTimeout(() => {
        if (containerRef.current) {
          const iframes = containerRef.current.querySelectorAll('iframe');
          console.log('📊 Container analysis after chat open:', {
            childElementCount: containerRef.current.childElementCount,
            iframeCount: iframes.length,
            authHostname: authProvider.hostname,
            origin: window.location.origin
          });
          
          // Check if iframe failed to load due to CSP
          if (iframes.length > 0) {
            iframes.forEach((iframe, index) => {
              iframe.addEventListener('error', () => {
                console.error(`❌ Iframe ${index} failed to load - likely CSP issue`);
                setCspError(true);
              });
              
              // Also monitor iframe content loading
              iframe.addEventListener('load', () => {
                console.log(`✅ Iframe ${index} loaded successfully`);
                // Debug iframe contents after a delay
                setTimeout(() => {
                  debugContainerContents();
                }, 2000);
              });
            });
          }
        }
      }, 2000);
      
    } catch (err) {
      console.error('❌ Error opening copilot chat:', err);
      
      // Check if it's a CSP-related error
      const errorMessage = err instanceof Error ? err.message : String(err);
      if (isCSPError(errorMessage)) {
        setCspError(true);
        onError(`SharePoint CSP Error: ${errorMessage}`);
      } else {
        // For "Failed to fetch site URL" errors, this is often transient - the reset mechanism will retry
        console.log('🔄 Chat initialization failed, reset mechanism will attempt recovery...');
        onError('Failed to initialize chat');
      }
      setComponentReady(false);
    }
  }, [chatConfig, containerId, authProvider.hostname, debugContainerContents, isCSPError, onError]);
  
  // Effect to handle chat initialization and reset
  useEffect(() => {
    console.log('🔍 Copilot initialization state:', {
      isOpen,
      hasChatApi: !!chatApi,
      componentReady,
      chatInitialized: chatInitializedRef.current,
      cspError
    });
    
    // Initialize chat when dialog opens and API is ready
    if (isOpen && chatApi && componentReady && !chatInitializedRef.current && !cspError) {
      console.log('🎯 Initializing copilot chat...');
      chatInitializedRef.current = true;
      initializeCopilotChat(chatApi);
    }
    
    if (!isOpen) {
      chatInitializedRef.current = false;
      setComponentReady(false);
      setCspError(false);
    }
  }, [isOpen, chatApi, componentReady, cspError, initializeCopilotChat]);
  
  // Reset chat when requested
  const handleResetChat = useCallback(() => {
    if (onResetChat) {
      console.log('🔄 Resetting copilot chat component');
      chatInitializedRef.current = false;
      setComponentReady(false);
      setCspError(false);
      onResetChat();
    }
  }, [onResetChat]);

  // Early return if not authenticated - AFTER all hooks
  if (!isAuthenticated) {
    console.log('CopilotDesktopView: Not rendering because not authenticated');
    return null;
  }

  if (!isOpen) return null;
  
  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-border bg-muted/50">
        <div>
          <h3 className="font-semibold text-sm">SharePoint Embedded Copilot</h3>
          <p className="text-xs text-muted-foreground">
            Connected to: {siteName || 'SharePoint Site'}
          </p>
        </div>
        {onResetChat && isAuthenticated && (
          <Button variant="ghost" size="sm" onClick={handleResetChat}>
            <RefreshCw className="w-4 h-4 mr-2" />
            Refresh
          </Button>
        )}
      </div>

      {/* Content */}
      <div className="flex-1 overflow-hidden" ref={containerRef}>
        {cspError ? (
          <div className="flex flex-col items-center justify-center h-full p-6 text-center">
            <Alert variant="destructive" className="max-w-md">
              <AlertTriangle className="h-4 w-4" />
              <AlertDescription>
                <p className="font-semibold mb-2">SharePoint Content Security Policy Restriction</p>
                <p className="text-xs mb-3">
                  Refused to frame '{debugInfo?.hostname}' because an ancestor violates CSP directive "frame-ancestors"
                </p>
                
                <p className="font-medium text-xs mb-1">What this means:</p>
                <ul className="text-xs text-left list-disc list-inside mb-3 space-y-1">
                  <li>SharePoint is configured to only allow framing from specific domains</li>
                  <li>Your app (origin: {debugInfo?.origin}) is not in the allowed list</li>
                  <li>This is a SharePoint server-side security feature that cannot be bypassed</li>
                </ul>
                
                <p className="font-medium text-xs mb-1">Required SharePoint Admin Actions:</p>
                <ul className="text-xs text-left list-disc list-inside mb-3 space-y-1">
                  <li>Update CSP headers to include: {debugInfo?.origin}</li>
                  <li>Or configure wildcard domain patterns if applicable</li>
                  <li>Or use SharePoint's native Copilot interface instead</li>
                </ul>
                
                <div className="bg-muted/50 p-2 rounded text-xs">
                  <p className="mb-1">Alternative: Access Copilot directly in SharePoint at:</p>
                  <a 
                    href={siteUrl || `https://${debugInfo?.hostname}`}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-primary hover:underline inline-flex items-center gap-1"
                  >
                    {debugInfo?.hostname}
                    <ExternalLink className="w-3 h-3" />
                  </a>
                </div>
              </AlertDescription>
            </Alert>
            {onResetChat && (
              <Button variant="outline" size="sm" onClick={handleResetChat} className="mt-4">
                <RefreshCw className="w-4 h-4 mr-2" />
                Try Again
              </Button>
            )}
          </div>
        ) : isLoading ? (
          <div className="flex flex-col items-center justify-center h-full">
            <Loader2 className="w-8 h-8 animate-spin text-primary mb-4" />
            <p className="text-sm text-muted-foreground">Loading...</p>
          </div>
        ) : error ? (
          <div className="flex flex-col items-center justify-center h-full p-6">
            <div className="p-4 rounded-full bg-destructive/10 mb-4">
              <AlertTriangle className="w-8 h-8 text-destructive" />
            </div>
            <p className="text-sm text-muted-foreground mb-4">
              {error || "Unable to load the chat. Please try again."}
            </p>
            {onResetChat && (
              <Button variant="outline" size="sm" onClick={handleResetChat}>
                <RefreshCw className="w-4 h-4 mr-2" />
                Reset Chat
              </Button>
            )}
          </div>
        ) : (
          <div className="w-full h-full" id="copilot-chat-container">
            <ChatEmbedded
              key={chatKey}
              onApiReady={handleApiReady}
              authProvider={authProvider}
              containerId={containerId}
              style={{ width: '100%', height: '100%' }}
            />
          </div>
        )}
      </div>
    </div>
  );
};

export default CopilotDesktopView;
