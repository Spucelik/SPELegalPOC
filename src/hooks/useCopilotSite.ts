import { useState, useEffect } from 'react';
import { useAuth } from '@/context/AuthContext';
import { APP_CONFIG, ENDPOINTS, SCOPES } from '@/config/appConfig';

interface CopilotSiteState {
  isLoading: boolean;
  error: string | null;
  containerId: string | null;
  containerName: string | null;
  webUrl: string | null;
  sharePointHostname: string;
}

/**
 * Hook to fetch SharePoint container/site information for Copilot.
 * 
 * - Normalizes container ID (adds b! prefix if missing)
 * - Fetches container name and webUrl via Graph API
 * - Extracts SharePoint hostname for authentication
 */
export function useCopilotSite(rawContainerId: string | null): CopilotSiteState {
  const { getAccessToken, isAuthenticated } = useAuth();
  const [state, setState] = useState<CopilotSiteState>({
    isLoading: false,
    error: null,
    containerId: null,
    containerName: null,
    webUrl: null,
    sharePointHostname: APP_CONFIG.sharePointHostname,
  });

  useEffect(() => {
    let cancelled = false;

    const fetchContainerInfo = async () => {
      if (!rawContainerId || !isAuthenticated) {
        setState(prev => ({
          ...prev,
          isLoading: false,
          error: !rawContainerId ? null : 'Not authenticated',
          containerId: null,
          containerName: null,
          webUrl: null,
        }));
        return;
      }

      setState(prev => ({ ...prev, isLoading: true, error: null }));

      try {
        // Strip b! prefix if present - Graph API expects raw container GUID
        const normalizedId = rawContainerId.startsWith('b!') 
          ? rawContainerId.slice(2) 
          : rawContainerId;

        // Get token with Graph scopes for container access
        const token = await getAccessToken(SCOPES.graph);

        if (!token) {
          if (!cancelled) {
            setState(prev => ({
              ...prev,
              isLoading: false,
              error: 'Failed to acquire access token',
            }));
          }
          return;
        }

        // Fetch container metadata
        const response = await fetch(
          `${ENDPOINTS.graph}/storage/fileStorage/containers/${normalizedId}`,
          {
            headers: {
              Authorization: `Bearer ${token}`,
              'Content-Type': 'application/json',
            },
          }
        );

        if (!response.ok) {
          const errorText = await response.text();
          console.error('Container fetch error:', response.status, errorText);
          if (!cancelled) {
            setState(prev => ({
              ...prev,
              isLoading: false,
              error: `Container not accessible: ${response.status}`,
            }));
          }
          return;
        }

        const containerData = await response.json();
        console.log('📦 Container metadata:', {
          id: containerData.id,
          displayName: containerData.displayName,
          containerTypeId: containerData.containerTypeId,
          settings: containerData.settings,
          status: containerData.status,
          // Log full response for debugging
          fullResponse: JSON.stringify(containerData, null, 2),
        });
        
        // Check if container has the expected structure for Copilot
        if (!containerData.id) {
          console.warn('⚠️ Container response missing ID - may indicate configuration issue');
        }

        // Try to get the drive webUrl
        let webUrl: string | null = null;
        try {
          const driveResponse = await fetch(
            `${ENDPOINTS.graph}/drives/${normalizedId}`,
            {
              headers: {
                Authorization: `Bearer ${token}`,
              },
            }
          );

          if (driveResponse.ok) {
            const driveData = await driveResponse.json();
            webUrl = driveData.webUrl || null;
          }
        } catch (driveError) {
          console.warn('Could not fetch drive URL:', driveError);
        }

        if (!cancelled) {
          setState({
            isLoading: false,
            error: null,
            containerId: normalizedId,
            containerName: containerData.displayName || 'SharePoint Container',
            webUrl,
            sharePointHostname: APP_CONFIG.sharePointHostname,
          });
        }
      } catch (err) {
        console.error('Error fetching container info:', err);
        if (!cancelled) {
          setState(prev => ({
            ...prev,
            isLoading: false,
            error: err instanceof Error ? err.message : 'Failed to fetch container info',
          }));
        }
      }
    };

    fetchContainerInfo();

    return () => {
      cancelled = true;
    };
  }, [rawContainerId, isAuthenticated, getAccessToken]);

  return state;
}
