import { useState, useEffect } from 'react';
import { useAuth } from '@/contexts/AuthContext';
import { SHAREPOINT_CONFIG, GRAPH_ENDPOINT } from '@/config/sharepoint';

interface CopilotSiteState {
  isLoading: boolean;
  error: string | null;
  siteUrl: string | null;
  siteName: string | null;
  sharePointHostname: string;
  driveId: string | null;
}

/**
 * Hook to fetch SharePoint container/site information for Copilot
 * Validates container access and retrieves metadata needed for the SDK
 */
export function useCopilotSite(containerId: string): CopilotSiteState {
  const { getAccessToken, isAuthenticated } = useAuth();
  const [state, setState] = useState<CopilotSiteState>({
    isLoading: true,
    error: null,
    siteUrl: null,
    siteName: null,
    sharePointHostname: SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME,
    driveId: null,
  });

  useEffect(() => {
    let cancelled = false;

    const fetchContainerInfo = async () => {
      if (!containerId || !isAuthenticated) {
        setState(prev => ({
          ...prev,
          isLoading: false,
          error: !containerId ? 'No container ID provided' : 'Not authenticated',
        }));
        return;
      }

      setState(prev => ({ ...prev, isLoading: true, error: null }));

      try {
        // Get token with Graph scopes for container access
        const token = await getAccessToken([
          'https://graph.microsoft.com/Files.Read.All',
          'https://graph.microsoft.com/Sites.Read.All',
        ]);

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
          `${GRAPH_ENDPOINT}/storage/fileStorage/containers/${containerId}`,
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
        console.log('📦 Container metadata fetched:', {
          id: containerData.id,
          displayName: containerData.displayName,
          status: containerData.status,
        });

        // Try to get the drive/site URL for the container
        let siteUrl: string | null = null;
        try {
          const driveResponse = await fetch(
            `${GRAPH_ENDPOINT}/drives/${containerId}`,
            {
              headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json',
              },
            }
          );

          if (driveResponse.ok) {
            const driveData = await driveResponse.json();
            siteUrl = driveData.webUrl || null;
            console.log('🔗 Drive webUrl:', siteUrl);
          }
        } catch (driveError) {
          console.warn('Could not fetch drive URL, using hostname:', driveError);
        }

        if (!cancelled) {
          setState({
            isLoading: false,
            error: null,
            siteUrl: siteUrl,
            siteName: containerData.displayName || 'SharePoint Container',
            sharePointHostname: SHAREPOINT_CONFIG.SHAREPOINT_HOSTNAME,
            driveId: containerData.id,
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
  }, [containerId, isAuthenticated, getAccessToken]);

  return state;
}
