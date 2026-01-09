import { useState, useCallback } from "react";
import { useAuth } from "@/contexts/AuthContext";
import { fetchFolderFiles, SharePointFile } from "@/services/sharepoint";

const GRAPH_SCOPES = ["FileStorageContainer.Selected"];

export function useFiles(containerId: string | null) {
  const { getAccessToken } = useAuth();
  const [files, setFiles] = useState<SharePointFile[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadFiles = useCallback(async (folderId: string | null) => {
    if (!containerId) {
      setFiles([]);
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      const token = await getAccessToken(GRAPH_SCOPES);
      const folderFiles = await fetchFolderFiles(token, containerId, folderId);
      setFiles(folderFiles);
    } catch (err) {
      console.error("Error loading files:", err);
      setError(err instanceof Error ? err.message : "Failed to load files");
      setFiles([]);
    } finally {
      setIsLoading(false);
    }
  }, [containerId, getAccessToken]);

  const clearFiles = useCallback(() => {
    setFiles([]);
    setError(null);
  }, []);

  return {
    files,
    isLoading,
    error,
    loadFiles,
    clearFiles,
  };
}
