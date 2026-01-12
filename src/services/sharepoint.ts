import { SHAREPOINT_CONFIG, GRAPH_ENDPOINT } from "@/config/sharepoint";

export interface SharePointContainer {
  id: string;
  displayName: string;
  description?: string;
  createdDateTime: string;
  containerTypeId: string;
}

export interface SharePointFolder {
  id: string;
  name: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  folder?: {
    childCount: number;
  };
  parentReference?: {
    id: string;
    path: string;
  };
}

interface ContainersResponse {
  value: SharePointContainer[];
}

interface DriveItemsResponse {
  value: SharePointFolder[];
}

// Fetch all containers for the configured container type
export async function fetchContainers(accessToken: string): Promise<SharePointContainer[]> {
  const url = `${GRAPH_ENDPOINT}/storage/fileStorage/containers?$filter=containerTypeId eq ${SHAREPOINT_CONFIG.CONTAINER_TYPE_ID}`;
  
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to fetch containers:", error);
    throw new Error(`Failed to fetch containers: ${response.status}`);
  }

  const data: ContainersResponse = await response.json();
  return data.value || [];
}

// Fetch root folders for a container (drive)
export async function fetchRootFolders(
  accessToken: string,
  containerId: string
): Promise<SharePointFolder[]> {
  // First get the drive ID for this container
  const driveUrl = `${GRAPH_ENDPOINT}/storage/fileStorage/containers/${containerId}/drive`;
  
  const driveResponse = await fetch(driveUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!driveResponse.ok) {
    const error = await driveResponse.text();
    console.error("Failed to fetch drive:", error);
    throw new Error(`Failed to fetch drive: ${driveResponse.status}`);
  }

  const driveData = await driveResponse.json();
  const driveId = driveData.id;

  // Now fetch root children, filtering to only folders
  const rootUrl = `${GRAPH_ENDPOINT}/drives/${driveId}/root/children?$filter=folder ne null`;
  
  const response = await fetch(rootUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to fetch root folders:", error);
    throw new Error(`Failed to fetch root folders: ${response.status}`);
  }

  const data: DriveItemsResponse = await response.json();
  return data.value || [];
}

// Fetch child folders for a specific folder
export async function fetchChildFolders(
  accessToken: string,
  containerId: string,
  folderId: string
): Promise<SharePointFolder[]> {
  // First get the drive ID for this container
  const driveUrl = `${GRAPH_ENDPOINT}/storage/fileStorage/containers/${containerId}/drive`;
  
  const driveResponse = await fetch(driveUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!driveResponse.ok) {
    throw new Error(`Failed to fetch drive: ${driveResponse.status}`);
  }

  const driveData = await driveResponse.json();
  const driveId = driveData.id;

  // Fetch children of the specific folder, filtering to only folders
  const folderUrl = `${GRAPH_ENDPOINT}/drives/${driveId}/items/${folderId}/children?$filter=folder ne null`;
  
  const response = await fetch(folderUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to fetch child folders:", error);
    throw new Error(`Failed to fetch child folders: ${response.status}`);
  }

  const data: DriveItemsResponse = await response.json();
  return data.value || [];
}

// File item interface for folder contents
export interface SharePointFile {
  id: string;
  name: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  size?: number;
  webUrl: string;
  "@microsoft.graph.downloadUrl"?: string;
  createdBy?: {
    user?: {
      displayName?: string;
      email?: string;
    };
  };
  lastModifiedBy?: {
    user?: {
      displayName?: string;
      email?: string;
    };
  };
  file?: {
    mimeType: string;
  };
  folder?: {
    childCount: number;
  };
  parentReference?: {
    driveId?: string;
  };
}

interface FolderContentsResponse {
  value: SharePointFile[];
}

// Fetch all items (files and folders) in a folder
export async function fetchFolderContents(
  accessToken: string,
  containerId: string,
  folderId: string | null
): Promise<SharePointFile[]> {
  // First get the drive ID for this container
  const driveUrl = `${GRAPH_ENDPOINT}/storage/fileStorage/containers/${containerId}/drive`;
  
  const driveResponse = await fetch(driveUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!driveResponse.ok) {
    throw new Error(`Failed to fetch drive: ${driveResponse.status}`);
  }

  const driveData = await driveResponse.json();
  const driveId = driveData.id;

  // If no folderId, get root children; otherwise get specific folder children
  const contentsUrl = folderId
    ? `${GRAPH_ENDPOINT}/drives/${driveId}/items/${folderId}/children`
    : `${GRAPH_ENDPOINT}/drives/${driveId}/root/children`;
  
  const response = await fetch(contentsUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to fetch folder contents:", error);
    throw new Error(`Failed to fetch folder contents: ${response.status}`);
  }

  const data: FolderContentsResponse = await response.json();
  return data.value || [];
}

// Get preview URL for a file (embeddable in iframe)
export async function getFilePreviewUrl(
  accessToken: string,
  driveId: string,
  itemId: string
): Promise<string | null> {
  const previewUrl = `${GRAPH_ENDPOINT}/drives/${driveId}/items/${itemId}/preview`;
  
  const response = await fetch(previewUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({}),
  });

  if (!response.ok) {
    console.error("Failed to get preview URL:", await response.text());
    return null;
  }

  const data = await response.json();
  // Add nb=true to remove the banner
  const getUrl = data.getUrl;
  if (getUrl) {
    return getUrl.includes("?") ? `${getUrl}&nb=true` : `${getUrl}?nb=true`;
  }
  return null;
}

// Get drive ID for a container
export async function getDriveId(
  accessToken: string,
  containerId: string
): Promise<string> {
  const driveUrl = `${GRAPH_ENDPOINT}/storage/fileStorage/containers/${containerId}/drive`;
  
  const driveResponse = await fetch(driveUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!driveResponse.ok) {
    throw new Error(`Failed to fetch drive: ${driveResponse.status}`);
  }

  const driveData = await driveResponse.json();
  return driveData.id;
}

// Create a new folder in a container
export async function createFolder(
  accessToken: string,
  containerId: string,
  parentFolderId: string | null,
  folderName: string
): Promise<SharePointFolder> {
  const driveId = await getDriveId(accessToken, containerId);
  
  // Use root or specific folder as parent
  const createUrl = parentFolderId
    ? `${GRAPH_ENDPOINT}/drives/${driveId}/items/${parentFolderId}/children`
    : `${GRAPH_ENDPOINT}/drives/${driveId}/root/children`;
  
  const response = await fetch(createUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename"
    }),
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to create folder:", error);
    throw new Error(`Failed to create folder: ${response.status}`);
  }

  return await response.json();
}

// Create a new empty Office file
export async function createEmptyFile(
  accessToken: string,
  containerId: string,
  parentFolderId: string | null,
  fileName: string
): Promise<SharePointFile> {
  const driveId = await getDriveId(accessToken, containerId);
  
  // Use root or specific folder as parent
  const createUrl = parentFolderId
    ? `${GRAPH_ENDPOINT}/drives/${driveId}/items/${parentFolderId}:/${fileName}:/content`
    : `${GRAPH_ENDPOINT}/drives/${driveId}/root:/${fileName}:/content`;
  
  const response = await fetch(createUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/octet-stream",
    },
    body: new Blob([]),
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to create file:", error);
    throw new Error(`Failed to create file: ${response.status}`);
  }

  return await response.json();
}

// Check if a file exists in a folder
export async function checkFileExists(
  accessToken: string,
  containerId: string,
  parentFolderId: string | null,
  fileName: string
): Promise<boolean> {
  const driveId = await getDriveId(accessToken, containerId);
  
  const checkUrl = parentFolderId
    ? `${GRAPH_ENDPOINT}/drives/${driveId}/items/${parentFolderId}:/${fileName}`
    : `${GRAPH_ENDPOINT}/drives/${driveId}/root:/${fileName}`;
  
  const response = await fetch(checkUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  return response.ok;
}

// Upload a file to a container/folder
export interface UploadProgressCallback {
  (fileName: string, progress: number): void;
}

export async function uploadFile(
  accessToken: string,
  containerId: string,
  parentFolderId: string | null,
  file: File,
  conflictBehavior: "replace" | "rename" = "rename",
  onProgress?: UploadProgressCallback
): Promise<SharePointFile> {
  const driveId = await getDriveId(accessToken, containerId);
  
  // For files larger than 4MB, we should use upload session, but for simplicity
  // we'll use direct PUT for smaller files (most common case)
  const maxDirectUploadSize = 4 * 1024 * 1024; // 4MB
  
  if (file.size > maxDirectUploadSize) {
    return uploadLargeFile(accessToken, driveId, parentFolderId, file, conflictBehavior, onProgress);
  }
  
  // Direct upload for smaller files
  const uploadUrl = parentFolderId
    ? `${GRAPH_ENDPOINT}/drives/${driveId}/items/${parentFolderId}:/${file.name}:/content?@microsoft.graph.conflictBehavior=${conflictBehavior}`
    : `${GRAPH_ENDPOINT}/drives/${driveId}/root:/${file.name}:/content?@microsoft.graph.conflictBehavior=${conflictBehavior}`;
  
  onProgress?.(file.name, 50); // Simulate progress for small files
  
  const response = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": file.type || "application/octet-stream",
    },
    body: file,
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to upload file:", error);
    throw new Error(`Failed to upload file: ${response.status}`);
  }

  onProgress?.(file.name, 100);
  return await response.json();
}

// Upload large files using upload session
async function uploadLargeFile(
  accessToken: string,
  driveId: string,
  parentFolderId: string | null,
  file: File,
  conflictBehavior: "replace" | "rename",
  onProgress?: UploadProgressCallback
): Promise<SharePointFile> {
  // Create upload session
  const sessionUrl = parentFolderId
    ? `${GRAPH_ENDPOINT}/drives/${driveId}/items/${parentFolderId}:/${file.name}:/createUploadSession`
    : `${GRAPH_ENDPOINT}/drives/${driveId}/root:/${file.name}:/createUploadSession`;
  
  const sessionResponse = await fetch(sessionUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      item: {
        "@microsoft.graph.conflictBehavior": conflictBehavior,
        name: file.name,
      },
    }),
  });

  if (!sessionResponse.ok) {
    throw new Error(`Failed to create upload session: ${sessionResponse.status}`);
  }

  const session = await sessionResponse.json();
  const uploadUrl = session.uploadUrl;

  // Upload in chunks
  const chunkSize = 320 * 1024 * 10; // 3.2MB chunks (must be multiple of 320KB)
  let uploadedBytes = 0;
  let result: SharePointFile | null = null;

  while (uploadedBytes < file.size) {
    const chunk = file.slice(uploadedBytes, uploadedBytes + chunkSize);
    const chunkEnd = Math.min(uploadedBytes + chunkSize, file.size) - 1;

    const chunkResponse = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": chunk.size.toString(),
        "Content-Range": `bytes ${uploadedBytes}-${chunkEnd}/${file.size}`,
      },
      body: chunk,
    });

    if (!chunkResponse.ok && chunkResponse.status !== 202) {
      throw new Error(`Failed to upload chunk: ${chunkResponse.status}`);
    }

    uploadedBytes += chunk.size;
    const progress = Math.round((uploadedBytes / file.size) * 100);
    onProgress?.(file.name, progress);

    // If upload is complete, the response will contain the file metadata
    if (chunkResponse.status === 200 || chunkResponse.status === 201) {
      result = await chunkResponse.json();
    }
  }

  if (!result) {
    throw new Error("Upload completed but no file metadata received");
  }

  return result;
}

// Copilot retrieval response interface
export interface CopilotRetrievalResponse {
  retrievalHits?: Array<{
    webUrl?: string;
    extracts?: Array<{
      text?: string;
      relevanceScore?: number;
    }>;
    resourceType?: string;
  }>;
}

// Fetch case summary using Microsoft Copilot retrieval API
export async function fetchCaseSummary(
  accessToken: string,
  caseTitle: string
): Promise<string | null> {
  const url = "https://graph.microsoft.com/beta/copilot/microsoft.graph.retrieval";
  
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      queryString: `Summarize the case for ${caseTitle}`,
      dataSource: "sharePointEmbedded",
      dataSourceConfiguration: {
        SharePointEmbedded: {
          ContainerTypeId: SHAREPOINT_CONFIG.CONTAINER_TYPE_ID,
        },
      },
    }),
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to fetch case summary:", error);
    return null;
  }

  const data: CopilotRetrievalResponse = await response.json();
  
  // Extract and return the text from the first retrieval hit's extract
  if (data.retrievalHits && data.retrievalHits.length > 0) {
    const firstHit = data.retrievalHits[0];
    if (firstHit.extracts && firstHit.extracts.length > 0 && firstHit.extracts[0].text) {
      // Clean up the text by removing page markers and excessive whitespace
      let text = firstHit.extracts[0].text;
      text = text.replace(/<page_\d+>/g, '').replace(/<\/page_\d+>/g, '');
      text = text.replace(/\\_/g, '_').replace(/\\-/g, '-').replace(/\\[/g, '[').replace(/\\]/g, ']').replace(/\\(/g, '(').replace(/\\)/g, ')');
      text = text.replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
      return text;
    }
  }
  
  return null;
}
