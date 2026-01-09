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

export interface SharePointFile {
  id: string;
  name: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  size: number;
  webUrl: string;
  file?: {
    mimeType: string;
  };
  createdBy?: {
    user?: {
      displayName: string;
      email: string;
    };
  };
  lastModifiedBy?: {
    user?: {
      displayName: string;
      email: string;
    };
  };
}

interface ContainersResponse {
  value: SharePointContainer[];
}

interface DriveItemsResponse {
  value: SharePointFolder[];
}

interface FilesResponse {
  value: SharePointFile[];
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

// Fetch files for a specific folder (excludes folders)
export async function fetchFolderFiles(
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

  // Fetch children - if folderId is null, fetch from root, otherwise from specific folder
  // Note: $filter is not supported on children endpoint, so we filter client-side
  const filesUrl = folderId 
    ? `${GRAPH_ENDPOINT}/drives/${driveId}/items/${folderId}/children`
    : `${GRAPH_ENDPOINT}/drives/${driveId}/root/children`;
  
  const response = await fetch(filesUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.text();
    console.error("Failed to fetch folder files:", error);
    throw new Error(`Failed to fetch folder files: ${response.status}`);
  }

  const data: FilesResponse = await response.json();
  // Filter to only include files (items with file property), exclude folders
  return (data.value || []).filter(item => item.file);
}
