import { useEffect, useState, useCallback } from "react";
import { SharePointContainer, createFolder, createEmptyFile } from "@/services/sharepoint";
import { FolderNode } from "@/hooks/useFolders";
import { useFiles } from "@/hooks/useFiles";
import { useAuth } from "@/contexts/AuthContext";
import { 
  Folder, 
  Home, 
  Upload, 
  ChevronRight,
  ChevronDown,
  FolderPlus,
  FilePlus,
  Plus
} from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import FileGrid from "@/components/FileGrid";
import NewFolderDialog from "@/components/NewFolderDialog";
import NewDocumentDialog from "@/components/NewDocumentDialog";
import { toast } from "sonner";

interface BreadcrumbItem {
  id: string | null;
  name: string;
}

interface CaseDetailsProps {
  container: SharePointContainer;
  selectedFolder: FolderNode | null;
}

export default function CaseDetails({ container, selectedFolder }: CaseDetailsProps) {
  const { getAccessToken } = useAuth();
  const { files, isLoading, loadFolderContents } = useFiles(container?.id || null);
  const [currentFolderId, setCurrentFolderId] = useState<string | null>(null);
  const [breadcrumbs, setBreadcrumbs] = useState<BreadcrumbItem[]>([]);
  const [isNewFolderDialogOpen, setIsNewFolderDialogOpen] = useState(false);
  const [isNewDocumentDialogOpen, setIsNewDocumentDialogOpen] = useState(false);
  const [isCreating, setIsCreating] = useState(false);

  // Reset when selected folder changes from sidebar
  useEffect(() => {
    if (selectedFolder) {
      setCurrentFolderId(selectedFolder.id);
      setBreadcrumbs([{ id: selectedFolder.id, name: selectedFolder.name }]);
    }
  }, [selectedFolder]);

  // Load folder contents when currentFolderId changes
  useEffect(() => {
    if (container?.id && currentFolderId) {
      loadFolderContents(currentFolderId);
    }
  }, [container?.id, currentFolderId, loadFolderContents]);

  const handleFolderClick = useCallback((folderId: string, folderName: string) => {
    setCurrentFolderId(folderId);
    setBreadcrumbs(prev => [...prev, { id: folderId, name: folderName }]);
  }, []);

  const handleBreadcrumbClick = useCallback((index: number) => {
    if (index < breadcrumbs.length - 1) {
      const targetCrumb = breadcrumbs[index];
      setCurrentFolderId(targetCrumb.id);
      setBreadcrumbs(prev => prev.slice(0, index + 1));
    }
  }, [breadcrumbs]);

  const handleHomeClick = useCallback(() => {
    if (selectedFolder) {
      setCurrentFolderId(selectedFolder.id);
      setBreadcrumbs([{ id: selectedFolder.id, name: selectedFolder.name }]);
    }
  }, [selectedFolder]);

  const handleCreateFolder = useCallback(async (folderName: string) => {
    if (!container?.id) return;
    
    setIsCreating(true);
    try {
      const accessToken = await getAccessToken(["FileStorageContainer.Selected"]);
      if (!accessToken) {
        toast.error("Failed to get access token");
        return;
      }

      const newFolder = await createFolder(
        accessToken,
        container.id,
        currentFolderId,
        folderName
      );

      toast.success(`Folder "${folderName}" created successfully`);
      setIsNewFolderDialogOpen(false);
      
      // Navigate to the newly created folder
      setCurrentFolderId(newFolder.id);
      setBreadcrumbs(prev => [...prev, { id: newFolder.id, name: newFolder.name }]);
    } catch (error) {
      console.error("Failed to create folder:", error);
      toast.error("Failed to create folder");
    } finally {
      setIsCreating(false);
    }
  }, [container?.id, currentFolderId, getAccessToken]);

  const handleCreateFile = useCallback(async (fileName: string, extension: string) => {
    if (!container?.id) return;
    
    setIsCreating(true);
    try {
      const accessToken = await getAccessToken(["FileStorageContainer.Selected"]);
      if (!accessToken) {
        toast.error("Failed to get access token");
        return;
      }

      const fullFileName = `${fileName}.${extension}`;
      await createEmptyFile(
        accessToken,
        container.id,
        currentFolderId,
        fullFileName
      );

      toast.success(`File "${fullFileName}" created successfully`);
      setIsNewDocumentDialogOpen(false);
      
      // Refresh the folder contents to show the new file
      if (currentFolderId) {
        loadFolderContents(currentFolderId);
      }
    } catch (error) {
      console.error("Failed to create file:", error);
      toast.error("Failed to create file");
    } finally {
      setIsCreating(false);
    }
  }, [container?.id, currentFolderId, getAccessToken, loadFolderContents]);

  if (!container) {
    return null;
  }

  const lastUpdated = new Date().toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });

  const currentFolderName = breadcrumbs.length > 0 
    ? breadcrumbs[breadcrumbs.length - 1].name 
    : selectedFolder?.name || "Root";

  return (
    <div className="h-full flex flex-col">
      {/* Header */}
      <div className="border-b border-border p-4">
        <div className="flex items-center gap-2 text-sm text-muted-foreground mb-3">
          <Folder className="w-5 h-5 text-legal-gold" />
          <span className="font-medium text-foreground">{container.displayName}</span>
          <span>/</span>
          <span className="text-primary">{currentFolderName}</span>
          <span className="ml-auto text-xs">Last Updated {lastUpdated}</span>
        </div>

        {/* Toolbar */}
        <div className="flex items-center gap-2 flex-wrap">
          <Button variant="ghost" size="sm" className="h-8" onClick={handleHomeClick}>
            <Home className="w-4 h-4 mr-1.5" />
            Home
          </Button>
          
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <Button size="sm" className="h-8 bg-primary">
                <Plus className="w-4 h-4 mr-1.5" />
                New
                <ChevronDown className="w-4 h-4 ml-1" />
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="start" className="w-48">
              <DropdownMenuItem onClick={() => setIsNewFolderDialogOpen(true)}>
                <FolderPlus className="w-4 h-4 mr-2" />
                New Folder
              </DropdownMenuItem>
              <DropdownMenuItem onClick={() => setIsNewDocumentDialogOpen(true)}>
                <FilePlus className="w-4 h-4 mr-2" />
                New Document
              </DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>

          <Button variant="outline" size="sm" className="h-8">
            <Upload className="w-4 h-4 mr-1.5" />
            Upload
          </Button>
        </div>
      </div>

      {/* Breadcrumb */}
      <div className="px-4 py-2 border-b border-border flex items-center gap-1 text-sm overflow-x-auto">
        <button 
          className="hover:text-primary transition-colors flex items-center gap-1"
          onClick={handleHomeClick}
        >
          <Home className="w-4 h-4 text-muted-foreground" />
        </button>
        {breadcrumbs.map((crumb, index) => (
          <span key={crumb.id || index} className="flex items-center gap-1">
            <ChevronRight className="w-4 h-4 text-muted-foreground" />
            <button 
              className={`hover:text-primary transition-colors ${
                index === breadcrumbs.length - 1 ? "font-medium" : ""
              }`}
              onClick={() => handleBreadcrumbClick(index)}
            >
              {crumb.name}
            </button>
          </span>
        ))}
      </div>

      {/* Content Area */}
      {selectedFolder ? (
        <FileGrid 
          files={files} 
          isLoading={isLoading} 
          folderName={currentFolderName}
          onFolderClick={handleFolderClick}
        />
      ) : (
        <div className="flex-1 overflow-auto flex items-center justify-center">
          <div className="text-center text-muted-foreground">
            <Folder className="w-16 h-16 mx-auto mb-4 opacity-30" />
            <p className="text-lg">Select a folder to view contents</p>
            <p className="text-sm mt-1">Click on a folder in the sidebar</p>
          </div>
        </div>
      )}

      {/* Dialogs */}
      <NewFolderDialog
        isOpen={isNewFolderDialogOpen}
        onClose={() => setIsNewFolderDialogOpen(false)}
        onCreateFolder={handleCreateFolder}
        isCreating={isCreating}
      />

      <NewDocumentDialog
        isOpen={isNewDocumentDialogOpen}
        onClose={() => setIsNewDocumentDialogOpen(false)}
        onCreateFile={handleCreateFile}
        isCreating={isCreating}
      />
    </div>
  );
}
