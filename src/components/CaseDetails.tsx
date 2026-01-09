import { useEffect } from "react";
import { SharePointContainer } from "@/services/sharepoint";
import { FolderNode } from "@/hooks/useFolders";
import { useFiles } from "@/hooks/useFiles";
import FileGrid from "@/components/FileGrid";
import { 
  Folder, 
  Home, 
  Upload, 
  ChevronRight,
  FileText
} from "lucide-react";
import { Button } from "@/components/ui/button";

interface CaseDetailsProps {
  container: SharePointContainer;
  selectedFolder: FolderNode | null;
}

export default function CaseDetails({ container, selectedFolder }: CaseDetailsProps) {
  const { files, isLoading, error, loadFiles } = useFiles(container?.id || null);

  useEffect(() => {
    if (container && selectedFolder) {
      loadFiles(selectedFolder.id);
    }
  }, [container?.id, selectedFolder?.id, loadFiles]);

  if (!container) {
    return null;
  }

  const lastUpdated = new Date().toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });

  const folderName = selectedFolder?.name || "Root";

  return (
    <div className="h-full flex flex-col">
      {/* Header */}
      <div className="border-b border-border p-4">
        <div className="flex items-center gap-2 text-sm text-muted-foreground mb-3">
          <Folder className="w-5 h-5 text-legal-gold" />
          <span className="font-medium text-foreground">{container.displayName}</span>
          <span>/</span>
          <span className="text-primary">{folderName}</span>
          <span className="ml-auto text-xs">Last Updated {lastUpdated}</span>
        </div>

        {/* Toolbar */}
        <div className="flex items-center gap-2 flex-wrap">
          <Button variant="ghost" size="sm" className="h-8">
            <Home className="w-4 h-4 mr-1.5" />
            Home
          </Button>
          <Button size="sm" className="h-8 bg-primary">
            <FileText className="w-4 h-4 mr-1.5" />
            New
          </Button>
          <Button variant="outline" size="sm" className="h-8">
            <Upload className="w-4 h-4 mr-1.5" />
            Upload
          </Button>
        </div>
      </div>

      {/* Breadcrumb */}
      <div className="px-4 py-2 border-b border-border flex items-center gap-1 text-sm">
        <Home className="w-4 h-4 text-muted-foreground" />
        <ChevronRight className="w-4 h-4 text-muted-foreground" />
        <button className="hover:text-primary transition-colors">
          {folderName}
        </button>
      </div>

      {/* Content Area */}
      {selectedFolder ? (
        <FileGrid 
          files={files} 
          isLoading={isLoading} 
          error={error} 
          folderName={folderName} 
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
    </div>
  );
}
