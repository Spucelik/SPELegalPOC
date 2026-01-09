import { SharePointFile, getFilePreviewUrl } from "@/services/sharepoint";
import { X, Maximize2, Minimize2, ExternalLink, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Sheet,
  SheetContent,
  SheetHeader,
  SheetTitle,
} from "@/components/ui/sheet";
import { useState, useEffect } from "react";
import { cn } from "@/lib/utils";
import { useAuth } from "@/contexts/AuthContext";

interface FileViewerProps {
  file: SharePointFile | null;
  isOpen: boolean;
  onClose: () => void;
}

export default function FileViewer({ file, isOpen, onClose }: FileViewerProps) {
  const [isExpanded, setIsExpanded] = useState(false);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const { getAccessToken } = useAuth();

  useEffect(() => {
    async function fetchPreviewUrl() {
      if (!file || !isOpen) {
        setPreviewUrl(null);
        return;
      }

      const driveId = file.parentReference?.driveId;
      if (!driveId) {
        console.error("No driveId available for file");
        setPreviewUrl(null);
        return;
      }

      setIsLoading(true);
      try {
        const token = await getAccessToken(["Files.Read.All"]);
        if (token) {
          const url = await getFilePreviewUrl(token, driveId, file.id);
          setPreviewUrl(url);
        }
      } catch (error) {
        console.error("Failed to get preview URL:", error);
        setPreviewUrl(null);
      } finally {
        setIsLoading(false);
      }
    }

    fetchPreviewUrl();
  }, [file, isOpen, getAccessToken]);

  if (!file) return null;

  return (
    <Sheet open={isOpen} onOpenChange={(open) => !open && onClose()}>
      <SheetContent 
        className={cn(
          "flex flex-col p-0 transition-all duration-300",
          isExpanded ? "sm:max-w-[80vw]" : "sm:max-w-[500px]"
        )}
        side="right"
      >
        <SheetHeader className="px-4 py-3 border-b flex-shrink-0">
          <div className="flex items-center justify-between">
            <SheetTitle className="text-base font-medium truncate pr-4">
              {file.name}
            </SheetTitle>
            <div className="flex items-center gap-1">
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                onClick={() => window.open(file.webUrl, "_blank")}
                title="Open in SharePoint"
              >
                <ExternalLink className="h-4 w-4" />
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                onClick={() => setIsExpanded(!isExpanded)}
              >
                {isExpanded ? (
                  <Minimize2 className="h-4 w-4" />
                ) : (
                  <Maximize2 className="h-4 w-4" />
                )}
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                onClick={onClose}
              >
                <X className="h-4 w-4" />
              </Button>
            </div>
          </div>
        </SheetHeader>
        
        <div className="flex-1 overflow-hidden bg-muted/30">
          {isLoading ? (
            <div className="w-full h-full flex items-center justify-center">
              <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
            </div>
          ) : previewUrl ? (
            <iframe
              src={previewUrl}
              className="w-full h-full border-0"
              title={file.name}
            />
          ) : (
            <div className="w-full h-full flex flex-col items-center justify-center p-6 text-center">
              <p className="text-muted-foreground mb-4">
                Preview not available for this file type.
              </p>
              <Button
                variant="outline"
                onClick={() => window.open(file.webUrl, "_blank")}
              >
                <ExternalLink className="h-4 w-4 mr-2" />
                Open in SharePoint
              </Button>
            </div>
          )}
        </div>
      </SheetContent>
    </Sheet>
  );
}
