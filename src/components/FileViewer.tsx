import { SharePointFile } from "@/services/sharepoint";
import { X, Maximize2, Minimize2, ExternalLink, Download } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Sheet,
  SheetContent,
  SheetHeader,
  SheetTitle,
} from "@/components/ui/sheet";
import { useState } from "react";
import { cn } from "@/lib/utils";

interface FileViewerProps {
  file: SharePointFile | null;
  isOpen: boolean;
  onClose: () => void;
}

function getDirectDownloadUrl(file: SharePointFile): string | null {
  // Use the @microsoft.graph.downloadUrl for direct access (bypasses iframe restrictions)
  return file["@microsoft.graph.downloadUrl"] || null;
}

function isImageFile(file: SharePointFile): boolean {
  const mimeType = file.file?.mimeType || "";
  const name = file.name.toLowerCase();
  return mimeType.includes("image") || !!name.match(/\.(jpg|jpeg|png|gif|bmp|webp)$/);
}

function isPdfFile(file: SharePointFile): boolean {
  const mimeType = file.file?.mimeType || "";
  const name = file.name.toLowerCase();
  return mimeType.includes("pdf") || name.endsWith(".pdf");
}

export default function FileViewer({ file, isOpen, onClose }: FileViewerProps) {
  const [isExpanded, setIsExpanded] = useState(false);

  if (!file) return null;

  const downloadUrl = getDirectDownloadUrl(file);

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
              {downloadUrl && (
                <Button
                  variant="ghost"
                  size="icon"
                  className="h-8 w-8"
                  onClick={() => window.open(downloadUrl, "_blank")}
                  title="Download"
                >
                  <Download className="h-4 w-4" />
                </Button>
              )}
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
          {downloadUrl && isImageFile(file) ? (
            <div className="w-full h-full flex items-center justify-center p-4">
              <img
                src={downloadUrl}
                alt={file.name}
                className="max-w-full max-h-full object-contain rounded"
              />
            </div>
          ) : downloadUrl && isPdfFile(file) ? (
            <iframe
              src={downloadUrl}
              className="w-full h-full border-0"
              title={file.name}
            />
          ) : (
            <div className="w-full h-full flex flex-col items-center justify-center p-6 text-center">
              <p className="text-muted-foreground mb-4">
                {downloadUrl 
                  ? "Preview not available for this file type."
                  : "Direct preview not available. Use the buttons above to view or download."}
              </p>
              <div className="flex gap-2">
                <Button
                  variant="outline"
                  onClick={() => window.open(file.webUrl, "_blank")}
                >
                  <ExternalLink className="h-4 w-4 mr-2" />
                  Open in SharePoint
                </Button>
                {downloadUrl && (
                  <Button
                    variant="outline"
                    onClick={() => window.open(downloadUrl, "_blank")}
                  >
                    <Download className="h-4 w-4 mr-2" />
                    Download
                  </Button>
                )}
              </div>
            </div>
          )}
        </div>
      </SheetContent>
    </Sheet>
  );
}
