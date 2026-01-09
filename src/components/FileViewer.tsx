import { SharePointFile } from "@/services/sharepoint";
import { X, Maximize2, Minimize2 } from "lucide-react";
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

function getPreviewUrl(file: SharePointFile): string | null {
  // Use the webUrl for embedding - for SharePoint files this should work
  // For PDFs and images, we can embed directly
  const mimeType = file.file?.mimeType || "";
  const name = file.name.toLowerCase();
  
  if (mimeType.includes("pdf") || name.endsWith(".pdf")) {
    // For PDFs, use the webUrl with embed parameter
    return file.webUrl;
  }
  
  if (mimeType.includes("image") || name.match(/\.(jpg|jpeg|png|gif|bmp|webp)$/)) {
    return file.webUrl;
  }
  
  // For text files
  if (mimeType.includes("text") || name.match(/\.(txt|csv|json|xml|html|css|js|ts)$/)) {
    return file.webUrl;
  }
  
  // Default to webUrl
  return file.webUrl;
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

  const previewUrl = getPreviewUrl(file);

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
          {isImageFile(file) ? (
            <div className="w-full h-full flex items-center justify-center p-4">
              <img
                src={previewUrl || ""}
                alt={file.name}
                className="max-w-full max-h-full object-contain rounded"
              />
            </div>
          ) : isPdfFile(file) ? (
            <iframe
              src={`${previewUrl}?action=embedview`}
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
                Open in Browser
              </Button>
            </div>
          )}
        </div>
      </SheetContent>
    </Sheet>
  );
}
