import { SharePointFile } from "@/services/sharepoint";
import { 
  Folder, 
  FileText, 
  FileSpreadsheet, 
  FileImage, 
  File,
  User,
  Loader2
} from "lucide-react";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Checkbox } from "@/components/ui/checkbox";

interface FileGridProps {
  files: SharePointFile[];
  isLoading: boolean;
  folderName: string;
}

function getFileIcon(file: SharePointFile) {
  if (file.folder) {
    return <Folder className="w-5 h-5 text-blue-500" />;
  }
  
  const mimeType = file.file?.mimeType || "";
  const name = file.name.toLowerCase();
  
  if (mimeType.includes("pdf") || name.endsWith(".pdf")) {
    return <FileText className="w-5 h-5 text-red-500" />;
  }
  if (mimeType.includes("word") || name.endsWith(".docx") || name.endsWith(".doc")) {
    return <FileText className="w-5 h-5 text-blue-600" />;
  }
  if (mimeType.includes("excel") || mimeType.includes("spreadsheet") || name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv")) {
    return <FileSpreadsheet className="w-5 h-5 text-green-600" />;
  }
  if (mimeType.includes("image") || name.match(/\.(jpg|jpeg|png|gif|bmp|webp)$/)) {
    return <FileImage className="w-5 h-5 text-purple-500" />;
  }
  
  return <File className="w-5 h-5 text-muted-foreground" />;
}

function formatDate(dateString: string): string {
  const date = new Date(dateString);
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
  });
}

function formatSize(bytes?: number, isFolder?: boolean, childCount?: number): string {
  if (isFolder && childCount !== undefined) {
    return `${childCount} item${childCount !== 1 ? "s" : ""}`;
  }
  if (!bytes) return "-";
  
  const units = ["B", "KB", "MB", "GB"];
  let size = bytes;
  let unitIndex = 0;
  
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex++;
  }
  
  return `${size.toFixed(unitIndex > 0 ? 2 : 0)} ${units[unitIndex]}`;
}

export default function FileGrid({ files, isLoading, folderName }: FileGridProps) {
  if (isLoading) {
    return (
      <div className="flex-1 flex items-center justify-center">
        <div className="flex items-center gap-3 text-muted-foreground">
          <Loader2 className="w-6 h-6 animate-spin" />
          <span>Loading contents...</span>
        </div>
      </div>
    );
  }

  if (files.length === 0) {
    return (
      <div className="flex-1 flex items-center justify-center">
        <div className="text-center text-muted-foreground">
          <Folder className="w-16 h-16 mx-auto mb-4 opacity-30" />
          <p className="text-lg">This folder is empty</p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex-1 overflow-auto">
      {/* Folder indicator */}
      <div className="px-4 py-2 flex items-center gap-2 text-sm text-muted-foreground">
        <Folder className="w-4 h-4 text-blue-500" />
        <span>{folderName}</span>
      </div>
      
      <Table>
        <TableHeader>
          <TableRow className="hover:bg-transparent">
            <TableHead className="w-[50%]">Name</TableHead>
            <TableHead>Modified</TableHead>
            <TableHead>Created</TableHead>
            <TableHead>Created By</TableHead>
            <TableHead className="text-right">Size/Items</TableHead>
          </TableRow>
        </TableHeader>
        <TableBody>
          {files.map((file) => (
            <TableRow key={file.id} className="cursor-pointer">
              <TableCell>
                <div className="flex items-center gap-3">
                  <Checkbox className="opacity-0 group-hover:opacity-100" />
                  {getFileIcon(file)}
                  <a 
                    href={file.webUrl} 
                    target="_blank" 
                    rel="noopener noreferrer"
                    className="hover:text-primary hover:underline transition-colors"
                  >
                    {file.name}
                  </a>
                </div>
              </TableCell>
              <TableCell className="text-muted-foreground">
                {formatDate(file.lastModifiedDateTime)}
              </TableCell>
              <TableCell className="text-muted-foreground">
                {formatDate(file.createdDateTime)}
              </TableCell>
              <TableCell>
                <div className="flex items-center gap-2 text-muted-foreground">
                  <User className="w-4 h-4" />
                  <span>{file.createdBy?.user?.displayName || "Unknown"}</span>
                </div>
              </TableCell>
              <TableCell className="text-right text-muted-foreground">
                {formatSize(file.size, !!file.folder, file.folder?.childCount)}
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  );
}
