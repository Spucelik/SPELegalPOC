import { SharePointFile } from "@/services/sharepoint";
import { 
  FileText, 
  FileImage, 
  FileSpreadsheet, 
  FileCode, 
  File as FileIcon,
  Loader2,
  FolderOpen
} from "lucide-react";
import { format } from "date-fns";

interface FileGridProps {
  files: SharePointFile[];
  isLoading: boolean;
  error: string | null;
  folderName: string;
}

function getFileIcon(mimeType?: string, fileName?: string) {
  if (!mimeType && fileName) {
    const ext = fileName.split('.').pop()?.toLowerCase();
    if (['jpg', 'jpeg', 'png', 'gif', 'svg', 'webp'].includes(ext || '')) {
      return <FileImage className="w-10 h-10 text-legal-gold" />;
    }
    if (['xls', 'xlsx', 'csv'].includes(ext || '')) {
      return <FileSpreadsheet className="w-10 h-10 text-green-600" />;
    }
    if (['doc', 'docx', 'pdf', 'txt'].includes(ext || '')) {
      return <FileText className="w-10 h-10 text-blue-600" />;
    }
    if (['js', 'ts', 'html', 'css', 'json'].includes(ext || '')) {
      return <FileCode className="w-10 h-10 text-purple-600" />;
    }
  }

  if (mimeType?.startsWith('image/')) {
    return <FileImage className="w-10 h-10 text-legal-gold" />;
  }
  if (mimeType?.includes('spreadsheet') || mimeType?.includes('excel')) {
    return <FileSpreadsheet className="w-10 h-10 text-green-600" />;
  }
  if (mimeType?.includes('document') || mimeType?.includes('pdf') || mimeType?.includes('word')) {
    return <FileText className="w-10 h-10 text-blue-600" />;
  }
  if (mimeType?.includes('text/') || mimeType?.includes('javascript') || mimeType?.includes('json')) {
    return <FileCode className="w-10 h-10 text-purple-600" />;
  }

  return <FileIcon className="w-10 h-10 text-muted-foreground" />;
}

function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

export default function FileGrid({ files, isLoading, error, folderName }: FileGridProps) {
  if (isLoading) {
    return (
      <div className="flex-1 flex items-center justify-center">
        <div className="text-center">
          <Loader2 className="w-12 h-12 animate-spin text-primary mx-auto mb-4" />
          <p className="text-muted-foreground">Loading files...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex-1 flex items-center justify-center">
        <div className="text-center text-destructive">
          <p className="text-lg font-medium">Failed to load files</p>
          <p className="text-sm mt-1">{error}</p>
        </div>
      </div>
    );
  }

  if (files.length === 0) {
    return (
      <div className="flex-1 flex items-center justify-center">
        <div className="text-center text-muted-foreground">
          <FolderOpen className="w-16 h-16 mx-auto mb-4 opacity-30" />
          <p className="text-lg">No files in this folder</p>
          <p className="text-sm mt-1">"{folderName}" is empty</p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex-1 overflow-auto p-6">
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-4">
        {files.map((file) => (
          <a
            key={file.id}
            href={file.webUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="group p-4 rounded-lg border border-border bg-card hover:border-primary hover:shadow-md transition-all duration-200 flex flex-col items-center text-center"
          >
            <div className="mb-3 group-hover:scale-110 transition-transform">
              {getFileIcon(file.file?.mimeType, file.name)}
            </div>
            <p className="text-sm font-medium text-foreground truncate w-full mb-1" title={file.name}>
              {file.name}
            </p>
            <p className="text-xs text-muted-foreground">
              {formatFileSize(file.size)}
            </p>
            <p className="text-xs text-muted-foreground mt-1">
              {format(new Date(file.lastModifiedDateTime), "MMM d, yyyy")}
            </p>
          </a>
        ))}
      </div>
    </div>
  );
}
