import { useState } from "react";
import { LegalCase, CaseFolder, CaseDocument } from "@/types/legal";
import { mockFolders, mockDocuments } from "@/data/mockData";
import { 
  Folder, 
  FileText, 
  Home, 
  Upload, 
  Download, 
  Trash2, 
  MoreHorizontal,
  ChevronRight,
  File,
  FileType
} from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Checkbox } from "@/components/ui/checkbox";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { cn } from "@/lib/utils";

interface CaseDetailsProps {
  legalCase: LegalCase;
}

export default function CaseDetails({ legalCase }: CaseDetailsProps) {
  const [currentPath, setCurrentPath] = useState<string[]>(["Root"]);
  const [selectedItems, setSelectedItems] = useState<Set<string>>(new Set());

  const folders = mockFolders[legalCase.id] || [];
  const documents = mockDocuments[legalCase.id] || [];

  const formatDate = (date: Date) => {
    return new Intl.DateTimeFormat("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    }).format(date);
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(2)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const getFileIcon = (fileType: string) => {
    switch (fileType) {
      case "pdf":
        return <FileType className="w-5 h-5 text-red-500" />;
      case "docx":
      case "doc":
        return <FileText className="w-5 h-5 text-blue-600" />;
      default:
        return <File className="w-5 h-5 text-muted-foreground" />;
    }
  };

  const toggleItemSelection = (id: string) => {
    const newSelected = new Set(selectedItems);
    if (newSelected.has(id)) {
      newSelected.delete(id);
    } else {
      newSelected.add(id);
    }
    setSelectedItems(newSelected);
  };

  const lastUpdated = new Date().toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });

  return (
    <div className="h-full flex flex-col">
      {/* Header */}
      <div className="border-b border-border p-4">
        <div className="flex items-center gap-2 text-sm text-muted-foreground mb-3">
          <Folder className="w-5 h-5 text-legal-gold" />
          <span className="font-medium text-foreground">{legalCase.name}</span>
          <span>/</span>
          <span className="text-primary">Root</span>
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
          <div className="h-6 w-px bg-border mx-1" />
          <Button variant="ghost" size="sm" className="h-8" disabled={selectedItems.size === 0}>
            <Download className="w-4 h-4 mr-1.5" />
            Download
          </Button>
          <Button variant="ghost" size="sm" className="h-8 text-destructive hover:text-destructive" disabled={selectedItems.size === 0}>
            <Trash2 className="w-4 h-4 mr-1.5" />
            Delete
          </Button>
        </div>
      </div>

      {/* Breadcrumb */}
      <div className="px-4 py-2 border-b border-border flex items-center gap-1 text-sm">
        <Home className="w-4 h-4 text-muted-foreground" />
        {currentPath.map((segment, index) => (
          <div key={index} className="flex items-center gap-1">
            <ChevronRight className="w-4 h-4 text-muted-foreground" />
            <button className="hover:text-primary transition-colors">
              {segment}
            </button>
          </div>
        ))}
      </div>

      {/* Content Table */}
      <div className="flex-1 overflow-auto">
        <Table>
          <TableHeader>
            <TableRow className="hover:bg-transparent">
              <TableHead className="w-12"></TableHead>
              <TableHead>Name</TableHead>
              <TableHead className="w-32">Modified</TableHead>
              <TableHead className="w-32">Created</TableHead>
              <TableHead className="w-32">Created By</TableHead>
              <TableHead className="w-24 text-right">Size/Items</TableHead>
              <TableHead className="w-12"></TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {/* Folders */}
            {folders.map((folder) => (
              <TableRow 
                key={folder.id}
                className={cn(
                  "cursor-pointer",
                  selectedItems.has(folder.id) && "bg-primary/5"
                )}
              >
                <TableCell onClick={(e) => e.stopPropagation()}>
                  <Checkbox 
                    checked={selectedItems.has(folder.id)}
                    onCheckedChange={() => toggleItemSelection(folder.id)}
                  />
                </TableCell>
                <TableCell className="font-medium">
                  <div className="flex items-center gap-2">
                    <Folder className="w-5 h-5 text-legal-gold" />
                    {folder.name}
                  </div>
                </TableCell>
                <TableCell className="text-muted-foreground text-sm">
                  {formatDate(folder.modifiedDate)}
                </TableCell>
                <TableCell className="text-muted-foreground text-sm">
                  {formatDate(folder.createdDate)}
                </TableCell>
                <TableCell className="text-muted-foreground text-sm">
                  —
                </TableCell>
                <TableCell className="text-right text-muted-foreground text-sm">
                  {folder.itemCount} items
                </TableCell>
                <TableCell>
                  <DropdownMenu>
                    <DropdownMenuTrigger asChild>
                      <Button variant="ghost" size="sm" className="h-8 w-8 p-0">
                        <MoreHorizontal className="w-4 h-4" />
                      </Button>
                    </DropdownMenuTrigger>
                    <DropdownMenuContent align="end">
                      <DropdownMenuItem>Open</DropdownMenuItem>
                      <DropdownMenuItem>Rename</DropdownMenuItem>
                      <DropdownMenuItem className="text-destructive">Delete</DropdownMenuItem>
                    </DropdownMenuContent>
                  </DropdownMenu>
                </TableCell>
              </TableRow>
            ))}

            {/* Documents */}
            {documents.map((doc) => (
              <TableRow 
                key={doc.id}
                className={cn(
                  "cursor-pointer",
                  selectedItems.has(doc.id) && "bg-primary/5"
                )}
              >
                <TableCell onClick={(e) => e.stopPropagation()}>
                  <Checkbox 
                    checked={selectedItems.has(doc.id)}
                    onCheckedChange={() => toggleItemSelection(doc.id)}
                  />
                </TableCell>
                <TableCell className="font-medium">
                  <div className="flex items-center gap-2">
                    {getFileIcon(doc.fileType)}
                    {doc.name}
                  </div>
                </TableCell>
                <TableCell className="text-muted-foreground text-sm">
                  {formatDate(doc.modifiedDate)}
                </TableCell>
                <TableCell className="text-muted-foreground text-sm">
                  {formatDate(doc.createdDate)}
                </TableCell>
                <TableCell className="text-muted-foreground text-sm">
                  {doc.createdBy}
                </TableCell>
                <TableCell className="text-right text-muted-foreground text-sm">
                  {formatSize(doc.size)}
                </TableCell>
                <TableCell>
                  <DropdownMenu>
                    <DropdownMenuTrigger asChild>
                      <Button variant="ghost" size="sm" className="h-8 w-8 p-0">
                        <MoreHorizontal className="w-4 h-4" />
                      </Button>
                    </DropdownMenuTrigger>
                    <DropdownMenuContent align="end">
                      <DropdownMenuItem>View</DropdownMenuItem>
                      <DropdownMenuItem>Download</DropdownMenuItem>
                      <DropdownMenuItem>Share</DropdownMenuItem>
                      <DropdownMenuItem className="text-destructive">Delete</DropdownMenuItem>
                    </DropdownMenuContent>
                  </DropdownMenu>
                </TableCell>
              </TableRow>
            ))}

            {folders.length === 0 && documents.length === 0 && (
              <TableRow>
                <TableCell colSpan={7} className="h-32 text-center">
                  <div className="text-muted-foreground">
                    <Folder className="w-12 h-12 mx-auto mb-2 opacity-30" />
                    <p>This case is empty</p>
                    <p className="text-sm">Upload files or create folders to get started</p>
                  </div>
                </TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </div>
    </div>
  );
}
