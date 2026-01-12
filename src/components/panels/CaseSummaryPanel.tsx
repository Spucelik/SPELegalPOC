import { useEffect, useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { FileText, Users, Calendar, Clock, Loader2 } from "lucide-react";
import { useAuth } from "@/contexts/AuthContext";
import { fetchCaseSummary } from "@/services/sharepoint";

interface CaseSummaryPanelProps {
  containerName?: string;
}

export default function CaseSummaryPanel({ containerName }: CaseSummaryPanelProps) {
  const { getAccessToken } = useAuth();
  const [summary, setSummary] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const loadSummary = async () => {
      if (!containerName) {
        setSummary(null);
        return;
      }

      setIsLoading(true);
      setError(null);

      try {
        const token = await getAccessToken([
          "https://graph.microsoft.com/.default"
        ]);
        
        if (token) {
          const result = await fetchCaseSummary(token, containerName);
          setSummary(result);
        }
      } catch (err) {
        console.error("Error loading case summary:", err);
        setError("Failed to load summary");
      } finally {
        setIsLoading(false);
      }
    };

    loadSummary();
  }, [containerName, getAccessToken]);

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-base flex items-center gap-2">
            <FileText className="h-4 w-4" />
            Case Overview
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          <div>
            <p className="text-sm text-muted-foreground">Case Name</p>
            <p className="font-medium">{containerName || "No case selected"}</p>
          </div>
          <div>
            <p className="text-sm text-muted-foreground">Summary</p>
            {isLoading ? (
              <div className="flex items-center gap-2 text-muted-foreground mt-1">
                <Loader2 className="h-4 w-4 animate-spin" />
                <span className="text-sm">Generating summary...</span>
              </div>
            ) : error ? (
              <p className="text-sm text-destructive mt-1">{error}</p>
            ) : summary ? (
              <p className="text-sm mt-1 leading-relaxed">{summary}</p>
            ) : (
              <p className="text-sm text-muted-foreground mt-1 italic">
                {containerName ? "No summary available" : "Select a case to view summary"}
              </p>
            )}
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-base flex items-center gap-2">
            <Users className="h-4 w-4" />
            Team Members
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="space-y-2 text-sm">
            <div className="flex justify-between">
              <span>Lead Attorney</span>
              <span className="text-muted-foreground">John Smith</span>
            </div>
            <div className="flex justify-between">
              <span>Paralegal</span>
              <span className="text-muted-foreground">Jane Doe</span>
            </div>
            <div className="flex justify-between">
              <span>Associate</span>
              <span className="text-muted-foreground">Mike Johnson</span>
            </div>
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-base flex items-center gap-2">
            <Calendar className="h-4 w-4" />
            Key Dates
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="space-y-2 text-sm">
            <div className="flex justify-between">
              <span>Filed</span>
              <span className="text-muted-foreground">Jan 15, 2024</span>
            </div>
            <div className="flex justify-between">
              <span>Next Hearing</span>
              <span className="text-muted-foreground">Mar 20, 2024</span>
            </div>
            <div className="flex justify-between">
              <span>Discovery Deadline</span>
              <span className="text-muted-foreground">Apr 30, 2024</span>
            </div>
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-base flex items-center gap-2">
            <Clock className="h-4 w-4" />
            Recent Activity
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="space-y-3 text-sm">
            <div className="border-l-2 border-primary pl-3">
              <p className="font-medium">Document uploaded</p>
              <p className="text-muted-foreground text-xs">2 hours ago</p>
            </div>
            <div className="border-l-2 border-muted pl-3">
              <p className="font-medium">Note added</p>
              <p className="text-muted-foreground text-xs">Yesterday</p>
            </div>
            <div className="border-l-2 border-muted pl-3">
              <p className="font-medium">Status updated</p>
              <p className="text-muted-foreground text-xs">3 days ago</p>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
