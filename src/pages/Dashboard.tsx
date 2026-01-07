import { useState, useEffect } from "react";
import { useAuth } from "@/contexts/AuthContext";
import { useNavigate } from "react-router-dom";
import Header from "@/components/Header";
import CaseCard from "@/components/CaseCard";
import CaseDetails from "@/components/CaseDetails";
import { LegalCase } from "@/types/legal";
import { mockCases } from "@/data/mockData";
import { Plus, Briefcase } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";

export default function Dashboard() {
  const { isAuthenticated, isInitialized } = useAuth();
  const navigate = useNavigate();
  const [cases, setCases] = useState<LegalCase[]>(mockCases);
  const [selectedCase, setSelectedCase] = useState<LegalCase | null>(null);
  const [isCreateDialogOpen, setIsCreateDialogOpen] = useState(false);
  const [newCaseName, setNewCaseName] = useState("");

  useEffect(() => {
    if (isInitialized && !isAuthenticated) {
      navigate("/");
    }
  }, [isInitialized, isAuthenticated, navigate]);

  useEffect(() => {
    // Select first case by default
    if (cases.length > 0 && !selectedCase) {
      setSelectedCase(cases[0]);
    }
  }, [cases, selectedCase]);

  const handleCreateCase = () => {
    if (!newCaseName.trim()) return;

    const newCase: LegalCase = {
      id: `case-${Date.now()}`,
      name: newCaseName.trim(),
      createdDate: new Date(),
      modifiedDate: new Date(),
      status: "active",
      folderCount: 0,
      documentCount: 0,
      containerId: `container-${Date.now()}`,
    };

    setCases([newCase, ...cases]);
    setSelectedCase(newCase);
    setNewCaseName("");
    setIsCreateDialogOpen(false);
  };

  if (!isInitialized) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-background">
        <div className="animate-pulse text-muted-foreground">Loading...</div>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex flex-col bg-background">
      <Header />

      <div className="flex-1 flex">
        {/* Sidebar - Cases List */}
        <aside className="w-80 border-r border-border bg-card flex flex-col">
          <div className="p-4 border-b border-border">
            <div className="flex items-center justify-between mb-1">
              <h2 className="text-lg font-semibold text-foreground">Cases</h2>
              <Dialog open={isCreateDialogOpen} onOpenChange={setIsCreateDialogOpen}>
                <DialogTrigger asChild>
                  <Button size="sm" className="h-8">
                    <Plus className="w-4 h-4 mr-1" />
                    New
                  </Button>
                </DialogTrigger>
                <DialogContent>
                  <DialogHeader>
                    <DialogTitle>Create New Legal Case</DialogTitle>
                    <DialogDescription>
                      Enter a name for the new legal case. This will create a new secure container for all case-related documents.
                    </DialogDescription>
                  </DialogHeader>
                  <div className="py-4">
                    <Label htmlFor="case-name">Case Name</Label>
                    <Input
                      id="case-name"
                      value={newCaseName}
                      onChange={(e) => setNewCaseName(e.target.value)}
                      placeholder="e.g., Smith vs Johnson Corp"
                      className="mt-2"
                      onKeyDown={(e) => e.key === "Enter" && handleCreateCase()}
                    />
                  </div>
                  <DialogFooter>
                    <Button variant="outline" onClick={() => setIsCreateDialogOpen(false)}>
                      Cancel
                    </Button>
                    <Button onClick={handleCreateCase} disabled={!newCaseName.trim()}>
                      Create Case
                    </Button>
                  </DialogFooter>
                </DialogContent>
              </Dialog>
            </div>
            {selectedCase && (
              <p className="text-sm text-muted-foreground truncate">
                Selected: {selectedCase.name}
              </p>
            )}
          </div>

          <div className="flex-1 overflow-y-auto p-3 space-y-2">
            {cases.length === 0 ? (
              <div className="text-center py-12">
                <Briefcase className="w-12 h-12 mx-auto text-muted-foreground/50 mb-3" />
                <p className="text-muted-foreground">No cases found</p>
                <p className="text-sm text-muted-foreground/70">Create a new case to get started</p>
              </div>
            ) : (
              cases.map((legalCase) => (
                <CaseCard
                  key={legalCase.id}
                  legalCase={legalCase}
                  isSelected={selectedCase?.id === legalCase.id}
                  onClick={() => setSelectedCase(legalCase)}
                />
              ))
            )}
          </div>
        </aside>

        {/* Main Content */}
        <main className="flex-1 overflow-hidden">
          {selectedCase ? (
            <CaseDetails legalCase={selectedCase} />
          ) : (
            <div className="h-full flex items-center justify-center">
              <div className="text-center">
                <Briefcase className="w-16 h-16 mx-auto text-muted-foreground/30 mb-4" />
                <p className="text-xl text-muted-foreground">Select a case to view details</p>
              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
