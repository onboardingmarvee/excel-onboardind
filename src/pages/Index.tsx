import { useState, useEffect, useCallback } from 'react';
import { FileSpreadsheet } from 'lucide-react';
import { supabase } from '@/integrations/supabase/client';
import type { Tables } from '@/integrations/supabase/types';
import type { ImportType } from '@/lib/templateProfiles';
import UploadSection from '@/components/UploadSection';
import ProcessingSection from '@/components/ProcessingSection';
import ResultsSection from '@/components/ResultsSection';

export default function Index() {
  const [uploads, setUploads] = useState<Tables<'uploads'>[]>([]);
  const [selectedUpload, setSelectedUpload] = useState<Tables<'uploads'> | null>(null);
  const [importType, setImportType] = useState<ImportType | null>(null);
  const [currentRun, setCurrentRun] = useState<Tables<'runs'> | null>(null);

  const fetchUploads = useCallback(async () => {
    const { data } = await supabase
      .from('uploads')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(20);
    if (data) setUploads(data);
  }, []);

  useEffect(() => {
    fetchUploads();
  }, [fetchUploads]);

  const handleSelectUpload = (upload: Tables<'uploads'>, type: ImportType) => {
    setSelectedUpload(upload);
    setImportType(type);
    setCurrentRun(null);
  };

  const handleCleanup = useCallback(() => {
    setSelectedUpload(null);
    setImportType(null);
    setCurrentRun(null);
    fetchUploads();
  }, [fetchUploads]);

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="border-b border-border bg-card">
        <div className="container max-w-7xl mx-auto px-4 py-4 flex items-center gap-3">
          <div className="h-9 w-9 rounded-lg bg-primary flex items-center justify-center">
            <FileSpreadsheet className="h-5 w-5 text-primary-foreground" />
          </div>
          <div>
            <h1 className="text-lg font-bold text-card-foreground">Migrador Excel</h1>
            <p className="text-xs text-muted-foreground">Ferramenta interna de migração de dados</p>
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="container max-w-7xl mx-auto px-4 py-6">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Col 1: Upload */}
          <div>
            <UploadSection
              uploads={uploads}
              onUploadComplete={fetchUploads}
              onSelectUpload={handleSelectUpload}
            />
          </div>

          {/* Col 2: Processing */}
          <div>
            <ProcessingSection
              selectedUpload={selectedUpload}
              importType={importType}
              onRunCreated={setCurrentRun}
            />
          </div>

          {/* Col 3: Results */}
          <div className="lg:col-span-1">
            <ResultsSection
              run={currentRun}
              selectedUpload={selectedUpload}
              onCleanup={handleCleanup}
            />
          </div>
        </div>
      </main>
    </div>
  );
}
