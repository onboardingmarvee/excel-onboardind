import { useState, useCallback } from 'react';
import { Upload, FileSpreadsheet, Loader2, Clock } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { supabase } from '@/integrations/supabase/client';
import { IMPORT_TYPE_LABELS, type ImportType } from '@/lib/templateProfiles';
import type { Tables } from '@/integrations/supabase/types';

interface UploadSectionProps {
  uploads: Tables<'uploads'>[];
  onUploadComplete: () => void;
  onSelectUpload: (upload: Tables<'uploads'>, importType: ImportType) => void;
}

export default function UploadSection({ uploads, onUploadComplete, onSelectUpload }: UploadSectionProps) {
  const [uploading, setUploading] = useState(false);
  const [importType, setImportType] = useState<ImportType | ''>('');
  const [dragOver, setDragOver] = useState(false);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      alert('Por favor, selecione um arquivo Excel (.xlsx ou .xls)');
      return;
    }
    if (!importType) {
      alert('Selecione o tipo de importação antes de enviar o arquivo.');
      return;
    }

    setUploading(true);
    try {
      const safeName = file.name
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')   // remove acentos
        .replace(/[^a-zA-Z0-9._-]/g, '_');                  // substitui espaços/parênteses/etc por _
      const path = `${crypto.randomUUID()}/${safeName}`;
      const { error: storageError } = await supabase.storage.from('uploads').upload(path, file);
      if (storageError) throw storageError;

      const { error: dbError } = await supabase.from('uploads').insert({
        original_filename: file.name,
        storage_path: path,
        status: 'uploaded',
      });
      if (dbError) throw dbError;

      onUploadComplete();
    } catch (err) {
      console.error('Upload error:', err);
      const msg = err instanceof Error ? err.message : (err as { message?: string })?.message ?? JSON.stringify(err);
      alert(`Erro ao fazer upload:\n${msg}`);
    } finally {
      setUploading(false);
    }
  }, [importType, onUploadComplete]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const statusClass = (status: string) => {
    const map: Record<string, string> = {
      uploaded: 'status-uploaded',
      processing: 'status-processing',
      done: 'status-done',
      error: 'status-error',
    };
    return map[status] || 'status-queued';
  };

  return (
    <div className="section-card space-y-5">
      <div className="flex items-center gap-2">
        <FileSpreadsheet className="h-5 w-5 text-primary" />
        <h2 className="text-lg font-semibold text-card-foreground">Upload de Excel</h2>
      </div>

      <div className="space-y-3">
        <label className="text-sm font-medium text-card-foreground">Tipo de Importação</label>
        <Select value={importType} onValueChange={(v) => setImportType(v as ImportType)}>
          <SelectTrigger>
            <SelectValue placeholder="Selecione o tipo..." />
          </SelectTrigger>
          <SelectContent>
            {Object.entries(IMPORT_TYPE_LABELS).map(([key, label]) => (
              <SelectItem key={key} value={key}>{label}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      <div
        className={`relative border-2 border-dashed rounded-lg p-8 text-center transition-colors cursor-pointer
          ${dragOver ? 'border-primary bg-primary/5' : 'border-border hover:border-primary/50'}
          ${!importType ? 'opacity-50 pointer-events-none' : ''}`}
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={handleDrop}
        onClick={() => {
          if (!importType) return;
          const input = document.createElement('input');
          input.type = 'file';
          input.accept = '.xlsx,.xls';
          input.onchange = (e) => {
            const file = (e.target as HTMLInputElement).files?.[0];
            if (file) handleFile(file);
          };
          input.click();
        }}
      >
        {uploading ? (
          <Loader2 className="h-8 w-8 mx-auto text-primary animate-spin" />
        ) : (
          <Upload className="h-8 w-8 mx-auto text-muted-foreground" />
        )}
        <p className="mt-2 text-sm text-muted-foreground">
          {uploading ? 'Enviando...' : 'Arraste um arquivo .xlsx ou clique para selecionar'}
        </p>
      </div>

      {uploads.length > 0 && (
        <div className="space-y-2">
          <h3 className="text-sm font-medium text-muted-foreground flex items-center gap-1.5">
            <Clock className="h-3.5 w-3.5" /> Uploads recentes
          </h3>
          <div className="space-y-1.5 max-h-48 overflow-y-auto">
            {uploads.map((u) => (
              <button
                key={u.id}
                disabled={!importType}
                onClick={() => importType && onSelectUpload(u, importType as ImportType)}
                className="w-full flex items-center justify-between gap-2 px-3 py-2 rounded-md text-sm
                  hover:bg-muted/60 transition-colors text-left disabled:opacity-40"
              >
                <span className="truncate text-card-foreground mono text-xs">{u.original_filename}</span>
                <span className={`status-badge ${statusClass(u.status)}`}>{u.status}</span>
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
