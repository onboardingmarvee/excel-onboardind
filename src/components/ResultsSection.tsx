import { useState } from 'react';
import { Download, AlertTriangle, CheckCircle2, Table2, AlertCircle, Trash2, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from '@/components/ui/alert-dialog';
import { supabase } from '@/integrations/supabase/client';
import type { Tables } from '@/integrations/supabase/types';

interface ResultsSectionProps {
  run: Tables<'runs'> | null;
  selectedUpload?: Tables<'uploads'> | null;
  onCleanup?: () => void;
}

interface ErrorItem {
  row?: number;
  field?: string;
  message?: string;
  documento?: string;
  tipo?: string;
  campo?: string;
  mensagem?: string;
}

export default function ResultsSection({ run, selectedUpload, onCleanup }: ResultsSectionProps) {
  const [cleaning, setCleaning] = useState(false);

  const handleCleanup = async () => {
    setCleaning(true);
    try {
      const { data, error } = await supabase.functions.invoke('cleanup-upload', {
        body: { clean_all: true },
      });
      if (error) throw error;
      if (data?.success) {
        onCleanup?.();
      } else {
        alert(data?.error || 'Erro ao limpar uploads');
      }
    } catch (err) {
      console.error('Cleanup error:', err);
      alert('Erro ao limpar uploads.');
    } finally {
      setCleaning(false);
    }
  };

  const cleanupButton = (
    <AlertDialog>
      <AlertDialogTrigger asChild>
        <Button
          variant="destructive"
          className="w-full"
          disabled={cleaning}
        >
          {cleaning ? (
            <>
              <Loader2 className="h-4 w-4 mr-2 animate-spin" />
              Limpando...
            </>
          ) : (
            <>
              <Trash2 className="h-4 w-4 mr-2" />
              Limpar uploads
            </>
          )}
        </Button>
      </AlertDialogTrigger>
      <AlertDialogContent>
        <AlertDialogHeader>
          <AlertDialogTitle>Limpar todos os uploads e outputs?</AlertDialogTitle>
          <AlertDialogDescription>
            Isso irá apagar <strong>todos</strong> os arquivos enviados, outputs gerados e registros de processamento.
            Os templates padrão de importação <strong>não serão afetados</strong>.
            Esta ação não pode ser desfeita.
          </AlertDialogDescription>
        </AlertDialogHeader>
        <AlertDialogFooter>
          <AlertDialogCancel>Cancelar</AlertDialogCancel>
          <AlertDialogAction onClick={handleCleanup}>
            Confirmar limpeza
          </AlertDialogAction>
        </AlertDialogFooter>
      </AlertDialogContent>
    </AlertDialog>
  );

  if (!run) {
    return (
      <div className="section-card space-y-5">
        <div className="flex items-center gap-2 mb-4">
          <Table2 className="h-5 w-5 text-primary" />
          <h2 className="text-lg font-semibold text-card-foreground">Prévia & Resultados</h2>
        </div>
        <div className="text-center py-12 text-muted-foreground text-sm">
          <Table2 className="h-10 w-10 mx-auto mb-2 opacity-30" />
          Execute um processamento para ver os resultados
        </div>
        {cleanupButton}
      </div>
    );
  }

  const preview = (run.preview_json as Record<string, unknown>[] | null) || [];
  const errors = (run.error_report_json as ErrorItem[] | null) || [];
  const columns = preview.length > 0 ? Object.keys(preview[0]) : [];

  const handleDownload = async (path: string | null, filename: string) => {
    if (!path) return;
    const { data, error } = await supabase.storage.from('outputs').download(path);
    if (error || !data) {
      alert('Erro ao baixar arquivo');
      return;
    }
    const url = URL.createObjectURL(data);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const statusIcon = run.status === 'done'
    ? <CheckCircle2 className="h-5 w-5 text-accent" />
    : run.status === 'error'
      ? <AlertCircle className="h-5 w-5 text-destructive" />
      : null;

  const showCleanup = run.status === 'done';

  return (
    <div className="section-card space-y-5">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2">
          <Table2 className="h-5 w-5 text-primary" />
          <h2 className="text-lg font-semibold text-card-foreground">Prévia & Resultados</h2>
        </div>
        <div className="flex items-center gap-2">
          {statusIcon}
          <span className={`status-badge status-${run.status}`}>{run.status}</span>
        </div>
      </div>

      {/* Preview Table */}
      {preview.length > 0 && (
        <div className="border border-border rounded-lg overflow-hidden">
          <div className="overflow-x-auto max-h-96">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-muted/60">
                  {columns.map((col) => (
                    <th key={col} className="px-3 py-2 text-left font-medium text-muted-foreground whitespace-nowrap mono">
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.map((row, i) => (
                  <tr key={i} className="border-t border-border hover:bg-muted/30 transition-colors">
                    {columns.map((col) => (
                      <td key={col} className="px-3 py-1.5 whitespace-nowrap text-card-foreground">
                        {String(row[col] ?? '')}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="bg-muted/40 px-3 py-1.5 text-xs text-muted-foreground">
            Mostrando {preview.length} linhas de prévia
          </div>
        </div>
      )}

      {/* Errors */}
      {errors.length > 0 && (
        <div className="space-y-2">
          <h3 className="text-sm font-medium text-destructive flex items-center gap-1.5">
            <AlertTriangle className="h-3.5 w-3.5" /> {errors.length} erro(s) encontrado(s)
          </h3>
          <div className="bg-destructive/5 border border-destructive/20 rounded-lg p-3 max-h-48 overflow-y-auto space-y-1.5">
            {errors.map((err, i) => (
              <div key={i} className="text-xs text-destructive">
                {err.documento && <span className="mono font-medium">[{err.documento}]</span>}
                {err.row != null && <span className="mono font-medium">Linha {err.row}</span>}
                {(err.field || err.campo) && <span className="mono"> → {err.field || err.campo}</span>}
                {(err.message || err.mensagem) && <span>: {err.message || err.mensagem}</span>}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Download Buttons */}
      {run.status === 'done' && (
        <div className="flex gap-3">
          <Button
            variant="outline"
            onClick={() => {
              const path = (run as any).output_storage_path || run.output_xlsx_path;
              handleDownload(path, 'output.xlsx');
            }}
            disabled={!(run as any).output_storage_path && !run.output_xlsx_path}
            className="flex-1"
          >
            <Download className="h-4 w-4 mr-2" /> Baixar XLSX
          </Button>
          {run.output_csv_path && (
            <Button
              variant="outline"
              onClick={() => handleDownload(run.output_csv_path, 'output.csv')}
              className="flex-1"
            >
              <Download className="h-4 w-4 mr-2" /> Baixar CSV
            </Button>
          )}
        </div>
      )}

      {/* Cleanup Button - always visible */}
      {cleanupButton}
    </div>
  );
}
