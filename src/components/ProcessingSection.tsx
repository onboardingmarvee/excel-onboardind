import { useState } from 'react';
import { Play, Loader2, Cog } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import { supabase } from '@/integrations/supabase/client';
import { IMPORT_TYPE_LABELS, type ImportType } from '@/lib/templateProfiles';
import type { Tables } from '@/integrations/supabase/types';

interface ProcessingSectionProps {
  selectedUpload: Tables<'uploads'> | null;
  importType: ImportType | null;
  onRunCreated: (run: Tables<'runs'>) => void;
}

export default function ProcessingSection({ selectedUpload, importType, onRunCreated }: ProcessingSectionProps) {
  const [instructions, setInstructions] = useState('');
  const [processing, setProcessing] = useState(false);
  const [contaPadrao, setContaPadrao] = useState('');
  const [centroCustoPadrao, setCentroCustoPadrao] = useState('');

  const handleProcess = async () => {
    if (!selectedUpload || !importType) return;
    setProcessing(true);

    try {
      // Determine which edge function to call based on import type
      const functionName = importType === 'clientes_fornecedores'
        ? 'process-clientes-fornecedores'
        : importType === 'vendas'
          ? 'process-vendas'
          : importType === 'contratos'
            ? 'process-contratos'
            : importType === 'movimentacoes'
              ? 'process-movimentacoes'
              : 'process-excel';

      if (importType === 'clientes_fornecedores' || importType === 'vendas' || importType === 'contratos' || importType === 'movimentacoes') {
        // New flow: edge function creates the run internally
        const { data, error } = await supabase.functions.invoke(functionName, {
          body: {
            upload_id: selectedUpload.id,
            ...(contaPadrao && { conta_padrao: contaPadrao }),
            ...(centroCustoPadrao && { centro_custo_padrao: centroCustoPadrao }),
          },
        });

        if (error) throw error;

        // Fetch the created run
        if (data?.run_id) {
          const { data: updatedRun } = await supabase
            .from('runs')
            .select('*')
            .eq('id', data.run_id)
            .single();

          if (updatedRun) onRunCreated(updatedRun);
        }
      } else {
        // Legacy flow for other import types
        const { data: run, error: runError } = await supabase.from('runs').insert({
          upload_id: selectedUpload.id,
          import_type: importType,
          transform_instructions: instructions || null,
          status: 'queued',
        }).select().single();

        if (runError) throw runError;

        const { data, error } = await supabase.functions.invoke(functionName, {
          body: { run_id: run.id },
        });

        if (error) throw error;

        const { data: updatedRun } = await supabase
          .from('runs')
          .select('*')
          .eq('id', run.id)
          .single();

        if (updatedRun) onRunCreated(updatedRun);
      }
    } catch (err) {
      console.error('Processing error:', err);
      alert('Erro ao processar. Verifique os logs.');
    } finally {
      setProcessing(false);
    }
  };

  const showInstructions = importType && importType !== 'clientes_fornecedores' && importType !== 'vendas' && importType !== 'contratos' && importType !== 'movimentacoes';

  return (
    <div className="section-card space-y-5">
      <div className="flex items-center gap-2">
        <Cog className="h-5 w-5 text-primary" />
        <h2 className="text-lg font-semibold text-card-foreground">Processamento</h2>
      </div>

      {selectedUpload ? (
        <div className="space-y-4">
          <div className="bg-muted/50 rounded-lg p-3 space-y-1">
            <p className="text-sm text-muted-foreground">Arquivo selecionado:</p>
            <p className="text-sm font-medium text-card-foreground mono">{selectedUpload.original_filename}</p>
            {importType && (
              <p className="text-xs text-primary font-medium">
                Template: {IMPORT_TYPE_LABELS[importType]}
              </p>
            )}
          </div>

          {showInstructions && (
            <div className="space-y-2">
              <label className="text-sm font-medium text-card-foreground">
                Instruções de transformação <span className="text-muted-foreground font-normal">(opcional)</span>
              </label>
              <Textarea
                placeholder="Ex: usar formato de data DD/MM/AAAA, separador decimal é vírgula, campos de nome em maiúsculas..."
                value={instructions}
                onChange={(e) => setInstructions(e.target.value)}
                rows={3}
                className="resize-none"
              />
            </div>
          )}

          {(importType === 'movimentacoes' || importType === 'vendas' || importType === 'contratos') && (
            <div className="border rounded-lg p-4 space-y-3 bg-muted/30">
              <p className="text-sm font-medium text-card-foreground">Configurações de importação</p>
              <div className="grid grid-cols-2 gap-3">
                <div className="space-y-1">
                  <label className="text-xs text-muted-foreground">Conta / Banco padrão</label>
                  <Input
                    placeholder="Ex: Nubank, Itaú, Bradesco..."
                    value={contaPadrao}
                    onChange={e => setContaPadrao(e.target.value)}
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs text-muted-foreground">Centro de custo padrão</label>
                  <Input
                    placeholder="Ex: Administrativo, Operacional..."
                    value={centroCustoPadrao}
                    onChange={e => setCentroCustoPadrao(e.target.value)}
                  />
                </div>
              </div>
            </div>
          )}

      {(importType === 'clientes_fornecedores' || importType === 'vendas' || importType === 'contratos' || importType === 'movimentacoes') && (
            <div className="bg-primary/5 border border-primary/20 rounded-lg p-3 text-xs text-muted-foreground space-y-1">
              <p className="font-medium text-primary">Template oficial registrado</p>
              <p>
                {importType === 'vendas'
                  ? 'O output será gerado no formato oficial com as abas "Dados" e "Serviços" preenchidas automaticamente.'
                  : importType === 'contratos'
                    ? 'O output será gerado no formato oficial com as abas "Contratos" e "Serviços" preenchidas automaticamente.'
                    : importType === 'movimentacoes'
                      ? 'O output será gerado no formato oficial com a aba "Dados" preenchida automaticamente (somente despesas).'
                      : 'O output será gerado no formato oficial com as abas "Dados" e "Contatos" preenchidas automaticamente.'}
              </p>
            </div>
          )}

          <Button
            onClick={handleProcess}
            disabled={processing}
            className="w-full"
            size="lg"
          >
            {processing ? (
              <>
                <Loader2 className="h-4 w-4 mr-2 animate-spin" />
                Processando...
              </>
            ) : (
              <>
                <Play className="h-4 w-4 mr-2" />
                Transferir Dados
              </>
            )}
          </Button>
        </div>
      ) : (
        <div className="text-center py-8 text-muted-foreground text-sm">
          <Cog className="h-10 w-10 mx-auto mb-2 opacity-30" />
          Selecione um upload e tipo de importação para começar
        </div>
      )}
    </div>
  );
}
