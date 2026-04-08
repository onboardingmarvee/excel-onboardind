export type ImportType = 'clientes_fornecedores' | 'movimentacoes' | 'vendas' | 'contratos';

export interface TemplateColumn {
  name: string;
  required: boolean;
  type: 'text' | 'number' | 'date' | 'phone' | 'currency' | 'email' | 'cpf_cnpj' | 'list' | 'integer';
  aliases: string[]; // possible names in client files
}

export interface TemplateProfile {
  label: string;
  columns: TemplateColumn[];
}

export const TEMPLATE_PROFILES: Record<ImportType, TemplateProfile> = {
  clientes_fornecedores: {
    label: 'Clientes/Fornecedores',
    columns: [
      { name: 'Tipo de Pessoa', required: true, type: 'list', aliases: ['tipo pessoa', 'tipo de pessoa', 'person type', 'tipo'] },
      { name: 'CPF/CNPJ', required: false, type: 'cpf_cnpj', aliases: ['cpf', 'cnpj', 'cpf/cnpj', 'cpf_cnpj', 'documento', 'doc'] },
      { name: 'Consultar Receita Federal?', required: false, type: 'list', aliases: ['consultar receita', 'receita federal', 'consulta rf'] },
      { name: 'Razão Social', required: true, type: 'text', aliases: ['razao social', 'razao', 'nome', 'name', 'empresa', 'company', 'razão social'] },
      { name: 'Nome Fantasia', required: false, type: 'text', aliases: ['nome fantasia', 'fantasia', 'trade name'] },
      { name: 'Cliente/Fornecedor', required: true, type: 'list', aliases: ['cliente fornecedor', 'cliente/fornecedor', 'tipo cadastro'] },
      { name: 'Status', required: false, type: 'list', aliases: ['status', 'situacao', 'estado'] },
      { name: 'CEP', required: false, type: 'text', aliases: ['cep', 'zip', 'codigo postal'] },
      { name: 'Número', required: false, type: 'integer', aliases: ['numero', 'num', 'nro', 'number'] },
      { name: 'Complemento', required: false, type: 'text', aliases: ['complemento', 'compl'] },
      { name: 'DDD', required: false, type: 'integer', aliases: ['ddd', 'codigo area'] },
      { name: 'Telefone', required: false, type: 'text', aliases: ['telefone', 'tel', 'fone', 'phone'] },
      { name: 'E-mail', required: false, type: 'email', aliases: ['email', 'e-mail', 'mail'] },
      { name: 'Tipo Entidade', required: false, type: 'list', aliases: ['tipo entidade', 'entidade', 'entity type'] },
      { name: 'Regime Tributário', required: false, type: 'list', aliases: ['regime tributario', 'regime tributário', 'tax regime'] },
      { name: 'Modalidade de Lucro', required: false, type: 'list', aliases: ['modalidade lucro', 'modalidade de lucro', 'profit mode'] },
      { name: 'Regime Especial', required: false, type: 'list', aliases: ['regime especial', 'special regime'] },
      { name: 'Inscrição Municipal', required: false, type: 'text', aliases: ['inscricao municipal', 'inscrição municipal', 'im'] },
      { name: 'Perfil de Tributação', required: false, type: 'list', aliases: ['perfil tributacao', 'perfil de tributação', 'tax profile'] },
      { name: 'CNAE Principal (Código)', required: false, type: 'text', aliases: ['cnae', 'cnae principal', 'codigo cnae'] },
      { name: 'Data Abertura/Nascimento', required: false, type: 'date', aliases: ['data abertura', 'data nascimento', 'abertura', 'nascimento', 'birthday'] },
      { name: 'Tipo Chave Pix', required: false, type: 'list', aliases: ['tipo chave pix', 'tipo pix', 'pix type'] },
      { name: 'Chave Pix', required: false, type: 'text', aliases: ['chave pix', 'pix', 'pix key'] },
      { name: 'Id Integração', required: false, type: 'text', aliases: ['id integracao', 'id integração', 'integration id'] },
    ],
  },
  movimentacoes: {
    label: 'Movimentações',
    columns: [
      { name: 'Data', required: true, type: 'date', aliases: ['data', 'dt', 'date', 'data movimento'] },
      { name: 'Tipo', required: true, type: 'text', aliases: ['tipo', 'type', 'natureza'] },
      { name: 'Descricao', required: true, type: 'text', aliases: ['descricao', 'desc', 'historico', 'description'] },
      { name: 'Valor', required: true, type: 'currency', aliases: ['valor', 'value', 'amount', 'vlr'] },
      { name: 'ContaDebito', required: false, type: 'text', aliases: ['conta debito', 'debito', 'debit'] },
      { name: 'ContaCredito', required: false, type: 'text', aliases: ['conta credito', 'credito', 'credit'] },
      { name: 'CentroCusto', required: false, type: 'text', aliases: ['centro de custo', 'centro custo', 'cc'] },
      { name: 'Documento', required: false, type: 'text', aliases: ['documento', 'doc', 'nf', 'nota fiscal'] },
      { name: 'Observacao', required: false, type: 'text', aliases: ['observacao', 'obs', 'nota'] },
    ],
  },
  vendas: {
    label: 'Vendas',
    columns: [
      { name: 'NumeroVenda', required: true, type: 'text', aliases: ['numero venda', 'num venda', 'nv', 'pedido', 'order'] },
      { name: 'DataVenda', required: true, type: 'date', aliases: ['data venda', 'data', 'dt venda', 'date'] },
      { name: 'Cliente', required: true, type: 'text', aliases: ['cliente', 'customer', 'comprador'] },
      { name: 'CPFCNPJCliente', required: false, type: 'cpf_cnpj', aliases: ['cpf cliente', 'cnpj cliente', 'cpf/cnpj', 'documento'] },
      { name: 'Produto', required: true, type: 'text', aliases: ['produto', 'item', 'product', 'descricao'] },
      { name: 'Quantidade', required: true, type: 'number', aliases: ['quantidade', 'qtd', 'qty', 'quant'] },
      { name: 'ValorUnitario', required: true, type: 'currency', aliases: ['valor unitario', 'preco', 'unit', 'price'] },
      { name: 'ValorTotal', required: true, type: 'currency', aliases: ['valor total', 'total', 'subtotal'] },
      { name: 'Desconto', required: false, type: 'currency', aliases: ['desconto', 'discount', 'desc'] },
      { name: 'FormaPagamento', required: false, type: 'text', aliases: ['forma pagamento', 'pagamento', 'payment'] },
      { name: 'Observacao', required: false, type: 'text', aliases: ['observacao', 'obs', 'nota'] },
    ],
  },
  contratos: {
    label: 'Contratos',
    columns: [
      { name: 'NumeroContrato', required: true, type: 'text', aliases: ['numero contrato', 'num contrato', 'contrato', 'contract'] },
      { name: 'Cliente', required: true, type: 'text', aliases: ['cliente', 'contratante', 'customer'] },
      { name: 'CPFCNPJCliente', required: false, type: 'cpf_cnpj', aliases: ['cpf', 'cnpj', 'cpf/cnpj', 'documento'] },
      { name: 'DataInicio', required: true, type: 'date', aliases: ['data inicio', 'inicio', 'start', 'dt inicio'] },
      { name: 'DataFim', required: false, type: 'date', aliases: ['data fim', 'fim', 'end', 'dt fim', 'vencimento'] },
      { name: 'ValorMensal', required: true, type: 'currency', aliases: ['valor mensal', 'mensalidade', 'monthly'] },
      { name: 'ValorTotal', required: false, type: 'currency', aliases: ['valor total', 'total'] },
      { name: 'Descricao', required: true, type: 'text', aliases: ['descricao', 'objeto', 'description'] },
      { name: 'Status', required: false, type: 'text', aliases: ['status', 'situacao', 'estado'] },
      { name: 'Observacao', required: false, type: 'text', aliases: ['observacao', 'obs', 'nota'] },
    ],
  },
};

export const IMPORT_TYPE_LABELS: Record<ImportType, string> = {
  clientes_fornecedores: 'Clientes/Fornecedores',
  movimentacoes: 'Movimentações',
  vendas: 'Vendas',
  contratos: 'Contratos',
};
