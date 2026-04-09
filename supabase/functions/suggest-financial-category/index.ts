import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

const STOPWORDS = new Set(["de","da","do","das","dos","para","com","sem","e","a","o","as","os","em","no","na","nos","nas","por","um","uma","uns","umas","ao","aos","que","se","ou"]);

function norm(text: string): string {
  return text.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ").trim();
}

function toTokens(text: string): string[] {
  const n = norm(text);
  return [...new Set(n.split(" ").filter(t => t.length >= 3 && !STOPWORDS.has(t)))];
}

function cleanCategoryQueryText(desc: string): string {
  let c = desc;
  c = c.replace(/^(PROVISÃO\s*-\s*|Cartão de Crédito\s*-\s*)/i, "");
  c = c.replace(/\b(provisao|provisão|cartao|cartão|credito|crédito)\b/gi, "");
  return c.replace(/\s+/g, " ").trim();
}

// --- ETAPA A: Overrides determinísticos ---

const MARKETING_SIGNALS = ["marketing","ads","google","meta","facebook","instagram","vendas","comissao","comissão","crm","leads","comercial","publicidade","propaganda","midia","mídia"];
const ADM_SIGNALS = ["administrativo","adm","financeiro","contabil","contábil","rh","recursos humanos","juridico","jurídico","escritorio","escritório"];
const SERVICO_SIGNALS = ["obra","construcao","construção","projeto","prestacao","prestação","cliente","instalacao","instalação","manutencao","manutenção"];

function hasAny(text: string, terms: string[]): boolean {
  const n = norm(text);
  return terms.some(t => n.includes(norm(t)));
}

function hasAnyToken(tokens: string[], terms: string[]): boolean {
  const normTerms = terms.map(t => norm(t));
  return tokens.some(tk => normTerms.some(nt => tk.includes(nt) || nt.includes(tk)));
}

function matchesAny(raw: string, cleaned: string, tokens: string[], terms: string[]): boolean {
  return hasAny(raw, terms) || hasAny(cleaned, terms) || hasAnyToken(tokens, terms);
}

function detectContext(raw: string, cleaned: string, tokens: string[]): "marketing" | "adm" | "servico" | "default" {
  if (matchesAny(raw, cleaned, tokens, MARKETING_SIGNALS)) return "marketing";
  if (matchesAny(raw, cleaned, tokens, ADM_SIGNALS)) return "adm";
  if (matchesAny(raw, cleaned, tokens, SERVICO_SIGNALS)) return "servico";
  return "default";
}

interface Override { code: string; name: string; reason: string }

function ruleBasedCategoryOverride(raw: string, cleaned: string, tokens: string[]): Override | null {
  const m = (terms: string[]) => matchesAny(raw, cleaned, tokens, terms);
  const ctx = () => detectContext(raw, cleaned, tokens);

  // ── 01. RECEITAS ──────────────────────────────────────────────────────────

  // Receitas recorrentes (fee mensal, mensalidade de serviço)
  if (m(["fee mensal","fee recorrente","mensalidade servico","mensalidade serviço","receita recorrente","recorrencia","recorrência"])) {
    return { code: "01.01.01", name: "Receitas Recorrentes (fee mensal)", reason: "override_receita_recorrente" };
  }
  // Receitas pontuais / jobs
  if (m(["job pontual","receita pontual","projeto pontual","venda servico","venda serviço","receita servico","receita serviço"])) {
    return { code: "01.01.02", name: "Receitas Pontuais (jobs)", reason: "override_receita_pontual" };
  }
  // Comissões recebidas
  if (m(["comissao recebida","comissão recebida","comissao sobre venda recebida"])) {
    return { code: "01.02.01", name: "Comissões", reason: "override_comissao_receita" };
  }

  // ── 02.01. IMPOSTOS SOBRE VENDAS E SERVIÇOS ───────────────────────────────

  if (m(["simples nacional","das simples","guia simples","pgdas","simples"])) {
    return { code: "02.01.01", name: "Simples Nacional", reason: "override_simples_nacional" };
  }
  if (m(["csll"])) {
    return { code: "02.01.02", name: "CSLL", reason: "override_csll" };
  }
  if (m(["icms"])) {
    return { code: "02.01.03", name: "ICMS", reason: "override_icms" };
  }
  if (m(["ipi"])) {
    return { code: "02.01.04", name: "IPI", reason: "override_ipi" };
  }
  if (m(["irpj"])) {
    return { code: "02.01.05", name: "IRPJ", reason: "override_irpj" };
  }
  if (m(["iss a recolher","iss proprio","imposto sobre servico","imposto sobre serviço","nota fiscal iss","iss municipal","iss retido a recolher"]) ||
      (m(["iss"]) && !m(["retido"]))) {
    return { code: "02.01.06", name: "ISS", reason: "override_iss" };
  }
  if (m(["pis a recolher","pis proprio"]) || (m(["pis"]) && !m(["csrf","retido"]))) {
    return { code: "02.01.07", name: "PIS", reason: "override_pis" };
  }
  if (m(["cofins a recolher","cofins proprio"]) || (m(["cofins"]) && !m(["csrf","retido"]))) {
    return { code: "02.01.08", name: "COFINS", reason: "override_cofins" };
  }
  if (m(["irrf a pagar","irrf retencao","irrf recolher"]) || (m(["irrf"]) && !m(["aplicacao","aplicação","financeira"]))) {
    return { code: "02.01.09", name: "IRRF a Pagar", reason: "override_irrf_pagar" };
  }
  if (m(["csrf","pis cofins csll","retidos pis","retencao pis cofins"])) {
    return { code: "02.01.10", name: "CSRF - Retidos Pis Cofins Csll a Recolher", reason: "override_csrf" };
  }
  if (m(["iss retido a recolher","iss retido"])) {
    return { code: "02.01.11", name: "ISS Retido a Recolher", reason: "override_iss_retido" };
  }
  if (m(["inss retido a recolher","inss retido","inss terceiros","retencao inss"])) {
    return { code: "02.01.12", name: "INSS Retido a Recolher", reason: "override_inss_retido" };
  }

  // ── 02.02. DESPESAS COM PRESTAÇÃO DE SERVIÇO ──────────────────────────────

  if (m(["prestador pontual","freelancer","autonomo","autônomo","pj externo","terceirizado","subcontratado"])) {
    return { code: "02.02.01", name: "Prestadores de Serviço Pontuais", reason: "override_prestador_pontual" };
  }
  if (m(["aluguel espaco","aluguel espaço","aluguel ambiente","locacao espaco","locação espaço","locacao ambiente","coworking"])) {
    return { code: "02.02.05", name: "Aluguel de Espaços e Ambientes", reason: "override_aluguel_espaco" };
  }

  // ── 02.03. SALÁRIOS, ENCARGOS E PESSOAL ──────────────────────────────────

  if (m(["pro-labore","pro labore","prolabore","remuneracao socio","remuneração sócio","socio pro","sócio pro"])) {
    return { code: "02.03.01", name: "Pró-labore", reason: "override_prolabore" };
  }
  const remuTerms = ["salario","salário","folha pagamento","clt","pagamento colaborador","remuneracao","remuneração","funcionario","funcionário","empregado","admissao","admissão","rescisao","rescisão"];
  if (m(remuTerms)) {
    const c = ctx();
    if (c === "marketing") return { code: "02.03.04", name: "Salários e Remunerações Marketing e Vendas", reason: "override_remuneracao_mkt" };
    if (c === "adm") return { code: "02.03.02", name: "Salários e Remunerações Administrativo", reason: "override_remuneracao_adm" };
    return { code: "02.03.03", name: "Salários e Remunerações Operacional", reason: "override_remuneracao_op" };
  }
  if (m(["inss patronal","inss empresa","inss folha","inss mensal","contribuicao previdenciaria","contribuição previdenciária"]) ||
      (m(["inss"]) && !m(["retido","retencao","recolher"]))) {
    return { code: "02.03.05", name: "INSS", reason: "override_inss" };
  }
  if (m(["fgts"])) {
    return { code: "02.03.06", name: "FGTS", reason: "override_fgts" };
  }
  if (m(["irrf folha","irrf salario","irrf funcionario"]) || (m(["irrf"]) && m(["funcionario","funcionário","colaborador","salario","salário","folha"]))) {
    return { code: "02.03.07", name: "IRRF", reason: "override_irrf_folha" };
  }
  if (m(["vale transporte","vt colaborador","vt folha"])) {
    return { code: "02.03.08", name: "Vale Transporte", reason: "override_vale_transporte" };
  }
  if (m(["vale refeicao","vale refeição","vale alimentacao","vale alimentação","vr folha","va folha"])) {
    return { code: "02.03.09", name: "Vale Refeição/Alimentação", reason: "override_vale_refeicao" };
  }
  if (m(["assistencia medica","assistência médica","plano saude","plano de saude","plano de saúde","plano medico","convenio medico","convênio médico"])) {
    return { code: "02.03.10", name: "Assistência Médica", reason: "override_assistencia_medica" };
  }
  if (m(["gratificacao meta","gratificação meta","bonus meta","bônus meta","premiacao","premiação"])) {
    return { code: "02.03.11", name: "Gratificação por Metas Atingidas", reason: "override_gratificacao" };
  }
  if (m(["plr","participacao nos lucros","participação nos lucros"])) {
    return { code: "02.03.12", name: "PLR (Participação nos Lucros e Resultados)", reason: "override_plr" };
  }

  // ── 02.04. DESPESAS COM PESSOAL ───────────────────────────────────────────

  if (m(["brinde pessoal","presente colaborador","presente funcionario","brinde colaborador","kit colaborador","cesta basica","cesta básica"])) {
    return { code: "02.04.01", name: "Brindes e Presentes para Pessoal", reason: "override_brinde_pessoal" };
  }
  if (m(["confraternizacao","confraternização","festa colaborador","happy hour","comemoracao interna","comemoração interna","jantar equipe","almoco equipe","almoço equipe"])) {
    return { code: "02.04.02", name: "Confraternizações", reason: "override_confraternizacao" };
  }
  if (m(["curso","treinamento","capacitacao","capacitação","workshop","formacao","formação","especializacao","especialização","pos graduacao","pós-graduação","certificacao","certificação","palestra"])) {
    return { code: "02.04.03", name: "Cursos e Treinamentos", reason: "override_curso_treinamento" };
  }
  if (m(["endomarketing","comunicacao interna","comunicação interna","clima organizacional"])) {
    return { code: "02.04.04", name: "Despesas com Endomarketing", reason: "override_endomarketing" };
  }
  if (m(["exame medico","exame admissional","exame demissional","medicina ocupacional","pcmso","aso"])) {
    return { code: "02.04.05", name: "Exames Médicos (ref. Medicina Ocupacional)", reason: "override_exame_medico" };
  }
  if (m(["uniforme","epi","equipamento protecao","equipamento proteção","vestuario profissional","vestuário profissional"])) {
    return { code: "02.04.07", name: "Uniformes", reason: "override_uniforme" };
  }

  // ── 02.05. DESPESAS COM MARKETING E VENDAS ────────────────────────────────

  if (m(["comissao sobre venda","comissão sobre venda","comissao vendedor","comissão vendedor","comissao representante","comissão representante"])) {
    return { code: "02.05.01", name: "Comissões sobre Vendas", reason: "override_comissao_vendas" };
  }
  if (m(["trafego pago","tráfego pago","google ads","meta ads","facebook ads","instagram ads","linkedin ads","ads campanha","impulsionamento","anuncio pago","anúncio pago"])) {
    return { code: "02.05.08", name: "Tráfego Pago", reason: "override_trafego_pago" };
  }
  if (m(["brinde cliente","presente cliente","kit cliente","mimo cliente"])) {
    return { code: "02.05.03", name: "Brindes e Presentes para Clientes", reason: "override_brinde_cliente" };
  }
  if (m(["assessoria marketing","assessoria de imprensa","relacoes publicas","relações públicas","agencia marketing","agência de marketing","assessoria comunicacao","assessoria comunicação"])) {
    return { code: "02.05.07", name: "Assessoria de Marketing e Imprensa", reason: "override_assessoria_marketing" };
  }
  if (m(["feira","evento comercial","evento marketing","exposicao","exposição","stand","booth"])) {
    return { code: "02.05.09", name: "Feiras e Eventos", reason: "override_feiras_eventos" };
  }
  if (m(["patrocinio","patrocínio","sponsor"])) {
    return { code: "02.05.10", name: "Patrocínios", reason: "override_patrocinio" };
  }

  // ── 02.06. DESPESAS ADMINISTRATIVAS ──────────────────────────────────────

  if (m(["financeiro por assinatura","assessoria financeira","consultoria financeira","gestor financeiro","gestao financeira","gestão financeira"])) {
    return { code: "02.06.01", name: "Assessoria Financeira", reason: "override_assessoria_financeira" };
  }
  if (m(["honorarios consultoria","honorários consultoria","consultoria empresarial","consultoria estrategica","consultoria estratégica"])) {
    return { code: "02.06.02", name: "Honorários Consultoria", reason: "override_honorarios_consultoria" };
  }
  if (m(["contador","contabilidade","honorarios contabilidade","honorários contabilidade","escritorio contabil","escritório contábil","contabil","contábil"])) {
    return { code: "02.06.03", name: "Honorários Contabilidade", reason: "override_contabilidade" };
  }
  if (m(["advogado","advocacia","juridico","jurídico","honorarios advocaticios","honorários advocatícios","assessoria juridica","assessoria jurídica","contrato juridico","contrato jurídico"])) {
    return { code: "02.06.04", name: "Honorários Advogado", reason: "override_advogado" };
  }
  if (m(["aluguel imovel","aluguel imóvel","aluguel escritorio","aluguel do escritório","locacao imovel","locação imóvel","aluguel sala","aluguel galpao","aluguel galpão"]) ||
      (m(["aluguel","locacao","locação"]) && !m(["espaco","espaço","ambiente","veiculo","veículo","carro"]))) {
    return { code: "02.06.05", name: "Aluguel", reason: "override_aluguel" };
  }
  if (m(["condominio","condomínio","taxa condominial","taxa condomínio"])) {
    return { code: "02.06.06", name: "Condomínio", reason: "override_condominio" };
  }
  if (m(["iptu","imposto predial","imposto territorial urbano"])) {
    return { code: "02.06.07", name: "IPTU", reason: "override_iptu" };
  }
  if (m(["agua","água","sabesp","caesb","embasa","saneamento","conta agua","conta água","fatura agua"])) {
    return { code: "02.06.08", name: "Água", reason: "override_agua" };
  }
  if (m(["energia eletrica","energia elétrica","luz","conta de luz","conta energia","eletricidade","enel","cemig","copel","celpe","coelba","celesc","elektro","equatorial","neoenergia"])) {
    return { code: "02.06.09", name: "Energia Elétrica", reason: "override_energia" };
  }
  if (m(["internet","telefone","telefonia","plano tel","conta tel","banda larga","vivo","claro","oi telecom","tim","net combo","operadora","fibra optica","fibra óptica"])) {
    return { code: "02.06.10", name: "Internet/Telefone", reason: "override_internet_telefone" };
  }
  if (m(["informatica","informática","suporte ti","suporte de ti","ti externo","servico ti","serviço de ti","manutencao computador","manutenção computador","rede informatica","rede informática"])) {
    return { code: "02.06.12", name: "Serviços de Informática e TI", reason: "override_ti" };
  }
  if (m(["material escritorio","material de escritório","papelaria","caneta","papel a4","toner","cartucho","impressao","impressão administrativa"])) {
    return { code: "02.06.13", name: "Material de Escritório", reason: "override_material_escritorio" };
  }
  if (m(["diarista","faxina","limpeza escritorio","limpeza do escritório","servico limpeza","serviço de limpeza","higienizacao","higienização","zeladoria"])) {
    return { code: "02.06.21", name: "Diarista/Limpeza", reason: "override_limpeza" };
  }
  if (m(["manutencao predial","manutenção predial","reparo imovel","reparo imóvel","reforma escritorio","reforma do escritório","obra escritorio","conserto predial"])) {
    return { code: "02.06.22", name: "Manutenção Predial", reason: "override_manutencao_predial" };
  }
  if (m(["motoboy","correios","sedex","pac correios","malote","documentos entrega","courrier","courier","entrega documento","cartorio","cartório"])) {
    return { code: "02.06.23", name: "Motoboy, Correios, Malotes, Documentos e afins", reason: "override_motoboy_correios" };
  }
  if (m(["seguro imovel","seguro imóvel","seguro veiculo","seguro veículo","seguro empresarial","seguro responsabilidade","seguro vida empresarial","seguro"])) {
    return { code: "02.06.17", name: "Seguros (Imóvel, Veículos e afins)", reason: "override_seguro" };
  }
  if (m(["alvara","alvará"])) {
    return { code: "02.06.18", name: "Alvarás", reason: "override_alvara" };
  }
  if (m(["licenca","licença","licença software","licenca operacional","licença operacional"])) {
    return { code: "02.06.19", name: "Licenças", reason: "override_licencas" };
  }
  if (m(["taxa administrativa","taxa cartorio","taxa cartório","taxa registro","taxa junta comercial","taxa receita federal","taxa prefeitura"])) {
    return { code: "02.06.20", name: "Outras Taxas Administrativas", reason: "override_taxas_adm" };
  }
  if (m(["higiene","limpeza material","produto limpeza","sabonete","papel higienico","papel higiênico","desinfetante","alcool gel","álcool gel"])) {
    return { code: "02.06.25", name: "Higiene e Limpeza", reason: "override_higiene_limpeza" };
  }
  if (m(["material uso","material de uso","material consumo","material de consumo","insumo","suprimento"])) {
    return { code: "02.06.26", name: "Material de Uso ou Consumo", reason: "override_material_consumo" };
  }
  if (m(["servico terceiro","serviço de terceiro","terceiros administrativo","prestacao servico administrativo","prestação serviço administrativo"])) {
    return { code: "02.06.27", name: "Serviços de Terceiros", reason: "override_servicos_terceiros" };
  }

  // ── 02.07. DESPESAS FINANCEIRAS ───────────────────────────────────────────

  if (m(["juros bancario","juros bancário","juros banco","juro emprestimo","juro empréstimo","encargo financeiro","mora bancaria","mora bancária","juros sobre divida","juros sobre dívida"])) {
    return { code: "02.07.01", name: "Juros Bancários", reason: "override_juros_bancarios" };
  }
  if (m(["tarifa bancaria","tarifa bancária","taxa bancaria","taxa bancária","ted tarifa","pix tarifa","manutencao conta","manutenção de conta","anuidade cartao","anuidade cartão"])) {
    return { code: "02.07.02", name: "Tarifas bancárias", reason: "override_tarifas_bancarias" };
  }
  if (m(["iof"])) {
    return { code: "02.07.03", name: "IOF", reason: "override_iof" };
  }
  if (m(["taxa maquininha","taxa gateway","taxa adquirente","taxa stone","taxa cielo","taxa pagseguro","taxa getnet","taxa rede","comissao gateway","comissão gateway","mdr"])) {
    return { code: "02.07.04", name: "Taxas sobre comissões", reason: "override_taxa_gateway" };
  }
  if (m(["multa paga","multa fiscal","multa administrativa","multa atraso","penalidade"])) {
    return { code: "02.07.05", name: "Multas Pagas", reason: "override_multa" };
  }
  if (m(["irrf aplicacao","irrf aplicação","irrf financeiro","imposto renda aplicacao","imposto renda aplicação"])) {
    return { code: "02.07.06", name: "IRRF s/ Aplicações Financeiras", reason: "override_irrf_aplicacao" };
  }
  if (m(["variacao cambial passiva","variação cambial passiva","perda cambial","cambio passivo","câmbio passivo"])) {
    return { code: "02.07.07", name: "Variações Cambiais Passivas", reason: "override_cambio_passivo" };
  }

  // ── 03. RECEITAS NÃO OPERACIONAIS ────────────────────────────────────────

  if (m(["aporte capital","aporte de capital","aporte socio","aporte sócio","capitalizacao","capitalização","integralizacao","integralização","aporte dos socios","aporte dos sócios"])) {
    return { code: "03.01.01", name: "Aporte de Capital dos Sócios", reason: "override_aporte_capital" };
  }
  if (m(["credito emprestimo bancario","crédito empréstimo bancário","emprestimo banco","empréstimo banco","caixa emprestimo recebido","financiamento recebido","bndes recebido","credito caixa","crédito caixa"])) {
    return { code: "03.02.01", name: "Crédito de Empréstimos Bancários", reason: "override_credito_emprestimo_banco" };
  }
  if (m(["emprestimo terceiro recebido","empréstimo terceiro recebido","credito emprestimo terceiro","crédito empréstimo terceiro"])) {
    return { code: "03.02.02", name: "Crédito de Empréstimos de Terceiros", reason: "override_credito_emprestimo_terceiro" };
  }
  if (m(["emprestimo socio recebido","empréstimo sócio recebido","credito emprestimo socio","crédito empréstimo sócio","socio emprestou","sócio emprestou"])) {
    return { code: "03.02.03", name: "Crédito de Empréstimos de Sócios", reason: "override_credito_emprestimo_socio" };
  }
  if (m(["reembolso cliente","reembolso de cliente","ressarcimento cliente","cliente reembolsou","cliente ressarcimento"])) {
    return { code: "03.03.01", name: "Reembolso de Clientes referente Compras e Pagamentos", reason: "override_reembolso_cliente" };
  }
  if (m(["adiantamento cliente","adiantamento de cliente","cliente adiantou","entrada cliente","sinal cliente"])) {
    return { code: "03.03.02", name: "Adiantamento de Clientes para Compras e Pagamentos", reason: "override_adiantamento_cliente" };
  }
  if (m(["rentabilidade","rendimento investimento","rendimento aplicacao","rendimento aplicação","juros aplicacao","juros aplicação","retorno investimento","yield"])) {
    return { code: "03.04.01", name: "Rentabilidade sobre Investimentos", reason: "override_rentabilidade" };
  }
  if (m(["devolucao socio","devolução sócio","socio devolveu","sócio devolveu","retorno emprestimo socio","retorno empréstimo sócio"])) {
    return { code: "03.04.02", name: "Devolução de Valores Emprestados para Sócios", reason: "override_devolucao_socio" };
  }
  if (m(["devolucao terceiro","devolução terceiro","terceiro devolveu","retorno emprestimo terceiro","retorno empréstimo terceiro"])) {
    return { code: "03.04.03", name: "Devolução de Valores Emprestados para Terceiros", reason: "override_devolucao_terceiro" };
  }
  if (m(["cashback","cash back","cashback recebido"])) {
    return { code: "03.04.06", name: "Cashback Recebido", reason: "override_cashback" };
  }
  if (m(["variacao cambial ativa","variação cambial ativa","ganho cambial","cambio ativo","câmbio ativo"])) {
    return { code: "03.04.09", name: "Variações Cambiais Ativas", reason: "override_cambio_ativo" };
  }
  if (m(["venda ativo","venda de ativo","venda bem","venda de bem","venda equipamento usado","venda mobiliario","venda mobiliário"])) {
    return { code: "03.04.04", name: "Venda de Bens de Pequeno Valor, Ativos e Outros", reason: "override_venda_ativo" };
  }
  if (m(["transferencia filial","transferência filial","transferencia unidade","transferência unidade","transferencia grupo","transferência grupo","transferencia empresa","transferência empresa"])) {
    return { code: "03.04.05", name: "Transferências de Filiais, Unidades e Empresas do Grupo", reason: "override_transferencia_filial_receita" };
  }

  // ── 04. DESPESAS NÃO OPERACIONAIS ────────────────────────────────────────

  if (m(["emprestimo bancario pago","empréstimo bancário pago","parcela emprestimo banco","parcela empréstimo banco","amortizacao banco","amortização banco","financiamento banco pago"])) {
    return { code: "04.01.01", name: "Empréstimos Bancários", reason: "override_emprestimo_banco" };
  }
  if (m(["emprestimo terceiro pago","empréstimo terceiro pago","parcela emprestimo terceiro","parcela empréstimo terceiro"])) {
    return { code: "04.01.02", name: "Empréstimos de Terceiros", reason: "override_emprestimo_terceiro" };
  }
  if (m(["emprestimo socio pago","empréstimo sócio pago","parcela emprestimo socio","parcela empréstimo sócio","devolucao emprestimo socio","devolução empréstimo sócio"])) {
    return { code: "04.01.03", name: "Empréstimos de Sócios", reason: "override_emprestimo_socio" };
  }
  if (m(["computador","notebook","laptop","monitor","periférico","periferico","impressora","scanner","hd externo","teclado","mouse profissional","servidor fisico","servidor físico"])) {
    return { code: "04.02.01", name: "Computadores e Periféricos", reason: "override_computador" };
  }
  if (m(["maquina","máquina","equipamento comprado","equipamento adquirido","ativo fixo","ativo imobilizado"])) {
    return { code: "04.02.02", name: "Máquinas e Equipamentos", reason: "override_maquinas" };
  }
  if (m(["movel","móvel","mesa comprada","cadeira comprada","armario comprado","armário comprado","utensilio","utensílio","instalacao fisica","instalação física","decoracao escritorio","decoração escritório"])) {
    return { code: "04.02.03", name: "Móveis, Utensílios e Instalações", reason: "override_moveis" };
  }
  if (m(["compra carro","compra veiculo","compra veículo","aquisicao veiculo","aquisição veículo","parcela carro","financiamento veiculo","financiamento veículo","automovel","automóvel"])) {
    return { code: "04.02.04", name: "Compra de Veículos", reason: "override_compra_veiculo" };
  }
  if (m(["compra imovel","compra imóvel","aquisicao imovel","aquisição imóvel","financiamento imovel","financiamento imóvel","parcela imovel","parcela imóvel"])) {
    return { code: "04.02.05", name: "Compra de Imóveis", reason: "override_compra_imovel" };
  }
  if (m(["compra cliente","pagamento cliente","despesa cliente reembolsar","pago para cliente","adiantamento para cliente"])) {
    return { code: "04.03.01", name: "Compras e Pagamentos para Clientes (Cliente vai Reembolsar)", reason: "override_despesa_cliente_reembolso" };
  }
  if (m(["compra adiantada cliente","pagamento adiantado cliente","despesa adiantada cliente","ja adiantado cliente","já adiantado cliente"])) {
    return { code: "04.03.02", name: "Compras e Pagamentos para Clientes (Já Adiantado pelo Cliente)", reason: "override_despesa_cliente_adiantado" };
  }
  if (m(["emprestimo para socio","empréstimo para sócio","adiantamento socio","adiantamento sócio","emprestei socio","emprestamos sócio"])) {
    return { code: "04.04.01", name: "Empréstimos para Sócios", reason: "override_emprestimo_para_socio" };
  }
  if (m(["emprestimo para terceiro","empréstimo para terceiro","adiantamento terceiro"])) {
    return { code: "04.04.02", name: "Empréstimos para Terceiros", reason: "override_emprestimo_para_terceiro" };
  }
  if (m(["transferencia para filial","transferência para filial","transferencia para unidade","transferência para unidade","transferencia para grupo","transferência para grupo"])) {
    return { code: "04.04.03", name: "Transferências para Filiais, Unidades e Empresas do Grupo", reason: "override_transferencia_filial_despesa" };
  }
  if (m(["doacao","doação","filantropia","contribuicao social","contribuição social","responsabilidade social"])) {
    return { code: "04.04.04", name: "Doações", reason: "override_doacao" };
  }

  // ── 05. DISTRIBUIÇÃO DE LUCROS ────────────────────────────────────────────

  if (m(["distribuicao lucro","distribuição lucro","dividendo","dividendo mensal","lucro distribuido","lucro distribuído","retirada socio","retirada sócio","retirada mensal socio"])) {
    return { code: "05.01.03", name: "Distribuição de Lucros / Dividendos", reason: "override_distribuicao_lucros" };
  }
  if (m(["despesa pessoal socio","despesa pessoal sócio","despesa particular socio","despesa particular sócio","gasto pessoal socio","gasto pessoal sócio"])) {
    return { code: "05.01.02", name: "Despesas Pessoais dos Sócios", reason: "override_despesa_pessoal_socio" };
  }
  if (m(["remuneracao mensal socio","remuneração mensal sócio","remuneracao dos socios","remuneração dos sócios"])) {
    return { code: "05.01.01", name: "Remuneração Mensal dos Sócios", reason: "override_remuneracao_mensal_socio" };
  }

  // ── CONTEXTO: Ferramentas / Software (com contexto) ────── (fim dos overrides específicos)

  const ferrTerms = ["ferramenta","software","saas","plataforma","sistema","app","aplicativo","subscricao","subscrição"];
  if (m(ferrTerms)) {
    const c = ctx();
    if (c === "marketing") return { code: "02.05.02", name: "Ferramentas, Softwares e Sistemas - Comercial e Marketing", reason: "override_ferramentas_mkt" };
    if (c === "adm") return { code: "02.06.11", name: "Ferramentas, Softwares e Sistemas - Administrativo/Financeiro", reason: "override_ferramentas_adm" };
    return { code: "02.02.06", name: "Ferramentas, Softwares e Sistemas - Prestação de Serviço", reason: "override_ferramentas_op" };
  }

  // ── CONTEXTO: Combustível / Veículo (com contexto) ────────────────────────

  if (m(["combustivel","combustível","gasolina","etanol","diesel","abastecimento","pedagio","pedágio","estacionamento","uber","99app","taxi"])) {
    const c = ctx();
    if (c === "marketing") return { code: "02.05.05", name: "Combustível e Despesas com Veículos (Visitas Comerciais)", reason: "override_combustivel_mkt" };
    if (c === "servico") return { code: "02.02.03", name: "Combustível e Despesas com Veículos para Prestação de Serviço", reason: "override_combustivel_servico" };
    return { code: "02.06.15", name: "Combustível e Despesas com Veículos (Administrativo)", reason: "override_combustivel_adm" };
  }

  // ── CONTEXTO: Lanche / Refeição (com contexto) ────────────────────────────

  if (m(["lanche","refeicao","refeição","almoco","almoço","jantar","restaurante","ifood","rappi","alimentacao reuniao","alimentação reunião","coffee break","cafe","café reuniao"])) {
    const c = ctx();
    if (c === "marketing") return { code: "02.05.04", name: "Lanches e Refeições (Visitas Comerciais)", reason: "override_lanche_mkt" };
    if (c === "servico") return { code: "02.02.02", name: "Lanches e Refeições para Prestação de Serviço", reason: "override_lanche_servico" };
    return { code: "02.06.14", name: "Lanches e Refeições (Administrativo)", reason: "override_lanche_adm" };
  }

  // ── CONTEXTO: Hospedagem / Viagem (com contexto) ──────────────────────────

  if (m(["hospedagem","hotel","airbnb","passagem aerea","passagem aérea","voo","milhas","viagem","flight","translado","transfer aeroporto"])) {
    const c = ctx();
    if (c === "marketing") return { code: "02.05.06", name: "Hospedagem e Outras Despesas com Viagem (Visitas Comerciais)", reason: "override_hospedagem_mkt" };
    if (c === "servico") return { code: "02.02.04", name: "Hospedagem e Outras Despesas com Viagem para Prestação de Serviço", reason: "override_hospedagem_servico" };
    return { code: "02.06.16", name: "Hospedagem e Outras Despesas com Viagem (Administrativo)", reason: "override_hospedagem_adm" };
  }

  return null;
}

// --- Main handler ---

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { description } = await req.json();
    if (!description || typeof description !== "string") {
      return new Response(JSON.stringify({ error: "description is required" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const raw = description;
    const cleaned = cleanCategoryQueryText(raw);
    const tokens = toTokens(cleaned);

    // ETAPA A: Override
    const override = ruleBasedCategoryOverride(raw, cleaned, tokens);
    if (override) {
      return new Response(JSON.stringify({
        code: override.code,
        name: override.name,
        score: 1,
        top: [{ code: override.code, name: override.name, score: 1 }],
        strategy: "override",
        reason: override.reason,
        cleaned_description: cleaned,
      }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }

    // ETAPA B: Similaridade (fallback)
    if (tokens.length === 0) {
      return new Response(JSON.stringify({
        code: null, name: null, score: 0, top: [],
        strategy: "similarity", reason: null, cleaned_description: cleaned,
      }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }

    const supabaseUrl = Deno.env.get("SUPABASE_URL")!;
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
    const sb = createClient(supabaseUrl, serviceKey);

    const { data: candidates, error } = await sb
      .from("financial_categories")
      .select("code, name, full_label, tokens")
      .overlaps("tokens", tokens);

    if (error) throw error;

    const codeMatch = raw.match(/\d{2}\.\d{2}\.\d{2}/);

    const scored = (candidates || []).map((c: any) => {
      const catTokens: string[] = c.tokens;
      const intersection = tokens.filter(t => catTokens.includes(t)).length;
      let score = (2 * intersection) / (catTokens.length + tokens.length);
      if (codeMatch && c.code === codeMatch[0]) score += 0.3;
      return { code: c.code, name: c.name, score: Math.round(score * 1000) / 1000 };
    });

    scored.sort((a: any, b: any) => b.score - a.score);
    const top = scored.slice(0, 5);
    const best = top[0] || { code: null, name: null, score: 0 };

    // ETAPA C: Guardrails
    const warning = best.score > 0 && best.score < 0.25 ? "categoria_suspeita" : undefined;

    return new Response(JSON.stringify({
      code: best.code,
      name: best.name,
      score: best.score,
      top,
      strategy: "similarity",
      reason: warning || null,
      cleaned_description: cleaned,
    }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
  } catch (err) {
    console.error("Suggest error:", err);
    return new Response(JSON.stringify({ error: (err as Error).message }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
