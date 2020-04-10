Attribute VB_Name = "ErrosEST"
Option Explicit

'Códigos de Erro - Reservado de 7100 até 7199
Public Const ERRO_LEITURA_ESTCONFIG2 = 7100 'Parâmetros: %s chave %d FilialEmpresa
'Erro na leitura da tabela ESTConfig. Codigo = %s Filial = %i.
Public Const ERRO_ESTCONFIG_INEXISTENTE = 7101 'Parâmetros: %s chave %d FilialEmpresa
'Não foi encontrado registro em ESTConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_ESTCONFIG = 7102 'Parâmetros: %s chave %d FilialEmpresa
'Erro na gravação da tabela ESTConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_ESTCONFIG = 7103 'Sem parâmetros
'Erro na leitura da tabela ESTConfig.
Public Const PRODUTO_INVENTARIADO_INEXISTENTE = 7104 'Sem Parametros
'Não foi possivel abrir a tela, pois nenhum produto inventariado foi cadastrado
Public Const PRODUTO_INVENTARIADO_INEXISTENTE2 = 7105 'Parametro: sCodProduto
'O produto %s não é um produto inventariado.
Public Const ERRO_VALOR_PERCENTUALPERDA_INVALIDO = 7106 'Sem Parametros
'O Percentual de Perda não pode ser de 100%.
Public Const ERRO_TOTAL_DIFERENTE = 7107 'Parâmetros: dTotalAux, dTotal
'O valor informado no campo Total %d é diferente da soma total dos valores da tela %d.
Public Const ERRO_OP_INICIAL_MAIOR = 7108
'A Ordem de Produção Inicial é maior que a Final.
Public Const ERRO_OP_INEXISTENTE = 7109
'A Ordem de Produção digitada não existe.
Public Const ERRO_OP_ABERTA_INEXISTENTE = 7110
'A Ordem de Produção digitada não é aberta ou não existe.
Public Const ERRO_OP_ENCERRADA_INEXISTENTE = 7111
'A Ordem de Produção digitada não é encerrada ou não existe.
Public Const ERRO_NAO_E_EMPRESATODA = 7112
'Para executar o relatório ordenado por Empresa, você deverá entrar no Sistema como Empresa Toda.
Public Const ERRO_ATUALIZACAO_ITENSNFISCAL = 7113 'Parametros: Filial, Série, Número, iTem
'Ocorreu um erro na tentativa de atualizar um registro na tabela de Itens de Nota Fiscal. Filial = %i, Série = %s, Número = %l, Item = %i
Public Const ERRO_FAIXA_INVALIDA2 = 7114 'Sem parâmetros
'As Faixas de Classificação devem estar entre 0 e 100.
Public Const ERRO_FAIXA_MAXIMA2 = 7115 'Sem parâmetros
'A soma dos valores das Faixas não pode ultrapassar o valor de 100.
Public Const ERRO_ITEMPEDIDO_INEXISTENTE1 = 7116 'Parametros lPedidoDeVenda, iFilialPedido, sProduto
'Não existe um item do Pedido de Venda=%l, da Filial = %i com Produto=%s.
Public Const ERRO_NOTA_FISCAL_EXTERNA = 7117 'Sem parâmetros
'Não é possível imprimir uma nota fiscal do tipo externa.
Public Const ERRO_RESERVA_BLOQUEIO_TOTAL_EXISTENTE = 7118 'Parametro: lPedidoDeVendas
'Pedido de Venda com código %l tem Bloqueio Total de estoque.
Public Const ERRO_SEM_PRODUTOS_CTL_ESTOQUE = 7119 'sem parametros
'Não há nenhum produto para o qual seja feito controle de estoque cadastrado.
Public Const ERRO_NFISCALFATENTRADA_SEM_TITULO_PAGAR = 7120 'Parametro: lNumNotaFiscal
'Nota Fiscal Fatura de Entrada com número %l não tem Título a Pagar associado.
Public Const ERRO_TRANSFERENCIA_MESMO_ALMOXARIFADO = 7121 'Parametro iLinhaGrid
'A linha %i do grid indica uma transferencia envolvendo o mesmo Almoxarifado.
Public Const ERRO_EXCLUSAO_FILIALFORNFILEMP = 7123 'Sem Parametros
'Erro na Exclusão da tabela FilialForFilEmp.
Public Const ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR_NF = 7124 'Parâmetro: iParcela, dtDataVencimento, dtDataEmissao
'Em Cobrança, Parcela número %i tem Data de Vencimento %dt anterior à Data Emissao %dt.
Public Const ERRO_MODULO_COMPRAS_INATIVO = 7133 'Sem parametros
'O Módulo de Compras está inativo.
Public Const ERRO_MODULO_ESTOQUE_INATIVO = 7134 'Sem parametros
'O Módulo de Estoque está inativo.
Public Const ERRO_TIPO_RECEBIMENTO_COMPRAS = 7135 'Parâmetros: iTipo
'O Recebimento com o tipo de código %i é de Compras e só deve ser utilizado
'nas notas fiscais de compras.
Public Const ERRO_RECEBIMENTO_DIFERENTE_COMPRAS = 7136 'Parâmetros: iTipo
'O Recebimento com tipo de código %i não é de Compras. Só deve ser
'utilizados nas Notas Fiscais que não são de Compras.
Public Const ERRO_TEMPOPRODUCAO_NAO_INFORMADO = 7137 'Sem parametros
'O tempo de produção não foi informado.
Public Const ERRO_TEMPOPRODUCAO_IGUAL_ZERO = 7138 'Sem parametros
'O tempo de produção não pode ser igual a zero.
Public Const ERRO_LOTE_NAO_TEM_PRODUTO_ALMOXARIFADO = 7139 'Parametros: sLote , sCodProduto, iAlmoxarifado
'O Lote / O.P. %s não trabalha com o Produto %s no Almoxarifado %s.
Public Const ERRO_NF_NAO_CADASTRADA3 = 7140 'Parâmetros: lNumNotaFiscal, sSerie, iTipoNFiscal, lCliente, iFilialCli, dtDataEmissao, dtDataEntrada
'A nota fiscal com os dados abaixo não está cadastrada. Número: %s, Série: %s, Tipo: %s, Cliente: %s, Filial: %s, Emissão em: %s e Entrada em: %s.
Public Const ERRO_NF_NAO_CADASTRADA1 = 7141 'Parâmetros: lNumNotaFiscal, sSerie, iTipoNFiscal, lFornecedor, iFilialForn, dtDataEmissao, dtDataEntrada
'A nota fiscal com os dados abaixo não está cadastrada. Número: %s, Série: %s, Tipo: %s, Fornecedor: %s, Filial: %s, Emissão em: %s e Entrada em: %s.



'Códigos de Avisos - Reservado de 5600 até 5699
Public Const AVISO_CLASSIFICACAOABC_ALTERADA = 5600 'Sem Parametros
'Classificação ABC alterada, não é possivel visualizar a Curva ABC sem antes gravar. Deseja Gravar ?
Public Const AVISO_CONFIRMA_EXCLUSAO_INVLOTE = 5601 ' Parâmetros: sCodigoInv
'Confirma a exclusão do Invetário de código = %s?
Public Const AVISO_ALTERACAO_NFISCAL_EXTERNA_CONTAB = 5602 'Parâmetros: lFornecedor, iFilialForn, lCliente, iFilialCli, sSerie,lNumNotaFiscal,dtDataEmissao
'Nota Fiscal Externa com os dados Código do Fornecedor = %l, Código da Filial Fornecedor = %i, Cliente = %l, Código da Filial Cliente = %i, Série = %s, Número = %l, Data Emissão = %dt está cadastrada, só é possivel alterar os dados relativos a contabilidade. Deseja proseguir na alteração?
Public Const AVISO_CANCELAR_NFISCALENTRADA = 5604 'Parâmetro: lNumNotaFiscal
'Deseja realmente cancelar a Nota Fiscal de Entrada %l ?.
Public Const AVISO_FILIAIS_APURADAS_MESANO_DIFERENTES = 5605 'Sem Parametros
'O custo para a Empresa Toda pode não ser o custo correto, pois as Filiais da Empresa não estão no mesmo mês de Apuração.
Public Const AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS2 = 5606 'Sem parametros
'Todos os campos com exceção do Rastreamento (Lote / O.P. e Filial O.P.) não serão alterados. Deseja proseguir na alteração ?




