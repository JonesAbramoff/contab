Attribute VB_Name = "ErrosEST"
Option Explicit

'C�digos de Erro - Reservado de 7100 at� 7199
Public Const ERRO_LEITURA_ESTCONFIG2 = 7100 'Par�metros: %s chave %d FilialEmpresa
'Erro na leitura da tabela ESTConfig. Codigo = %s Filial = %i.
Public Const ERRO_ESTCONFIG_INEXISTENTE = 7101 'Par�metros: %s chave %d FilialEmpresa
'N�o foi encontrado registro em ESTConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_ESTCONFIG = 7102 'Par�metros: %s chave %d FilialEmpresa
'Erro na grava��o da tabela ESTConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_ESTCONFIG = 7103 'Sem par�metros
'Erro na leitura da tabela ESTConfig.
Public Const PRODUTO_INVENTARIADO_INEXISTENTE = 7104 'Sem Parametros
'N�o foi possivel abrir a tela, pois nenhum produto inventariado foi cadastrado
Public Const PRODUTO_INVENTARIADO_INEXISTENTE2 = 7105 'Parametro: sCodProduto
'O produto %s n�o � um produto inventariado.
Public Const ERRO_VALOR_PERCENTUALPERDA_INVALIDO = 7106 'Sem Parametros
'O Percentual de Perda n�o pode ser de 100%.
Public Const ERRO_TOTAL_DIFERENTE = 7107 'Par�metros: dTotalAux, dTotal
'O valor informado no campo Total %d � diferente da soma total dos valores da tela %d.
Public Const ERRO_OP_INICIAL_MAIOR = 7108
'A Ordem de Produ��o Inicial � maior que a Final.
Public Const ERRO_OP_INEXISTENTE = 7109
'A Ordem de Produ��o digitada n�o existe.
Public Const ERRO_OP_ABERTA_INEXISTENTE = 7110
'A Ordem de Produ��o digitada n�o � aberta ou n�o existe.
Public Const ERRO_OP_ENCERRADA_INEXISTENTE = 7111
'A Ordem de Produ��o digitada n�o � encerrada ou n�o existe.
Public Const ERRO_NAO_E_EMPRESATODA = 7112
'Para executar o relat�rio ordenado por Empresa, voc� dever� entrar no Sistema como Empresa Toda.
Public Const ERRO_ATUALIZACAO_ITENSNFISCAL = 7113 'Parametros: Filial, S�rie, N�mero, iTem
'Ocorreu um erro na tentativa de atualizar um registro na tabela de Itens de Nota Fiscal. Filial = %i, S�rie = %s, N�mero = %l, Item = %i
Public Const ERRO_FAIXA_INVALIDA2 = 7114 'Sem par�metros
'As Faixas de Classifica��o devem estar entre 0 e 100.
Public Const ERRO_FAIXA_MAXIMA2 = 7115 'Sem par�metros
'A soma dos valores das Faixas n�o pode ultrapassar o valor de 100.
Public Const ERRO_ITEMPEDIDO_INEXISTENTE1 = 7116 'Parametros lPedidoDeVenda, iFilialPedido, sProduto
'N�o existe um item do Pedido de Venda=%l, da Filial = %i com Produto=%s.
Public Const ERRO_NOTA_FISCAL_EXTERNA = 7117 'Sem par�metros
'N�o � poss�vel imprimir uma nota fiscal do tipo externa.
Public Const ERRO_RESERVA_BLOQUEIO_TOTAL_EXISTENTE = 7118 'Parametro: lPedidoDeVendas
'Pedido de Venda com c�digo %l tem Bloqueio Total de estoque.
Public Const ERRO_SEM_PRODUTOS_CTL_ESTOQUE = 7119 'sem parametros
'N�o h� nenhum produto para o qual seja feito controle de estoque cadastrado.
Public Const ERRO_NFISCALFATENTRADA_SEM_TITULO_PAGAR = 7120 'Parametro: lNumNotaFiscal
'Nota Fiscal Fatura de Entrada com n�mero %l n�o tem T�tulo a Pagar associado.
Public Const ERRO_TRANSFERENCIA_MESMO_ALMOXARIFADO = 7121 'Parametro iLinhaGrid
'A linha %i do grid indica uma transferencia envolvendo o mesmo Almoxarifado.
Public Const ERRO_EXCLUSAO_FILIALFORNFILEMP = 7123 'Sem Parametros
'Erro na Exclus�o da tabela FilialForFilEmp.
Public Const ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR_NF = 7124 'Par�metro: iParcela, dtDataVencimento, dtDataEmissao
'Em Cobran�a, Parcela n�mero %i tem Data de Vencimento %dt anterior � Data Emissao %dt.
Public Const ERRO_MODULO_COMPRAS_INATIVO = 7133 'Sem parametros
'O M�dulo de Compras est� inativo.
Public Const ERRO_MODULO_ESTOQUE_INATIVO = 7134 'Sem parametros
'O M�dulo de Estoque est� inativo.
Public Const ERRO_TIPO_RECEBIMENTO_COMPRAS = 7135 'Par�metros: iTipo
'O Recebimento com o tipo de c�digo %i � de Compras e s� deve ser utilizado
'nas notas fiscais de compras.
Public Const ERRO_RECEBIMENTO_DIFERENTE_COMPRAS = 7136 'Par�metros: iTipo
'O Recebimento com tipo de c�digo %i n�o � de Compras. S� deve ser
'utilizados nas Notas Fiscais que n�o s�o de Compras.
Public Const ERRO_TEMPOPRODUCAO_NAO_INFORMADO = 7137 'Sem parametros
'O tempo de produ��o n�o foi informado.
Public Const ERRO_TEMPOPRODUCAO_IGUAL_ZERO = 7138 'Sem parametros
'O tempo de produ��o n�o pode ser igual a zero.
Public Const ERRO_LOTE_NAO_TEM_PRODUTO_ALMOXARIFADO = 7139 'Parametros: sLote , sCodProduto, iAlmoxarifado
'O Lote / O.P. %s n�o trabalha com o Produto %s no Almoxarifado %s.
Public Const ERRO_NF_NAO_CADASTRADA3 = 7140 'Par�metros: lNumNotaFiscal, sSerie, iTipoNFiscal, lCliente, iFilialCli, dtDataEmissao, dtDataEntrada
'A nota fiscal com os dados abaixo n�o est� cadastrada. N�mero: %s, S�rie: %s, Tipo: %s, Cliente: %s, Filial: %s, Emiss�o em: %s e Entrada em: %s.
Public Const ERRO_NF_NAO_CADASTRADA1 = 7141 'Par�metros: lNumNotaFiscal, sSerie, iTipoNFiscal, lFornecedor, iFilialForn, dtDataEmissao, dtDataEntrada
'A nota fiscal com os dados abaixo n�o est� cadastrada. N�mero: %s, S�rie: %s, Tipo: %s, Fornecedor: %s, Filial: %s, Emiss�o em: %s e Entrada em: %s.



'C�digos de Avisos - Reservado de 5600 at� 5699
Public Const AVISO_CLASSIFICACAOABC_ALTERADA = 5600 'Sem Parametros
'Classifica��o ABC alterada, n�o � possivel visualizar a Curva ABC sem antes gravar. Deseja Gravar ?
Public Const AVISO_CONFIRMA_EXCLUSAO_INVLOTE = 5601 ' Par�metros: sCodigoInv
'Confirma a exclus�o do Invet�rio de c�digo = %s?
Public Const AVISO_ALTERACAO_NFISCAL_EXTERNA_CONTAB = 5602 'Par�metros: lFornecedor, iFilialForn, lCliente, iFilialCli, sSerie,lNumNotaFiscal,dtDataEmissao
'Nota Fiscal Externa com os dados C�digo do Fornecedor = %l, C�digo da Filial Fornecedor = %i, Cliente = %l, C�digo da Filial Cliente = %i, S�rie = %s, N�mero = %l, Data Emiss�o = %dt est� cadastrada, s� � possivel alterar os dados relativos a contabilidade. Deseja proseguir na altera��o?
Public Const AVISO_CANCELAR_NFISCALENTRADA = 5604 'Par�metro: lNumNotaFiscal
'Deseja realmente cancelar a Nota Fiscal de Entrada %l ?.
Public Const AVISO_FILIAIS_APURADAS_MESANO_DIFERENTES = 5605 'Sem Parametros
'O custo para a Empresa Toda pode n�o ser o custo correto, pois as Filiais da Empresa n�o est�o no mesmo m�s de Apura��o.
Public Const AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS2 = 5606 'Sem parametros
'Todos os campos com exce��o do Rastreamento (Lote / O.P. e Filial O.P.) n�o ser�o alterados. Deseja proseguir na altera��o ?




