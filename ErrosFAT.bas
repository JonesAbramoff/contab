Attribute VB_Name = "ErrosFAT"
Option Explicit

'C�digos de Erro - Reservado de 8000 a 8299
Public Const ERRO_LEITURA_FATCONFIG2 = 8000 '%s chave %d FilialEmpresa
'Erro na leitura da tabela FATConfig. Codigo = %s Filial = %i
Public Const ERRO_FATCONFIG_INEXISTENTE = 8001 '%s chave %d FilialEmpresa
'N�o foi encontrado registro em FATConfig. Codigo = %s Filial = %i
Public Const ERRO_ATUALIZACAO_FATCONFIG = 8002 '%s chave %d FilialEmpresa
'Erro na grava��o da tabela FATConfig. Codigo = %s Filial = %i
Public Const ERRO_SEM_BLOQUEIOS_PV_SEL = 8004 'Sem par�metro
'N�o h� bloqueios dentro dos crit�rios de sele��o informados.
Public Const ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_NF = 8005 'Sem par�metro
'Erro na leitura de pedidos de venda para a gera��o de notas fiscais
Public Const ERRO_SEM_PEDIDOS_VENDA_ENCONTRADOS = 8006 'Sem par�metro
'N�o h� pedidos de venda dentro dos crit�rios de sele��o informados.
Public Const ERRO_SEM_NFSREC_GERACAO_FATURA = 8007 'Sem par�metro
'N�o h� notas fiscais a serem faturadas dentro dos crit�rios de sele��o informados.
Public Const ERRO_LEITURA_NFSREC_GERACAO_FATURA = 8008 'Sem par�metro
'Erro na leitura de notas fiscais para gera��o de fatura
Public Const ERRO_SERIE_NAO_PREENCHIDA = 8012
'A Serie deve ser preenchida.
Public Const ERRO_ATUALIZACAO_COMISSOESPEDVENDAS = 8013 'Parametro: lNumIntDoc
'Erro na tentativa de atualizar as Comiss�es do Pedido de Vendas de N�mero Interno %l
Public Const ERRO_VALOREMISSAO_COMISSAO_NAO_INFORMADO = 8014 'iComissao
'O Valor de Emiss�o da Comissao %i do T�tulo n�o foi informado.
Public Const ERRO_INCLUSAO_COMISSOESPEDVENDAS = 8015 'Sem parametro
'Erro na tentativa incluir registro na tabela de Comiss�es de Pedido de Vendas.
Public Const ERRO_VALOR_EMISSAO_MAIOR = 8017 'Sem Par�metros
'Valor da comiss�o na emiss�o n�o pode ser maior que o valor total da comiss�o.
Public Const ERRO_PERCENTUAL_EMISSAO_NAO_INFORMADO = 8020 'iComissao
'O Percentual de Emiss�o da Comiss�o %i n�o foi informado.
Public Const ERRO_PEDIDO_VENDA_BAIXADO_MODIFICACAO = 8021 'Parametro: lCodigo
'O Pedido de Venda n�mero %l j� foi faturado. N�o pode ser modificado.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI = 8022
'O item de c�digo n�o existe na tabela de Tipos de Tributa��o para IPI.
Public Const ERRO_FUNDAMENTACAO_NAO_PREENCHIDA = 8023 'Sem par�metro
'O preenchimento da Fundamenta��o � obrigat�rio.
Public Const ERRO_DESTINO_NAO_PREENCHIDO = 8024 'Sem par�metro
'O preenchimento do destino � obrigat�rio.
Public Const ERRO_LEITURA_CATEGORIAPRODUTOITEM1 = 8025 'Par�metro: sCategoria
'Erro na leitura dos registros da categoria %s da tabela Categoria Produto.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS = 8028 'Par�metro: iCodigo
'O item de c�digo %i n�o existe na tabela de Tipos de Tributa��o para ICMS.
Public Const ERRO_LOCK_ESTADOS = 8029  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de itens das Categorias de Clientes.
Public Const ERRO_PEDVENDA_FATURADO_ALTERACAO_CLIENTE = 8037 'Parametros: lCodigoPV, lClienteBD, lClienteTela
'Pedido de Venda %l tem �tens faturados. N�o � poss�vel alterar Cliente %l no Banco de Dados para %l da Tela.
Public Const ERRO_PEDVENDA_FATURADO_ALTERACAO_FILIAL_CLIENTE = 8038 'Parametros: lCodigoPV, iFilialBD, iFilialTela
'Pedido de Venda %l tem �tens faturados. N�o � poss�vel alterar Filial Cliente %i no Banco de Dados para %i da Tela.
Public Const ERRO_PEDVENDA_FATURADO_ALTERACAO_NATUREZA = 8039 'Par�metros: lCodigoPV, sNaturezaBD, sNaturezaTela
'Pedido de Venda %l tem �tens faturados. N�o � poss�vel alterar NaturezaFilial Cliente %i no Banco de Dados para %i da Tela.
Public Const ERRO_LEITURA_FATCONFIG = 8042 'Parametro Codigo
'Erro na leitura da tabela FATConfig. Codigo = %s.
Public Const ERRO_LEITURA_BLOQUEIOPV = 8045 'Parametros: iFilialEmpresa, lPedidoDeVendas, iTipoDeBloqueio
'Ocorreu um erro na leitura de um registro da tabela de Bloqueios de Pedido de Venda. Filial=%i, Pedido=%l, Tipo de Bloqueio=%i.
Public Const ERRO_LOCK_BLOQUEIOPV = 8046 'Parametros: iFilialEmpresa, lPedidoDeVendas, iSequencial
'Ocorreu um erro na tentativa de fazer um "lock" de um registro da tabela de Bloqueios de Pedido de Venda. Filial=%i, Pedido=%l, Sequencial=%i.
Public Const ERRO_LEITURA_CANALVENDA = 8047 'Sem parametros
'Erro de Leitura na tabela CanalVenda
Public Const ERRO_NOME_REDUZIDO_CANALVENDA_REPETIDO = 8050 'Sem parametros
'Erro Nome reduzido j� � utilizado por canal com outro C�digo
Public Const ERRO_INSERCAO_CANALVENDA = 8051 'Parametro objCanal.iCodigo
'Erro na inser��o do Canal %i
Public Const ERRO_ATUALIZACAO_CANALVENDA = 8052 'Parametro objCanal.iCodigo
'Erro na Atualiza��o do Canal %i
Public Const ERRO_EXCLUSAO_CANALVENDA = 8053 'Parametro objCanal.iCodigo
'Erro exclus�o Canal de Venda % i
Public Const ERRO_DESCRICAONF_NAO_PREENCHIDA = 8054 'Sem parametros
'O preenchimento da Descric�o da Nota Fiscal � obrigat�rio.
Public Const ERRO_LEITURA_TIPODEBLOQUEIO = 8055 'Sem parametros
'Erro de Leitura na tabela TipodeBloqueio
Public Const ERRO_LEITURA_TIPODEBLOQUEIO1 = 8056 'Parametros objTipo.iCodigo
'Erro na leitura do Tipo %i da tabela Tipo de Bloqueio
Public Const ERRO_NOME_REDUZIDO_TIPODEBLOQUEIO_REPETIDO = 8058 'Sem parametros
'Erro Nome reduzido j� � utilizado por Tipo com outro C�digo
Public Const ERRO_INSERCAO_TIPODEBLOQUEIO = 8059 'Parametro objTipo.iCodigo
'Erro na inser��o do Tipo %i
Public Const ERRO_ATUALIZACAO_TIPODEBLOQUEIO = 8060 'Parametro objTipo.iCodigo
'Erro na Atualiza��o do Tipo %i
Public Const ERRO_EXCLUSAO_TIPODEBLOQUEIO = 8061 'Parametro objTipo.iCodigo
'Erro na Exclus�o do Tipo %i
Public Const ERRO_LEITURA_TRIBNFISCAL = 8065
'Erro na leitura da tabela de Tributa��o de Nota Fiscal.
Public Const ERRO_LEITURA_TRIBCOMPLNFISCAL = 8066
'Erro na leitura da tabela de Tributa��o de Complemento de Nota Fiscal.
Public Const ERRO_LEITURA_TRIBITEMNFISCAL = 8067
'Erro na leitura da tabela de Tributa��o de Itens de Nota Fiscal.
Public Const ERRO_INCONSISTENCIA_TRIBITEMNFISCAL = 8068
'Tributa��o para o Item n�o confere com o existente.
Public Const ERRO_IR_FONTE_MAIOR_VALOR_TOTAL = 8069 'dValorIRRF, dValorTotal
'IR Fonte = %d n�o pode ultrapassar Valor Total = %d.
Public Const ERRO_ATUALIZACAO_PREVVENDA = 8071 'Par�metro sCodigo
'Erro na atualiza��o da Previs�o de Venda com o c�digo %s na tabela PrevVenda .
Public Const ERRO_INSERCAO_PREVVENDA = 8072 'Par�metro sCodigo
'Erro na inser��o da Previs�o de Venda com o c�digo %s na tabela PrevVenda .
Public Const ERRO_EXCLUSAO_PREVVENDA = 8073 'Par�metro sCodigo
'Erro na exclus�o da Previs�o de Venda com o c�digo %s na tabela PrevVenda .
Public Const ERRO_PREVVENDA_NAO_CADASTRADA = 8074 'Par�metro sCodigo
'Previs�o de Venda %s n�o cadastrada na tabela PrevVenda .
Public Const ERRO_LOCK_PREVVENDA = 8075 'Par�metro sCodigo
'N�o foi poss�vel fazer o Lock da Previs�o de Venda %s da tabela PrevVenda .
Public Const ERRO_CODIGO_PREVVENDA_NAO_PREENCHIDO = 8077
'O C�digo da Previs�o de Venda n�o est� preenchido.
Public Const ERRO_DATAPREVISAO_NAO_PREENCHIDA = 8078
'O campo Data de Previs�o n�o est� preenchido .
Public Const ERRO_PEDIDODEVENDA_INEXISTENTE = 8082 'Parametros = lCodPedido
'O Pedido de Venda %l n�o est� cadastrado.
Public Const ERRO_LOCKEXCLUSIVE_TIPODEBLOQUEIO = 8083 'Sem Par�metros
'N�o foi poss�vel fazer LockExclusive na tabela TiposDeBloqueio
Public Const ERRO_TIPODEBLOQUEIO_USADO = 8084  'Par�metro, o c�digo do Tipo de Bloqueio
'N�o � permitido excluir o Tipo de Bloqueio %i, pois ele � utilizado em Bloqueios de Pedido de Venda. Pedido de Venda = %l, Filial Empresa = %i e Sequencial = %l.
Public Const ERRO_LEITURA_BLOQUEIOSPV_BLOQUEIOSPVBAIXADOS = 8086 'Sem Par�metros
'Erro de Leitura na tabela BloqueiosPV ou BloqueiosPVBaixados
Public Const ERRO_LEITURA_PEDIDOSDEVENDA_PEDIDOSDEVENDABAIXADOS = 8087 'Sem Par�metros
'Erro de Leitura na Tabela PedidosDeVenda ou PedidosDeVendaBaixados
Public Const ERRO_LEITURA_NFISCAL_NFISCALBAIXADAS = 8088 'Sem par�mtros
'Erro de Leitura na Tabela NFiscal ou NFiscalBaixadas
Public Const ERRO_LOCKEXCLUSIVE_CANALVENDA = 8089 'Sem Par�metros
'N�o foi poss�vel fazer LockExclusive na Tabel CanalVenda
Public Const ERRO_CANAL_EM_PV = 8090 'Par�metro C�digo do Canal de Venda
'N�o � permitido excluir o Canal de Venda %i, pois ele � utilizado no Pedido de Venda %l da Filial Empresa %i.
Public Const ERRO_CANAL_EM_NF = 8091 'Par�metro C�digo do Canal de Venda
'N�o � permitido excluir o Canal de Venda %i, pois ele est� sendo utilizado na Nota Fiscal com N�mero %l, S�rie %s da Filial Empresa %i.
Public Const ERRO_LIBERACAOCREDITO_INEXISTENTE = 8093 'Parametros = sCodUsuario
'O usu�rio %s n�o tem autoriza��o para liberar por cr�dito este bloqueio.
Public Const ERRO_LIBERACAOCREDITO_LIMITEOPERACAO = 8094 'Parametros = sCodUsuario
'O usu�rio %s n�o tem autoriza��o para liberar por cr�dito este bloqueio (Limite por Opera��o Excedido).
Public Const ERRO_LIBERACAOCREDITO_LIMITEMENSAL = 8095 'Parametros = sCodUsuario
'O usu�rio %s n�o tem autoriza��o para liberar por cr�dito este bloqueio (Limite Mensal Excedido).
Public Const ERRO_ALTERACAO_CLIENTES = 8096 'Parametro lCodigo
'Erro na alteracao da tabela de Clientes. C�digo do Cliente = %l.
Public Const ERRO_LEITURA_LIBERACAOCREDITO = 8097 'Parametro: sCodUsuario
'Ocorreu um erro na leitura de um registro da tabela de Libera��o de Cr�dito. Usu�rio = %s.
Public Const ERRO_LEITURA_VALORLIBERADOCREDITO = 8098 'Parametro: sCodUsuario, iAno
'Ocorreu um erro na leitura de um registro da tabela ValorLiberadoCredito. Usu�rio = %s, Ano = %i.
Public Const ERRO_EXCLUSAO_VALORLIBERADOCREDITO = 8099
'Erro na tentativa de excluir registro da tabela ValorLiberadoCredito.
Public Const ERRO_LOCK_LIBERACAOCREDITO = 8100
'Erro na tentativa de fazer 'lock' na tabela LiberacaoCredito.
Public Const ERRO_LIMITES = 8101
'Limite de Opera��o precisa ser menor ou igual ao Limite Mensal.
Public Const ERRO_ATUALIZACAO_LIBERACAOCREDITO = 8102
'Erro na tentativa de atualizar a tabela LiberacaoCredito.
Public Const ERRO_INSERCAO_LIBERACAOCREDITO = 8103
'Erro na tentiva de inserir um registro na tabela LiberacaoCredito.
Public Const ERRO_EXCLUSAO_LIBERACAOCREDITO = 8104 'Parametro codigo usuario
'Erro na tentativa de exclus�o da al�ada do usu�rio %s da tabela LiberacaoCredito.
Public Const ERRO_VINCULADO_PEDIDOSVENDA = 8105 'Parametro codigo usuario.
'A al�ada n�o pode ser exclu�da pois o usu�rio %s est� vinculado a um Pedido de Venda.
Public Const ERRO_VINCULADO_VALORLIBERADOCREDITO = 8106 'Parametro codigo usuario
'A al�ada n�o pode ser exclu�da pois o usu�rio %s est� vinculado a um registro na tabela de Libera��o de Cr�ditos.
Public Const ERRO_LIBERACAOCREDITO_VAZIA = 8109
'Erro tabela LiberacaoCredito est� Vazia.
Public Const ERRO_LEITURA_TABELA_ALMOXARIFADOS = 8113 'Sem parametros
'Erro de leitura na Tabela Almoxarifados
Public Const ERRO_UNIDADE_NAO_CADASTRADA = 8114 'Sem parametros
'Unidade de medida n�o cadastrada
Public Const ERRO_TOTAL_RESERVADO_SEM_PREENCHIMENTO = 8116 'Sem parametors
'Total Reservado n�o foi informado
Public Const ERRO_NENHUM_ALMOXARIFADO_BD = 8117 'Sem parametros
'Nemhum almoxarifado est� cadastrado no Banco de Dados
Public Const ERRO_ITEM_NAO_EXISTE = 8118 'Parametro sCodItem
'Item com c�digo %s n�o existe
Public Const ERRO_PRODUTO_NAO_DISPONIVEL = 8119 'Parametro objItemPedido.sProduto
'N�o existe disponibilidade do produto %s
Public Const ERRO_QUANTRESERVADA_MAIOR = 8121 'Parametro dQuantReservada
'Quantidade reservada n�o pode ser maior do que a quantidade dispon�vel!
Public Const ERRO_TOTAL_RESERVADO_MAIOR = 8122 'Parametro dTotalReservado
'O Total reservado n�o pode ser maior que o total a reservar!
Public Const ERRO_TABELA_NAO_MARCADA = 8123 'Sem par�metros
'Deve ser selecionada pelo menos uma Tabela para atualiza��o.
Public Const ERRO_TABELAPRECO_INEXISTENTE1 = 8125 'Sem Parametro
'Nenhuma Tabela de Pre�o Cadastrada no Banco de Dados
Public Const ERRO_PERCENTUAL_NAOPREENCHIDO = 8126 'Sem parametro
'O valor do percentual de reajuste deve estar preenchido.
Public Const TELA_AUTCRED_CHAMADA_SEM_PARAMETRO = 8127 'Sem Parametro
'A tela Autoriza��o de Cr�dito foi chamada sem a passagem do parametro necess�rio.
Public Const ERRO_EMPRESA_INVALIDA = 8128
'Utiliza��o Empresa inv�lida
Public Const ERRO_BLOQUEIO_LIBERACAO_NAO_MARCADO = 8130 'Sem par�metro
'Pelo menos um Bloqueio deve estar marcado para Libera��o.
Public Const ERRO_BLOQUEIO_EDICAO_NAO_MARCADO = 8131 'Sem par�metro
'Deve haver um Bloqueio marcado para Edi��o.
Public Const ERRO_TIPOBLOQUEIO_NAO_MARCADO = 8132 'Sem par�metro
'Pelo menos 1 Tipo de Bloqueio deve estar marcado para esta opera��o.
Public Const ERRO_AUSENCIA_ALMOXARIFADO_FILIAL = 8134 'iFilialEmpresa
'Nunhum almoxarifado da Filial %i est� cadastrado no Banco de Dados.
Public Const ERRO_ITEM_INEXISTENTE = 8135
'�tem n�o existe
Public Const ERRO_QUANT_ALOCADA_MAIOR_DISPONIVEL = 8136
'Quantidade Alocada n�o pode ser maior do que a Quantidade Dispon�vel
Public Const ERRO_TOTAL_ALOCACAO_SUPERIOR_ALOCAR = 8137
'Total Alocado n�o pode ser maior do que a Quantidade a Alocar.
Public Const ERRO_TIPODOC_DIFERENTE_NF_SAIDA_DEVOLUCAO = 8159 'iTipoNFiscal
'Tipo de Documento %i n�o � Nota Fiscal de Saida de Devolu��o.
Public Const ERRO_NF_EXTERNA_NAO_CADASTRADA = 8161 'Par�metros: sSerie, lNumNotaFiscal, lFornecedor, iFilial
'Nota Fiscal Externa com s�rie %s, n�mero %l, fornecedor = %l, Cliente = %l e Filial = %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_DATASAIDA_NAO_PREENCHIDA = 8162
'A Data de Sa�da n�o foi preenchida.
Public Const ERRO_LEITURA_NOTA_FISCAL = 8163 'Par�metros: iTipoNFiscal, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscal na Nota Fiscal com Tipo = %i, Serie = %s e N�mero = %l.
Public Const ERRO_ALTERACAO_NFISCAL_SAIDA2 = 8164 'Par�metros: iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal com os dados Tipo = %i, Serie = %s, N�mero NF = %l, Data Emissao = %dt est� cadastrada no Banco de Dados. N�o � poss�vel alterar.
Public Const ERRO_PEDIDO_VENDA_BLOQUEIO_CREDITO = 8165 'Parametro: lCodPedido
'N�o � poss�vel gerar Nota Fiscal a partir do pedido %l pois ele possui um Bloqueio de Cr�dito.
Public Const ERRO_QUANTIDADE_FATURAR_MENOR = 8166 'Parametro: dQuantFaturar
'A quantidade do item n�o pode ultrapassar a quantidade a ser faturada que � %d.
Public Const ERRO_PEDIDO_VENDA_BLOQUEIO_TOTAL = 8167 'Parametro: lCodPedido
'N�o � poss�vel gerar Nota Fiscal a partir do pedido %l pois ele possui um Bloqueio Total.
Public Const ERRO_ALTERACAO_NFISCALPEDIDO = 8168
'N�o � poss�vel alterar nota fiscal gerada a partir de pedido.
Public Const ERRO_CODPEDIDO_NAO_INFORMADO = 8169
'O C�digo do Pedido deve ser informado.
Public Const ERRO_FILIALPEDIDO_NAO_INFORMADA = 8170
'A Filial do Pedido deve ser informada.
Public Const ERRO_PEDIDOVENDA_BAIXADO = 8171 'Par�metros:lCodPedido, iFilialPedido
'O Pedido de Venda com o c�digo %l da Filial Empresa %i j� est� baixado
Public Const ERRO_PEDIDO_VENDA_NAO_CADASTRADO1 = 8172 'Par�metros:lCodPedido, iFilialPedido
'O Pedido de Venda com o c�digo %l da Filial Empresa %i n�o est� cadastrado no Banco de Dados.
Public Const ERRO_FILIALFATURAMENTO_DIFERENTE = 8173 'Parametros: iFilialFaturamento,lCodPedido, iFilialEmpresa
'A Filial de Faturamento %i do Pedido de Venda %l, � diferente da Filial empresa Atual %i.
Public Const ERRO_NENHUM_PEDIDO_TRAZIDO = 8174
'Nenhum pedido de venda foi trazido para a tela.
Public Const ERRO_PEDIDO_VENDA_FATURA_INTEGRAL = 8175
'O Pedido de Venda deve ser faturado integralmente.
Public Const ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO = 8176
'N�o pode haver substituicao de produto.
Public Const ERRO_CANALVENDA_NAO_ENCONTRADO = 8177 '%sCanalVenda
'O Canal de Venda %s n�o foi encontrado.
Public Const ERRO_TIPODOC_DIFERENTE_NF_SAIDA_REMESSA = 8178 'Par�metro: iTipoNFiscal
'Tipo de Documento %i n�o � Nota Fiscal de Saida de Remessa.
Public Const ERRO_MODIFICACAO_SERIE = 8179 'Sem parametro
'Erro na modifica��o da Serie.
Public Const ERRO_INSERCAO_SERIE = 8180 'Sem parametro
'Erro na inser��o da Serie.
Public Const ERRO_EXCLUSAO_SERIE = 8181 'Parametro: sSerie
'Erro na tentativa de excluir Serie.
Public Const ERRO_PROXNUMNFISCAL_NAO_PREENCHIDA = 8182
'O Proximo N�mero para Nota Fiscal deve ser preenchido.
Public Const ERRO_SERIE_REL_NF_REC = 8183 'Parametro: sSerie
'Erro na exclus�o da Serie %s, relacionado com Nota Fiscal a Receber.
Public Const ERRO_SERIE_REL_NF_REC_BAIXADA = 8184 'Parametro: sSerie
'Erro na exclus�o da Serie %s, relacionado com Nota Fiscal a Receber Baixada.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_SERIE = 8185
'Erro na leitura da tabela de S�ries de Notas Fiscais.
Public Const ERRO_EXISTEM_NOTAS_FISCAIS_SERIE = 8186
'A s�rie n�o pode ser exclu�da, pois est� vinculada com a Nota Fiscal %l da Filial Empresa %i.
Public Const ERRO_NAO_EXISTE_RESERVAS1 = 8187 ' Parametros lCodNFiscal, sCodProduto
'N�o existe Reservas para a NFiscal  %l do Produto %s.
Public Const ERRO_PRODUTO_SEM_SUBSTITUTOS = 8189 'Sem parametros
'Produto n�o tem Produtos Substitutos associados
Public Const ERRO_PRODUTOS_SUBSTITUTOS_INVALIDOS1 = 8190 'Sem parametros
'Produtos substitutos j� fazem parte da Nota Fiscal ou est�o Inativos ou n�o participam do Faturamento
Public Const ERRO_PEDIDO_EDICAO_NAO_MARCADA = 8195 'Sem par�metros
'Deve haver uma Pedido marcado para Edi��o.
Public Const ERRO_DATA_FATURA_NAO_PREENCHIDA = 8197 'Sem parametros
'A data da fatura n�o foi preenchida.
Public Const ERRO_FATURA_SEM_NOTASFISCAIS = 8198 'S�m Par�metros
'N�o existem Parcelas no Grid de Parcelas Para faturar.
Public Const ERRO_NOTASFISCAIS_NAO_SELECIONADAS = 8199 'Sem Par�metros
'N�o existe nenhuma Nota Fiscal selecionada no Grid.
Public Const ERRO_TIPODEBLOQUEIO_EXCLUSAO = 8200 'Par�metro: iCodigo
'O Tipo De Bloqueio %i n�o pode ser exclu�do.
Public Const ERRO_PEDIDO_FATURAMENTO_PARCIAL = 8208 'Parametro CodigoPV
'O pedido %l n�o pode ser faturado totalmente. Use a transa��o de faturamento de nota fiscal a partir de um pedido para faturar este pedido.
Public Const ERRO_PEDIDO_JA_FATURADO_PARCIALMENTE = 8209 'Parametro CodigoPV
'O pedido %l j� foi faturado parcialmente. Use a transa��o de faturamento de nota fiscal a partir de um pedido para faturar este pedido.
Public Const ERRO_PEDIDO_SEM_COBRANCA = 8210 'Parametro: lCodPedido
'Foi solicitada a gera��o de uma nota fiscal fatura para o Pedido %l que n�o possui informa��es de cobran�a.
Public Const ERRO_INSERCAO_VALORLIBERADOCREDITO = 8211 'Parametros: sCodUsuario, iAno
'Ocorreu um erro na inser��o de um registro na tabela ValorLiberadoCredito. Usu�rio = %s, Ano = %i.
Public Const ERRO_ATUALIZACAO_VALORLIBERADOCREDITO = 8212 'Parametros: sCodUsuario, iAno
'Ocorreu um erro na atualiza��o de um registro da tabela ValorLiberadoCredito. Usu�rio = %s, Ano = %i.
Public Const ERRO_CLIENTE_SEM_CREDITO = 8213 'Parametros: sCodUsuario
'O cliente %s n�o possui cr�dito para efetuar esta opera��o.
Public Const ERRO_LEITURA_NFISCALBAIXADA3 = 8214 'Par�metros: iTipoNFiscal, lCliente, iFilialCli, sSerie, lNumNotaFiscal
'Ocorreu um erro na leitura da tabela NFiscalBaixadas na Nota Fiscal com Tipo = %i, Cliente = %l, Filial = %i, Serie = %s e N�mero = %l.
Public Const ERRO_EXCLUSAO_COMISSOESPEDVENDAS1 = 8215 'Parametro: iFilialEmpresa, lPedidoVenda
'Erro na tentativa de excluir registro da tabela de ComissoesPedVendas da Filial %i e Pedido %l.
Public Const ERRO_PRODUTOS_SUBSTITUTOS_INVALIDOS = 8218 'Sem parametros
'Produtos substitutos j� fazem parte do pedido de venda ou est�o Inativos ou n�o participam do Faturamento
Public Const ERRO_VARIAS_NFISCAL_EDICAO = 8219 'Sem par�metro
'S� deve haver uma Nota Fiscal marcada para Edi��o.
Public Const ERRO_FILIALDE_MAIOR_FILIALATE = 8220 'Sem par�metros
'Cliente De n�o pode ser maior do que o Cliente At�.
Public Const ERRO_TIPO_BLOQUEIO_PRE_DEFINIDO = 8221 'sem parametros
'Um tipo de bloqueio pr�-definido n�o pode ser alterado
Public Const ERRO_PEDIDO_VENDA_INICIAL_MAIOR = 8225
'O Pedido de Venda Inicial � maior que o Final.
Public Const ERRO_VENDEDOR_NAO_CADASTRADO2 = 8226
'O Vendedor n�o est� cadastrado.'
Public Const ERRO_TRANSPORTADORA_INICIAL_MAIOR = 8227
'A Transportadora Inicial � maior do que a Final.'
Public Const ERRO_PERIODO_PREENCHIDO_INCORRETO = 8228
'Os Periodos devem estar preenchidos em ordem, sem espa�os entre eles.
Public Const ERRO_PERIODO_ANTERIOR_MENOR = 8229
'O Periodo anterior n�o pode ser maior que o seguinte.
Public Const ERRO_INSERCAO_RELFATPRAZOPAG = 8230 'Par�metro: lCodigo
'Erro na Inser��o do Relatorio com o C�digo %l
Public Const ERRO_NUM_MAXIMO_BLOQUEIO_MAIOR_LIMITE = 8231 'Sem par�metros
'O n�mero de Bloqueios ultrapassa ao limite m�ximo que � de 1000.
Public Const ERRO_NATUREZAOP_SAIDA = 8232 'parametro: natop
'Informe uma natureza de opera��o de sa�da (>= 500).
Public Const ERRO_NATUREZAOP_ENTRADA = 8233 'parametro: natop
'Informe uma natureza de operacao de entrada (<500)
Public Const ERRO_NATUREZAOP_ITEM_TRIBUTACAO_NAO_PREENCHIDA = 8234 'iItem
'A Natureza de Opera��o da tributa��o do item %i n�o foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_ITEM_NAO_PREENCHIDO = 8235 'iItem
'O Tipo de Tributacao do item %i n�o foi preenchido.
Public Const ERRO_NATUREZAOP_DESCONTO_NAO_PRENCHIDA = 8236
'A Natureza de Opera��o da tributa��o do Valor Desconto n�o foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_DESCONTO_NAO_PREENCHIDO = 8237
'O Tipo de Tributa��o do Valor Desconto n�o foi preenchido.
Public Const ERRO_NATUREZAOP_FRETE_NAO_PRENCHIDA = 8238
'A Natureza de Opera��o da tributa��o do Valor Frete n�o foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_FRETE_NAO_PREENCHIDO = 8239
'O Tipo de Tributa��o do Valor Frete n�o foi preenchido.
Public Const ERRO_NATUREZAOP_DESPESAS_NAO_PRENCHIDA = 8240
'A Natureza de Opera��o da tributa��o do Valor Despesas n�o foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_DESPESAS_NAO_PREENCHIDO = 8241
'O Tipo de Tributa��o do Valor Despesas n�o foi preenchido.
Public Const ERRO_NATUREZAOP_SEGURO_NAO_PRENCHIDA = 8242
'A Natureza de Opera��o da tributa��o do Valor Seguro n�o foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_SEGURO_NAO_PREENCHIDO = 8243
'O Tipo de Tributa��o do Valor Seguro n�o foi preenchido.
Public Const ERRO_VARIOS_PEDIDO_EDICAO = 8244 'Sem par�metro
'S� deve haver um Pedido marcado para Edi��o.
Public Const ERRO_NOTAFISCAL_NAO_EDITAVEL = 8245 'Sem par�metros
'A Nota Fiscal n�o pode ser edit�vel.
Public Const ERRO_PEDIDOVENDA_NAO_INFORMADO = 8246 'Sem par�metros
'O n�mero do Pedido de Venda deve ser preenchido.
Public Const ERRO_NFISCAL_NAO_INFORMADA = 8247 'Sem par�metros
'O n�mero e a s�rie da Nota Fiscal devem estar preenchidos.
Public Const ERRO_LEITURA_NFISCAL5 = 8248 'Par�metros: iFilialEmpresa, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscal com Filial %i, S�rie %s e N�mero da Nota %l.
Public Const ERRO_SERIE_MAIOR_LIMITE_MAXIMO = 8249 'Sem par�metros
'Ultrapassou o limite m�ximo do tamanho da S�rie.
Public Const ERRO_VALORBASE_MAIOR_VALORDOC = 8250 'Par�metros: dValorBase, dValorDoc
'O Valor Base n�o pode ser maior que o Valor do Documento.
Public Const ERRO_DATA_DESCONTO_PARCELA_SUPERIOR_DATA_VENCIMENTO = 8251 'Par�metros: dtDataDesconto, iDesconto, iParcela
'A data %dt do desconto %i da parcela %i � superior a data de vencimento.
Public Const ERRO_DATAINICIAL_NAOPREENCHIDA = 8252 'Sem Par�metros
'Erro a Data Inicial n�o foi preenchida.
Public Const ERRO_DATAFINAL_NAOPREENCHIDA = 8253 'Sem Par�metros
'Erro a Data Final n�o foi preenchida.
Public Const ERRO_FILIALENTREGA_NAO_INFORMADA = 8255
'A Filial de Entrega do Cliente n�o foi preenchida.
Public Const ERRO_NATUREZAOP_INICIAL_MAIOR = 8256
'A Natureza de Opera��o Inicial � maior que a Final.
Public Const ERRO_PEDIDO_VENDA_BLOQUEIO = 8257 'Par�metro: lCodPedido
'N�o � poss�vel gerar Nota Fiscal a partir do pedido %l pois ele possui bloqueios que impedem o seu faturamento.
Public Const ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2 = 8258 'Parametro: sProduto
'O Produto %s n�o pode ser vendido.
Public Const ERRO_LOCK_COMISSOESNF = 8259 'Sem Parametros
'Ocorreu um erro na tentativa de fazer um lock de um registro da tabela de Comiss�es de Notas Fiscais.
Public Const ERRO_EXCLUSAO_COMISSOESNF2 = 8260 'Sem Parametros
'Erro na tentativa de excluir registro da tabela de Comiss�es de Notas Fiscais.
Public Const ERRO_EXCLUSAO_RESERVA1 = 8261 'Sem par�metros
'Erro na tentativa de excluir registro da tabela de Reservas.
Public Const ERRO_ATUALIZACAO_RESERVA1 = 8262 'Sem Par�metros
'Erro na atualiza��o de resgitsro na tabela de reservas.
Public Const ERRO_LEITURA_USUARIOFILIALEMPRESA_DIC = 8263 'Sem Parametros
'Ocorreu um erro na leitura da Query UsuarioFilialEmpresa do Dicionario de Dados.
Public Const ERRO_LEITURA_PEDIDOS_VENDA_BAIXA_PV = 8264 'Sem parametros
'Erro na leitura de Pedidos de Venda para baixa.
Public Const ERRO_NOTA_FISCAL_NAO_TEM_COMISSAO = 8265 'Parametros: lNumeroNFiscal, sSerie
'A Nota Fiscal %l da S�rie %s n�o pode ter comiss�o ou n�o est� cadastrada.
Public Const ERRO_NENHUM_PEDIDO_SELECIONADO = 8267
'Para editar um Pedido � necess�rio que uma das linhas do grid seja selecionada.
Public Const ERRO_FILIALPEDIDO_DIFERENTE_FILIALEMPRESA = 8268 'Par�metros: lCodPedido, iFilialEmpresaPedido, giFilialEmpresa
'N�o � poss�vel editar o Pedido de Venda %l da Filial Empresa %i pois estamos na Filial Empresa %i.
Public Const ERRO_PRODUTO_SEM_ALMOX_PADRAO1 = 8269 'sProduto, giFilialEmpresa
'O Produto %s n�o est� relacionado com nenhum Almoxarifado da Filial Empresa %i.
Public Const ERRO_QUANTIDADE_PREVVENDA_NAO_PREENCHIDA = 8270 'Sem Parametro
'A Quantidade da previs�o de Venda n�o foi preenchida.
Public Const ERRO_CODIGO_DESABILITADO_IMUTAVEL = 8271
'N�o � poss�vel alterar c�digo do Pedido de Vendas.
Public Const ERRO_FILIALENTREGA_FORNECEDOR_NAO_INFORMADA = 8272
'A Filial de Entrega do Fornecedor n�o foi preenchida.
Public Const ERRO_CONDICAO_PAGTO_ALTERADA_NUM_PARC = 8273 'parametro: codigo da condicao
'A condi��o de pagamento %s teve alterado o seu n�mero de parcelas.
Public Const ERRO_LEITURA_NFISCALTIPODOCINFO = 8274 'Sem par�metros
'Erro na leitura da tabela NfiscalTipoDocInfo.
Public Const ERRO_QUANT_FATURADA_MAIOR_QUANT_A_FATURAR = 8275 'parametro lNumIntItemPV
'Quantidade faturada � maior do que quantidade a faturar de �temPV com n�mero interno %l.
Public Const ERRO_TIPODOC_DIFERENTE_NF_FATURA_PEDIDO = 8276 'Par�metro: iTipoDocInfo
'Tipo de Documento %i n�o � Nota Fiscal Fatura Pedido.
Public Const ERRO_TIPODOC_DIFERENTE_NF_PEDIDO = 8277 'Par�metro: iTipoDocInfo
'Tipo de Documento %i n�o � Nota Fiscal Pedido.
Public Const ERRO_PRODUTO_SEM_TABELA_PADRAO = 8278 'Parametro: sCodProduto
'Produto com c�digo %s n�o tem Tabela de Pre�o Padr�o associada.
Public Const ERRO_TABELAPRECOITEM_INEXISTENTE3 = 8279 'Parametros: sCodProduto, dtDataFinal
'Inexiste no Banco de Dados Item de Tabela de Pre�o Padr�o para Produto %s com Data de Vig�ncia anterior ou igual a %dt.
Public Const ERRO_VALOR_EMISSAO_GRID_NAO_PREENCHIDO = 8280 'Parametro: iLinha
'O valor da comiss�o na emiss�o na linha %i do Grid de Comiss�es n�o foi preenchido.
Public Const ERRO_PERCENTAGEM_EMISSAO_GRID_NAO_PREENCHIDA = 8281
'A porcentagem da comiss�o na emiss�o na linha %i do Grid de Comiss�es n�o foi preenchida.
Public Const ERRO_PERCENTUAL_COMISSAO_NULO = 8282
'A porcentagem de comiss�o n�o pode ser nula.
Public Const ERRO_LIBERACAOCREDITO_LIMITEOPERACAO1 = 8283 'Parametros: lPedido, sCodUsuario
'O Pedido %l n�o teve o cr�dito liberado pois o usu�rio %s ultrapassou o limite por opera��o.
Public Const ERRO_LIBERACAOCREDITO_INEXISTENTE1 = 8284 'Parametros: sCodUsuario, lPedido
'O usu�rio %s n�o tem autoriza��o para liberar por cr�dito o bloqueio do Pedido %l.
Public Const ERRO_LIBERACAOCREDITO_LIMITEMENSAL1 = 8285 'Parametros: lPedido, sCodUsuario
'O Pedido %l n�o teve o cr�dito liberado, pois o usu�rio %s ultrapassou o limite mensal.
Public Const ERRO_LEITURA_ITEMPV = 8286 'Parametro: lNumIntDoc
'Erro na leitura da tabela ItensPedidoDeVenda, registro com n�mero interno %l.
Public Const ERRO_LOCK_ITEMPEDIDODEVENDA = 8287 'Parametro: lNumIntDoc
'Erro de lock na tabela ItensPedidoDeVenda, registro com n�mero interno %l.
Public Const ERRO_QUANT_LIBERADA_MAIOR_QUANT_RESERVADA = 8288 'Parametro: lNumIntDoc
'Quantidade de reserva liberada de ItemPV com n�mero interno %l superior a quantidade reservada do �tem.
Public Const ERRO_ATUALIZACAO_ITEMPV = 8289 'Parametro: lNumIntDoc
'Erro na atualiza��o de ItemPedidoDeVenda com n�mero interno %l.
Public Const ERRO_NUMERO_DE_NAO_PREENCHIDO = 8291 'Sem Parametros
'O n�mero de Nota Fiscal De n�o foi preenchido.
Public Const ERRO_NUMERO_ATE_NAO_PREENCHIDO = 8292 'Sem Parametros
'O n�mero de Nota Fiscal At� n�o foi preenchido.
Public Const ERRO_LOCK_SERIE_IMPRESSAO_NF = 8293 'Parametros : sS�rie
'A S�rie %s est� lockada para impress�o.
Public Const ERRO_LEITURA_TRIBUTACAOPVBAIXADO = 8294
'Erro na leitura da tabela de TributacaoPVBaixado.
Public Const ERRO_COMPL_PV_BAIXADO_TIPO_INVALIDO = 8295 'sem parametros
'registro na tabela TributacaoComplPVBaixado com tipo inv�lido
Public Const ERRO_SERIE_SEM_PADRAO = 8298 'Parametro: sSerie
'Para desmarcar a Serie %s como Padr�o marque outra S�rie como Padr�o.
Public Const ERRO_TIPO_NOTA_FISCAL_NAO_FATURA_PEDIDO = 8299 'Parametro: iCodigo
'C�digo %i n�o corresponde a Nota Fiscal Fatura de Venda a partir de Pedido.
Public Const ERRO_NFISCALFATURA_SEM_TITULO_RECEBER = 8300 'Parametro: lNumNotaFiscal
'Nota Fiscal Fatura com n�mero %l n�o tem T�tulo a Receber associado.
Public Const ERRO_TITULO_RECEBER_NAO_CADASTRADO = 8301 'Parametro: lNumIntDoc
'T�tulo Receber com n�mero interno %l n�o est� cadastrado.
Public Const ERRO_FATURAPAG_NAO_EXCLUIDA = 8302
'A nota fiscal n�o pode ser cancelada/exclu�da por estar vinculada a uma fatura a pagar.
'Para realizar o cancelamento da nota fiscal � preciso, antes, excluir a Fatura a Pagar.
Public Const ERRO_FATURAREC_NAO_EXCLUIDA = 8303 'Par�metro : lNumFatura
'A nota fiscal n�o pode ser cancelada/exclu�da por estar vinculada a fatura a Receber %l.
'Para realizar o cancelamento da nota fiscal � preciso, antes, excluir a Fatura a Receber.
Public Const ERRO_COMISSAO_BAIXADA_CANC_NFISCAL = 8304 'sem parametros
'N�o pode cancelar uma nota fiscal que teve a comissao j� baixada(paga).
Public Const ERRO_EXCLUSAO_FILIALCLIENTEFILEMP = 8305 'Sem Parametros
'Erro na Exclus�o da tabela FilialClienteFilEmp.
Public Const ERRO_FATURADE_NAO_PREENCHIDA = 8306 'Sem Parametros
'� obrigat�rio o preenchimento do campo Fatura De.
Public Const ERRO_FATURAATE_NAO_PREENCHIDA = 8307 'Sem Parametros
'� obrigat�rio o preenchimento do campo Fatura At�.
Public Const ERRO_TIPOPRODUTO_INICIAL_MAIOR = 8308 'Sem par�metros
'O tipo de produto inicial n�o pode ser maior que o tipo de produto final.
Public Const ERRO_INCLUSAO_NFISCAL_NUMAUTO = 8309 'sem parametros
'Para criar uma nota fiscal o numero da mesma tem que estar em branco
Public Const ERRO_NF_NAO_CADASTRADA2 = 8310 'Par�metros: lNumNotaFiscal, sSerie, iTipoNFiscal, lCliente, iFilialCli, dtDataEmissao
'A nota fiscal com os dados abaixo n�o est� cadastrada. N�mero: %s, S�rie: %s, Tipo: %s, Cliente: %s, Filial: %s e Emiss�o em: %s.
Public Const ERRO_NF_NAO_CADASTRADA4 = 8311 'Par�metros: lNumNotaFiscal, sSerie, iTipoNFiscal, lFornecedor, iFilialForn, dtDataEmissao
'A nota fiscal com os dados abaixo n�o est� cadastrada. N�mero: %s, S�rie: %s, Tipo: %s, Fornecedor: %s, Filial: %s e Emiss�o em: %s.
Public Const ERRO_VALORICMS_MAIOR_TOTAL = 8312 'Parametros: dValorICMS , dValorTotal
'O Valor do Imposto %s n�o pode ser maior que o Valor Total  %s.
Public Const ERRO_DESTINATARIO_NAO_PREENCHIDO = 8313
' O Preenchimento do Destinat�rio � obrigat�rio
Public Const ERRO_BASECALCULO_NAO_PREENCHIDA = 8314
'O Preenchimento da base de c�lculo � obrigat�rio.
Public Const ERRO_VALORFRETE_NAO_PREENCHIDO = 8315
'O Preenchimento do Valor do Frete � obrigat�rio.
Public Const ERRO_LEITURA_CONHECIMENTO_FRETE = 8316
'Erro de leitura na tabela de Conhecimento de Frete.
Public Const ERRO_CONHECIMENTOFRETE_NAO_CADASTRADO = 8317 'Parametros: sSerie, lNumNotaFiscal
'O Conhecimento de Transporte com a s�rie %s e N�mero %s n�o est� cadastrado.
Public Const ERRO_REMETENTE_NAO_PREENCHIDO = 8318
'O Preenchimento do Remetente � obrigat�rio.
Public Const ERRO_VALORBASE_MENOR_SUBTOTAL = 8319 'Parametros: dValorBase , dValorTotal
'O Valor Base %s n�o pode ser menos que a soma dos valores que � %s.
Public Const ERRO_INCLUSAO_NFISCAL = 8320 'Par�metros: sSerie, lNumNotaFuscal
'Erro na tentativa de incluir a Nota Fiscal  Serie %s N�mero %s.
Public Const ERRO_INCLUSAO_CONHECIMENTOFRETE = 8321 'Par�metros: sSerie, lNumNotaFuscal
'Erro na tentativa de incluir a Nota Fiscal  Serie %s N�mero %s.




'C�digos de Avisos - Reservado de 5900 at� 5999
Public Const AVISO_EXCLUSAO_NATUREZAOP = 5900 'Parametro: sCodigo
'Confirma a exclus�o da Natureza de Opera��o %s ?
Public Const AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS = 5901 'Sem Parametros
'N�o � poss�vel selecionar "Todos" para produtos e clientes simultaneamente.
Public Const AVISO_CONFIRMA_EXCLUSAO_CANAL = 5902 'Parametro objCanal.iCodigo
'Confirma exclus�o do Canal %i ?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPODEBLOQUEIO = 5903 'Parametro objTipo.iCodigo
'Confirma exclus�o do Tipo de Bloqueio %i ?
Public Const AVISO_EXCLUSAO_PREVVENDA = 5904 'Parametro sCodigo
'Confirma exclus�o da Previs�o %s do Banco de Dados ?
Public Const AVISO_NUM_MAX_BLOQUEIOS_LIBERACAO = 5905
'O n�mero m�ximo de Bloqueios poss�veis de exibi��o foi atingido . Ainda existem mais Bloqueios para Libera��o .
Public Const AVISO_CONFIRMA_EXCLUSAO_ALCADAFAT_USUARIO = 5906 'Par�metro: C�digo do Usu�rio
'Confirma a exclus�o da al�ada do usu�rio com c�digo %s?
Public Const AVISO_ITEM_ANTERIOR_ALTERADO = 5907 'Sem parametros
'As reservas do Item anterior foram alteradas!Deseja salvar altera��es?
Public Const AVISO_CANCELAR_NFISCAL = 5908 'Par�metro: lNumNotaFiscal
'Deseja realmente cancelar a Nota Fiscal de Sa�da %l ?.
Public Const AVISO_ALOCADO_MENOR_ALOCAR = 5909 'dTotalAlocado, dQuantAlocar
'A quantidade alocada %d � inferior � quantidade a alocar %d. Deseja prosseguir?
Public Const AVISO_NFISCAL_SAIDA_DEVOLUCAO_MESMO_NUMERO = 5910 'Par�metros: sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados exite Nota Fiscal  com os Dados S�rie NF =%s, N�mero NF =%l, Data Emiss�o =%dt.
'Deseja prosseguir na inser��o de Nota Fiscal com o mesmo n�mero?
Public Const AVISO_PEDIDO_TELA_NAO_UTILIZADO = 5911 'Parametros: lCodPedido, iFilialPedido
'A Nota Fiscal da tela vai ser gerada com base nos dados do Pedido %l da Filial %i trazidos anteriormente para a tela.
'O codigo e filial que est�o na tela ser�o ignorados. Deseja prosseguir na grava��o da nota fiscal?
Public Const AVISO_CRIAR_CANALVENDA = 5912 'iCanalVenda
'O Canal Venda %i n�o existe. Deseja Cri�-lo?
Public Const AVISO_NFISCAL_SAIDA_REMESSA_MESMO_NUMERO = 5913 'PAR�METROS: sSerie, lNumero, dtDataEmissao
'No Banco de Dados exite Nota Fiscal  com os Dados S�rie NF =%s, N�mero NF =%l, Data Emiss�o =%dt.
'Deseja prosseguir na inser��o de Nota Fiscal com o mesmo n�mero?
Public Const AVISO_EXCLUIR_SERIE = 5914 'Parametro: sSerie
'Confirma exclus�o da S�rie %s?
Public Const AVISO_EXISTE_NF_MAIOR_OU_IGUAL = 5915 ' Parametros: sSerie, lNumNotaFiscal, dtDataEmissao, iFilialEmpresa
'Existe Nota Fiscal com N�mero maior ou igual � Nota: S�rie = %s, N�mero = %l, Data de Emiss�o %dt e FilialEmpresa = %i. Confirma a Grava��o ?
Public Const AVISO_NF_ULTIMA_GRAVADA = 5916 ' Parametros: sSerie, lNumNotaFiscal, dtDataEmissao, iFilialEmpresa
'A �ltima Nota Fiscal gravada foi: S�rie = %s, N�mero = %l, Data de Emissao %dt e FilialEmpresa = %i.
Public Const AVISO_ALTERACAO_NFISCAL_SAIDA_CONTAB = 5918 'Par�metros: iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal com os dados Tipo = %i, Serie = %s, N�mero NF = %l, Data Emissao = %dt est� cadastrada no Banco de Dados, s� poder� ser alterado os dados relativos a contabilidade. Deseja proseguir na altera��o?
Public Const AVISO_NFISCAL_REIMPRESSA = 5919 'Parametros : lNumeroNFInicial, lNumNotaFinal
'Confirma a reimpress�o das Notas de n�mero %l at� %l ?
Public Const AVISO_NFISCAL_LOCKADA = 5920 'Sem Parametros
'A Impress�o da Nota Fiscal est� bloqueada. Est� havendo uma impress�o ou houve erro anterior.
'Continue somente em caso de Erro anterior. Deseja Continuar?
Public Const AVISO_SERIE_GRAVADA_PADRAO = 5921 'Parametro: sSerie
'A S�rie %s ser� a S�rie padr�o.



