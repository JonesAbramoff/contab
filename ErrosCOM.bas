Attribute VB_Name = "ErrosCOM"
Option Explicit

'C�digos de Erros - Reservado de 12000 a 12999
Public Const ERRO_PERCENT_MAIS_QUANTCOTACAO_ANTERIOR_NAO_PREENCHIDA = 12000
'O campo Percentagem a mais de Cota��es Anteriores n�o foi preenchido.
Public Const ERRO_PERCENT_MENOS_QUANTCOTACAO_ANTERIOR_NAO_PREENCHIDA = 12001
'O campo Percentagem a menos de Cota��es Anteriores n�o foi preenchido.
Public Const ERRO_PERCENT_MAIS_RECEB_NAO_PREENCHIDA = 12002
'O campo Percentagem a mais da Faixa de Recebimento n�o foi preenchido.
Public Const ERRO_PERCENT_MENOS_RECEB_NAO_PREENCHIDA = 12003
'O campo Percentagem a menos da Faixa de Recebimento n�o foi preenchido.
Public Const ERRO_LEITURA_ALCADA = 12009 'Par�metro: sCodUsuario
'Erro na busca da al�ada do usu�rio %s na tabela de Al�adas.
Public Const ERRO_LOCK_ALCADA = 12010 'Par�metro: sCodUsuario
'N�o conseguiu fazer o lock da al�ada do usu�rio %s.
Public Const ERRO_LEITURA_PEDIDOCOMPRABAIXADO = 12012 'Sem parametro
'Erro na leitura da tabela de Pedido de Compra Baixado
Public Const ERRO_LEITURA_VALORPCLIBERADO = 12013 'Sem parametro
'Erro na leitura de Valor de Pedido de Compra Liberado
Public Const ERRO_ALCADA_VINCULADA_PEDIDOCOMPRA = 12014 'Parametro sCodUsuario
'A al�ada n�o pode ser exclu�da pois o usu�rio %s est� vinculado a um Pedido de Compra
Public Const ERRO_ALCADA_VINCULADA_PEDIDOCOMPRABAIXADO = 12015 'Parametro sCodUsuario
'A al�ada n�o pode ser exclu�da pois o usu�rio %s est� vinculado a um Pedido de Compra Baixado
Public Const ERRO_ALCADA_VINCULADA_VALORPCLIBERADO = 12016 'Parametros:sCodUsuario, iAno
'A al�ada n�o pode ser exclu�da pois o usu�rio %s est� vinculado a um registro
'na tabela de Valor de Pedido de Compra Liberado
Public Const ERRO_EXCLUSAO_ALCADA = 12017 'Par�metro: sCodUsuario
'Erro na tentativa de exclus�o da al�ada do usu�rio %s da tabela de Al�adas.
Public Const ERRO_ATUALIZACAO_ALCADA = 12018 'Par�metro: sCodUsuario
'Erro na tentativa de atualizar a al�ada do usu�rio %s da tabela de Al�adas.
Public Const ERRO_INSERCAO_ALCADA = 12019 'Parametro: sCodUsuario
'Erro na tentiva de inserir al�ada do usuario %s.
Public Const ERRO_LEITURA_REQUISITANTE = 12021 'Parametro:  lCodigo
'Erro na leitura do Requisitante %l na tabela de Requisitantes.
Public Const ERRO_LEITURA_REQUISITANTE1 = 12022 'Par�metro: sNomeReduzido
'Erro na leitura do Requisitante %s na tabela de Requisitantes.
Public Const ERRO_LEITURA_REQUISICAOCOMPRAS = 12023 'Sem Par�metros
'Erro na leitura da tabela de Requisi��o de Compras.
Public Const ERRO_LEITURA_ITENSCONCORRENCIA = 12025 'Par�metro: lCodConcorrencia
'Erro na leitura dos itens de concorr�ncia da concorr�ncia com o c�digo %l da tabela de itens de concorr�ncia.
Public Const ERRO_REQCOMPRA_VINCULADA_CONCORRENCIA_NAO_ENCONTRADA = 12027
'Uma requisi��o de compras vinculada a concorr�ncia n�o foi encontrada.
Public Const ERRO_USUARIO_SEM_ALCADA = 12028 'parametro sCodUsuario
'O usu�rio %s n�o possui al�ada.
Public Const ERRO_LOCK_VALORPCLIBERADO = 12029 ' Parametro sCodUsuario
'Erro na tentativa de fazer lock no Valor de Pedido de Compra do usu�rio %s.
Public Const ERRO_INSERCAO_VALORPCLIBERADO = 12030 'parametro sCodUsuario
'Erro na tentativa de inserir Valor de Pedido de Compra Liberado do usu�rio %s
Public Const ERRO_ATUALIZACAO_VALORPCLIBERADO = 12031 'Parametro sCodUsuario
'Erro na atualiza��o do Valor de Pedido de Compra Liberado do usu�rio %s
Public Const ERRO_LIMITE_MENSAL_MENOR = 12032 'parametro dLimiteMensal
'O limite mensal %d � menor que o valor do pedido.
Public Const ERRO_LIMITE_OPERACAO_MENOR = 12033 'parametro dLimiteOperacao
'O limite de opera��o %d � menor que  o valor do pedido.
Public Const ERRO_AUSENCIA_BLOQUEIOS_LIBERAR = 12034 'Sem Parametros
'N�o h� bloqueios no Grid selecionados para a libera��o
Public Const ERRO_NUM_BLOQUEIOS_SELECIONADOS_SUPERIOR_MAXIMO = 12035  'Sem parametros
'O n�mero de bloqueios selecionados � superior ao n�mero m�ximo poss�vel para libera��o. Restrinja mais a sele��o para que o n�mero de bloqueios lidos diminua.
Public Const ERRO_SEM_BLOQUEIOS_PC_SEL = 12036 'Sem parametros
'N�o h� bloqueios dentro dos crit�rios de sele��o informados.
Public Const ERRO_LOCK_BLOQUEIOSPC = 12037 'Parametro: lPedCompras
'N�o conseguiu fazer lock no Bloqueio de Pedido de Compras %l.
Public Const ERRO_ATUALIZACAO_BLOQUEIOSPC = 12038 'Parametro lPedCompras
'Erro na atualiza��o do Bloqueio de Pedido de Compras %l.
Public Const ERRO_LEITURA_BLOQUEIOSPC1 = 12039 'Parametro: lPedCompras
'Erro na leitura do Bloqueio de Pedido de Compras %l.
Public Const ERRO_TIPOBLOQUEIOPC_NAO_MARCADO = 12040 'Sem parametros
'Pelo menos 1 Tipo de Bloqueio deve estar marcado para esta opera��o.
Public Const ERRO_ATUALIZACAO_PEDIDOCOMPRA = 12041 'Parametro lPedCompra
'Erro na atualiza��o do Pedido de Compra com o c�digo %l.
Public Const ERRO_BLOQUEIOPC_INEXISTENTE = 12042 'Parametro lPedCompra
'O Bloqueio de Pedido de Compras %l n�o existe.
Public Const ERRO_LOCK_PEDIDOCOMPRA = 12043 'Parametro lPedCompra
'N�o conseguiu fazer o lock do Pedido de Compra com o c�digo %l.
Public Const ERRO_LIMITE_MENSAL_ULTRAPASSADO = 12044 'Parametro dLimiteMensal, sCodUsuario
'O limite mensal %d do usu�rio %s foi ultrapassado.
Public Const ERRO_COMPRADOR_USUARIO = 12045 'Parametros: iCodComprador,sCodUsuario
'O c�digo de Comprador %i correspondente ao Usu�rio %s no Bando de Dados n�o confere com o Comprador %i da Tela.
Public Const ERRO_COMPRADOR_NAO_CADASTRADO = 12046  'Parametro: iCodigo
'O comprador com o codigo %i nao esta cadastrado
Public Const USUARIO_COMPRADOR_NAO_ALTERAVEL = 12047 'Parametro: iCodigo
'O comprador com o codigo %i nao esta cadastrado
Public Const ERRO_LEITURA_COMPRADOR2 = 12048
'Erro de leitura na tabela de compradores.
Public Const ERRO_LEITURA_COMPRADOR = 12049 'Parametro: iCodigo
'Erro de leitura do Comprador com o c�digo %i na tabela de compradores.
Public Const ERRO_LEITURA_COMPRADOR1 = 12050 'Parametro: sCodUsuario
'Erro de leitura do comprador com o c�digo de usu�rio %s.
Public Const ERRO_LOCK_COMPRADOR = 12051
'Erro na tentativa de fazer 'lock' na tabela Comprador
Public Const ERRO_ATUALIZACAO_COMPRADOR = 12052 'Parametro: iCodComprador
'Erro na tentativa de atualizar o comprador %i a tabela Comprador.
Public Const ERRO_INSERCAO_COMPRADOR = 12053 'Parametro: iCodComprador
'Erro na tentiva de inserir o Comprador %i na tabela Comprador.
Public Const ERRO_EXCLUSAO_COMPRADOR = 12054 ' Parametro: iCodComprador
'Erro na tentativa de excluir o comprador %i na tabela Comprador
Public Const ERRO_COMPRADOR_VINCULADO_CONCORRENCIA = 12055 'Parametro: icodigo
'Comprador %i nao pode ser excluido pois esta vinculado a um registro na tabela Concorrencia
Public Const ERRO_COMPRADOR_VINCULADO_CONCORRENCIABAIXADA = 12056 'Parametro: icodigo
'Comprador %i nao pode ser excluido pois esta vinculado a um registro na tabela Concorrenciabaixada
Public Const ERRO_COMPRADOR_VINCULADO_PEDIDOCOMPRA = 12057 'Parametro: icodigo
'Comprador %i nao pode ser excluido pois esta vinculado a um registro na tabela PedidodeCompra
Public Const ERRO_COMPRADOR_VINCULADO_PEDIDOCOMPRABAIXADO = 12058 'Parametro: icodigo
'Comprador %i nao pode ser excluido pois esta vinculado a um registro na tabela PedidodeCompraBaixado
Public Const ERRO_COMPRADOR_VINCULADO_COTACAO = 12059 'Parametro: icodigo
'Comprador %i nao pode ser excluido pois esta vinculado a um registro na tabela Cotacao
Public Const ERRO_COMPRADOR_VINCULADO_COTACAOBAIXADA = 12060 'Parametro: icodigo
'Comprador %i nao pode ser excluido pois esta vinculado a um registro na tabela CotacaoBaixada
Public Const ERRO_REQUISITANTE_NOME_DUPLICADO = 12062 'Parametro sNome
'O nome %s j� est� sendo utilizado por outro requisitante.
Public Const ERRO_LEITURA_REQUISICAOCOMPRABAIXADA1 = 12063 'Parametro: lCodRequisitante
'Erro na busca do Requisitante %l na tabela de Requisi��o de Compra Baixada.
Public Const ERRO_LEITURA_REQUISICAOMODELO1 = 12064 'Parametro lCodRequisitante
'Erro na busca do Requisitante %l na tabela de Requisi��o Modelo.
Public Const ERRO_REQUISITANTE_VINCULADO_REQCOMPRABAIXADA = 12065 'Parametro: lCodigo
'O Requisitante com c�digo %l n�o pode ser exclu�do pois est� relacionado a uma requisi��o de compra baixada.
Public Const ERRO_REQUISITANTE_VINCULADO_REQCOMPRAMODELO = 12066 'Parametro: lCodigo
'O Requisitante com c�digo %l n�o pode ser exclu�do pois est� relacionado a uma requisi��o de compra modelo.
Public Const ERRO_REQUISITANTE_NAO_CADASTRADO = 12067 'Parametro: lCodigo
'O Requisitante com c�digo %l n�o est� cadastrado.
Public Const ERRO_LEITURA_REQUISICAOCOMPRA1 = 12068 'Parametro: lCodRequisitante
'Erro na busca pelo Requisitante %l na tabela de Requisi��oCompra
Public Const ERRO_LOCK_REQUISITANTE = 12069 'Parametro: lCodigo
'N�o conseguiu fazer o lock do requisitante %l.
Public Const ERRO_EXCLUSAO_REQUISITANTE = 12070  'parametro: lCodigo
'Erro na exclus�o do requisitante %l na tabela de Requisitantes.
Public Const ERRO_REQUISITANTE_NOMERED_DUPLICADO = 12071 'Parametro: sNomeReduzido
'O nome reduzido %s j� est� sendo utilizado por outro requisitante.
Public Const ERRO_ATUALIZACAO_REQUISITANTE = 12072 'Par�metro : lCodigo
'Erro na atualiza��o do Requisitante %l na tabela de Requisitantes.
Public Const ERRO_INSERCAO_REQUISITANTE = 12073 'Par�metro: lCodigo
'Erro na Inser��o do Requisitante %l na tabela de Requisitantes.
Public Const ERRO_REQUISITANTE_VINCULADO_REQUISICAOCOMPRA = 12074 'parametro: lCodigo
'O Requisitante %l n�o pode ser exclu�do pois est� sendo utilizado em uma Requisi��o de Compra.
Public Const ERRO_EXCLUSAO_TIPO_BLOQUEIO_ALCADA = 12075  'Sem parametros
'N�o � poss�vel excluir o Tipo de Bloqueio de Al�ada.
Public Const ERRO_ALTERACAO_TIPO_BLOQUEIO_ALCADA = 12076 'Sem parametro
'N�o � poss�vel alterar o Tipo de Bloqueio de Al�ada.
Public Const ERRO_ATUALIZACAO_TIPODEBLOQUEIOPC = 12077 'Parametro: iCodigo
'Erro na atualiza��o do Tipo de Bloqueio %i na tabela de Tipo de Bloqueio de Pedido de Compra.
Public Const ERRO_TIPODEBLOQUEIOPC_MESMA_DESCRICAO = 12078 ' Parametro sDescricao
'J� existe no Banco de Dados Tipo de Bloqueio com a descri��o %s
Public Const ERRO_INSERCAO_TIPODEBLOQUEIOPC = 12079 'Parametro iCodigo
'Erro na inser��o do Tipo de Bloqueio %i na tabela de Tipos de BloqueioPC
Public Const ERRO_LEITURA_TIPODEBLOQUEIOPC = 12080 'Parametro iCodigo
'Erro na leitura do Tipo de Bloqueio %i na tabela de Tipos de BloqueioPC
Public Const ERRO_TIPODEBLOQUEIOPC_MESMO_NOME = 12081 'Parametro sNomeReduzido
'J� existe no Banco de Dados Tipo de Bloqueio com nome reduzido %s
Public Const ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO = 12082 'Parametro iCodigo
'O Tipo de Bloqueio de Pedido de Compra %i n�o est� cadastrado
Public Const ERRO_LOCK_TIPODEBLOQUEIOPC = 12083 'Parametro iCodigo
'N�o conseguiu fazer o lock do Tipo de Bloqueio de Pedido de Compra %i
Public Const ERRO_LEITURA_BLOQUEIOSPC = 12084 'Parametro iTipoBloqueio
'Erro na leitura do Tipo de Bloqueio %i da tabela de BloqueiosPC
Public Const ERRO_EXCLUSAO_TIPODEBLOQUEIOPC = 12085 'Parametro:iCodigo
'Erro na exclus�o do Tipo de Bloqueio de Pedido de Compra %i da tabela de TiposDeBloqueioPC
Public Const ERRO_TIPODEBLOQUEIOPC_VINCULADO_BLOQUEIOSPC = 12086 'Parametro iCodigoTipoBloqueioPC
'N�o � poss�vel excluir o Tipo de Bloqueio com o c�digo %i pois ele est� sendo utilizado por um Bloqueio de Pedido de Compra
Public Const ERRO_MESESMEDIATEMPORESSUP_NAO_PREENCHIDO = 12087 'Sem parametros
'Meses para tempo de ressuprimento n�o foi preenchido.
Public Const ERRO_MESESCONSUMOMEDIO_NAO_PREENCHIDO = 12088 'Sem parametros
'Meses de consumo m�dio n�o foi preenchido.
Public Const ERRO_LEITURA_BLOQUEIOS_PC_LIBERACAO = 12089
'Erro na leitura dos Bloqueios de Pedidos de Compras para liberar.
Public Const ERRO_PEDIDOCOMPRA_NAO_CADASTRADO = 12090 'Parametro lCodigo
'O Pedido de Compra com c�digo %l n�o est� cadastrado.
Public Const ERRO_USUARIO_NAO_COMPRADOR = 12091 'Parametro sCodUsuario
'O usu�rio %s n�o tem acesso a essa tela pois s� � acess�vel � compradores
Public Const ERRO_PEDIDOCOTACAO_NAO_SELECIONADO = 12092 'Sem parametros
'Um Pedido de Cota��o deve ser selecionado.
Public Const ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO = 12093 'Parametros lCodigo
'O Pedido de Cota��o %l n�o est� cadastrado no Banco de Dados.
Public Const ERRO_ITEM_PEDCOTACAO_VINCULADO_ITEM_PEDCOMPRA = 12094 'Parametros lCodigo
'Um item do Pedido de Cota��o %l est� vinculado � um item de um Pedido de Compra.
Public Const ERRO_ITEM_PEDCOTACAO_VINCULADO_CONCORRENCIA = 12095 'Parametros lCodigo
'Um item do Pedido de Cota��o %l est� vinculado a uma Concorr�ncia.
Public Const ERRO_PEDIDO_COTACAO_AVULSO = 12096 'Sem parametros
'Cota��es desvinculadas de Requisi��es n�o geram Pedidos de Compra.
Public Const ERRO_CONDICAO_PAGTO_NAO_PREENCHIDA = 12097 'Sem parametros
'A Condi��o de Pagamento deve ser informada.
Public Const ERRO_PRECOS_ITENS_CONDPAGTO_NAO_PREENCHIDOS = 12098 'Sem parametros
'Os pre�os dos itens para a Condi��o de Pagamento escolhida devem estar preenchidos.
Public Const ERRO_QUANTENTREGA_MAIOR_QUANTCOTACAO = 12099 'Sem parametros
'A quantidade dispon�vel para entrega deve ser menor ou igual a quantidade de cota��o.
Public Const ERRO_ITEMPEDCOTACAO_VINCULADO_CONCORRENCIA1 = 12100 'Parametros iCondPagtoPrazo, iIndice1
'O pre�o unit�rio para a condi��o de pagamento %i do item %i n�o pode ser apagado pois esse item j� est� relacionado com uma concorr�ncia.
Public Const ERRO_LOCK_PEDIDO_COTACAO = 12107 'Parametro lCodigo.
'N�o conseguiu fazer o lock do pedido de cotacao %l .
Public Const ERRO_LEITURA_OBSERVACAO = 12110 'Parametro sObservacao
'Erro na leitura da Observacao %s.
Public Const ERRO_INSERCAO_OBSERVACAO = 12111 'Parametros sObservacao
'Erro na inser��o da observa��o %s.
Public Const ERRO_INSERCAO_ITENSCOTACAO = 12112 'Parametro lCodigo
'Erro na inser��o dos itens de cota��o do pedido de cota��o %l.
Public Const ERRO_ATUALIZACAO_PEDIDOCOTACAO = 12113 'Parametros lCodigo
'Erro na atualiza��o dos dados do pedido de cota��o %l
Public Const ERRO_ATUALIZACAO_ITENSCOTACAO = 12114 'Parametros lCodigo
'Erro na atualiza��o dos itens de cota��o do pedido de cota��o %l
Public Const ERRO_EXCLUSAO_PEDIDO_COTACAO = 12115 'Parametro lCodigo
'Erro na exclus�o do pedido de cota��o %l.
Public Const ERRO_AUSENCIA_ITENS_PEDIDOCOTACAO = 12118 'Sem parametros
'N�o existem itens para o Pedido de Cota��o.
Public Const ERRO_LEITURA_COTACAOPRODUTOITEMRC = 12119 'Sem parametros
'Erro na leitura da tabela CotacaoProdutoItemRC
Public Const ERRO_COTACAO_NAO_CADASTRADA = 12120 'Parametro lCotacao
'A cota��o %l n�o est� cadastrada.
Public Const ERRO_OBSERVACAO_NAO_CADASTRADA = 12121 'Parametros lObservacao
'A observa��o %l n�o est� cadastrada.
Public Const ERRO_ITEMPEDCOTACAO_NAO_ENCONTRADO = 12122 'Parametro lCodigo
'N�o foram encontrados itens para o pedido de cota��o %l.
Public Const ERRO_ITENSCOTACAO_NAO_ENCONTRADOS = 12123 'Parametro lCodigo
'N�o foram encontrados itens de cota��o para o pedido de cota��o %l.
Public Const ERRO_REQUISICAOCOMPRA_NAO_ENVIADA = 12124 'Par�metros: lCodigo
'A Requisi��o de Compras de c�digo %l n�o foi enviada.
Public Const ERRO_LOCK_REQUISICAOCOMPRA = 12126 'Par�metros: lCodigo
'Erro na tentativa de fazer "lock" na tabela Requisi��o Compra com Requisi��o de c�digo %l.
Public Const ERRO_QUANTCANCELAR_SUPERIOR_QUANTDISPONIVEL_CANCELAR = 12127 'Par�metros: sProduto e dQuantMaximaCancelar
'A quantidade informada do produto %s � inv�lida. A quantidade m�xima a cancelar � %d.
Public Const ERRO_REQUISICAO_NAO_CARREGADA = 12128 'Sem par�metros
'Nenhuma requisi��o foi trazida para a tela.
Public Const ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA = 12129 'Par�metros: lCodigo
'A Requisi��o de Compra de c�digo %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_REQUISITANTE_NAO_CADASTRADO1 = 12130 'Par�metros: sNomeReduzido
'O Requisitante com o Nome Reduzido %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_BAIXA_ITEMRC_VINCULADO_ITEMPC_NAO_BAIXADO = 12131 'Par�metros: lCodigo, lNumIntDoc
'A Requisi��o de c�digo %l n�o pode ser baixada pois o item de n�mero interno %l est� vinculado a um item
'de pedido de venda n�o baixado.
Public Const ERRO_BAIXA_ITEMRC_VINCULADO_ITEMCONCORRENCIA_NAO_BAIXADO = 12132 'Par�metros: lCodigo, lItemConcorrencia
'A Requisi��o de c�digo %l n�o pode ser baixada pois o Item de Concorr�ncia %l est� vinculado a um item de concorr�ncia n�o baixado.
Public Const ERRO_LEITURA_ITEMRCITEMCONCORRENCIA = 12133 'Par�metros: lItemReqCompra
'Erro na leitura da tabela Item Requisi��o Compra Item Concorr�ncia com Item Requisi��o de Compra %l.
Public Const ERRO_LOCK_ITENSREQCOMPRA = 12134 'Par�metros: lCodRequisicao
'Erro na tentativa de fazer lock na tabela Itens Requisi��o Compra
'com Requisi��o de c�digo %l.
Public Const ERRO_EXCLUSAO_ITENSREQCOMPRAS = 12135 'Par�metros: lCodRequisicao
'Erro na tentativa de excluir item da Requisi��o Compra com o c�digo %l
Public Const ERRO_EXCLUSAO_COTACAOPRODUTOITEMRC = 12136 'Par�metros: lItemReqCompra
'Erro na tentativa de excluir Cota��o Produto Item Requisi��o Compras com Item Requisi��o de Compra %l.
Public Const ERRO_ATUALIZACAO_ITENSREQCOMPRA = 12137 'Par�metros: lNumIntDoc
'Erro na atualiza��o de Itens Requisi��o Compra de n�mero interno %l.
Public Const ERRO_INSERCAO_ITENSREQCOMPRABAIXADOS = 12138 'Par�metros: lNumIntDoc
'Erro na Inser��o do item Requisi��o de Compra n�mero interno %l na tabela Itens Requisi��o Compra Baixados.
Public Const ERRO_EXCLUSAO_REQUISICAOCOMPRA = 12139 'Par�metros: lCOdigo
'Erro na tentativa de excluir a Requisi��o Compra de c�digo %l da tabela Requisi��o Compra.
Public Const ERRO_INSERCAO_REQUISICAOCOMPRABAIXADA = 12140 'Par�metros: lCodigo 'OK ??? passe o c�digo da req como par�metro
'Erro na tentativa de inserir a Requisi��o de c�digo %l na tabela Requisi��o Compra Baixada.
Public Const ERRO_ITEMRC_DESVINCULADO_ITEMPC = 12141 'Par�metros: lNumIntDocItemPC
'O Item %l de Requisi��o de Compras n�o possui nenhum v�nculo com Pedidos de Compra.
Public Const ERRO_ITENSREQCOMPRA_NAO_CADASTRADO = 12142 'Par�metros: lNumIntDocItemRC, lCodRequisicao
'O Item de n�mero interno %l da Requisi��o de Compras de c�digo %l n�o est� cadastrado no
'Banco de dados.
Public Const ERRO_QUANTCANCELADA_MAIOR = 12143 'Sem par�metros
'A quantidade a cancelar n�o pode ser maior que a quantidade a receber.
Public Const ERRO_COMPRADOR_NAO_CADASTRADO1 = 12144  'Parametro: sCodUsuario
'O comprador com o codigo de usu�rio %s n�o esta cadastrado
Public Const ERRO_PEDIDOCOMPRA_JA_CADASTRADO = 12145 'lCodigo
'O Pedido de Compra %l j� est� cadastrado no Banco de Dados.
Public Const ERRO_ATUALIZACAO_ITENSPEDCOMPRA = 12146 'Sem parametro
'Erro na atualiza��o da tabela de Itens de Pedido de Compra.
Public Const ERRO_LEITURA_LOCALIZACAOITENSPC = 12147 'Sem parametro
'Erro na leitura da tabela de Localiza��oItensPC.
Public Const ERRO_PEDIDOCOMPRA_GERADO = 12148 'Parametro lCodigo
'O Pedido de Compra %l � um pedido de compra gerado.
Public Const ERRO_LOCK_ITENSPEDCOMPRA = 12149 'Sem parametro
'Erro na tentativa de lock em ItensPedCompra.
Public Const ERRO_EXCLUSAO_ITENSPEDCOMPRA = 12150 'Parametro lCodigo
'Erro na exclus�o do Item com n�mero interno %l da tabela de Itens de Pedido de Compra.
Public Const ERRO_INSERCAO_BLOQUEIOPC = 12151 'Parametro lCodigo
'Erro na inser��o do Bloqueio de Pedido de Compra %l.
Public Const ERRO_INSERCAO_LOCALIZACAOITENSPC = 12152 'Parametros lNumIntDoc
'Erro na inser��o do item com n�mero interno %l na tabela de Localiza��oItensPC.
Public Const ERRO_INSERCAO_ITENSPEDCOMPRA = 12153 'Parametro lNumIntDoc
'Erro na inser��o do item com n�mero interno %l do Pedido de Compra.
Public Const ERRO_INSERCAO_PEDIDOCOMPRA = 12154 'Parametro lCodigo
'Erro na inser��o do Pedido de Compra %l na tabela de PedidoCompra.
Public Const ERRO_OBSERVACAO_INEXISTENTE = 12155 'Parametro lNumInt
'A observa��o com o n�mero interno %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_DATAENVIO_INFERIOR_DATAPEDIDO = 12156 'Sem parametro
'A Data de Envio deve ser maior ou igual a Data do Pedido.
Public Const ERRO_PRODUTO_DESVINCULADO_ITEM = 12157 'Parametro sDescricao
'O produto %s n�o est� presente em nenhum dos itens do Pedido de Compras
Public Const ERRO_PRODUTO_ITEM_DISTRIBUICAO_VAZIO = 12158 'Sem parametro
'N�o � poss�vel deixar um Item de Distribui��o sem Produto
Public Const ERRO_PEDCOMPRA_BAIXADO = 12159 'Parametro lCodigo
'O Pedido de Compra com c�digo %l est� baixado.
Public Const ERRO_LOCK_OBSERVACAO = 12160 'Sem Parametros
'Erro na tentativa de lock na tabela de Observa��o.
Public Const ERRO_ALMOXARIFADO_ITEM_DISTRIBUICAO_VAZIO = 12161 'Sem parametro
'N�o � poss�vel deixar um Item de Distribui��o sem Almoxarifado
Public Const ERRO_AUSENCIA_ITENS_PC = 12162 'Sem parametros
'N�o existem itens para o Pedido de Compra.
Public Const ERRO_DATALIMITE_ITEM_INFERIOR_DATAPEDIDO = 12163 'Parametro iItem 'OK que � iIndice (se eu estiver em outra tela e precisar desse erro n�o vou entender)
'A Data Limite do item %i � menor que a Data do Pedido
Public Const ERRO_LEITURA_PEDIDOCOMPRA_BAIXADO = 12164 'Parametro lCodigo
'Erro na leitura do Pedido de Compra Baixado %l.
Public Const ERRO_PEDCOMPRA_BAIXADO_EXCLUSAO = 12165 'Parametro lCodigo
'O Pedido de Compra com c�digo %l est� baixado. N�o pode ser exclu�do.
Public Const ERRO_DATALIMITE_INFERIOR_DATAPEDIDO = 12166 'Sem parametros
'A Data Limite deve ser maior ou igual a Data do Pedido
Public Const ERRO_ALIQUOTA_IGUAL_100 = 12167 'Sem parametros
'A al�quota deve ter um valor inferior a 100%
Public Const ERRO_VALORTOTAL_PC_NEGATIVO = 12168
'Valor Total do Pedido de Compra � negativo.
Public Const ERRO_PEDIDOCOMPRA_ENVIADO = 12169 'Parametro lcodigo
'O Pedido de Compra com c�digo %l j� foi enviado.
Public Const ERRO_ITEMPEDCOMPRA_INEXISTENTE = 12170 'Parametro lNumIntDoc
'Item do Pedido de Compra  com n�mero interno %l n�o est� cadastrado.
Public Const ERRO_PEDCOMPRA_AUSENCIA_ITENS = 12171 'Parametro lCodigo
'Pedido de Compra %l n�o possui itens.
Public Const ERRO_EXCLUSAO_PEDIDOCOMPRA = 12172 'Parametro lCodigo
'Erro na exclus�o do Pedido de Compra %s.
Public Const ERRO_EXCLUSAO_BLOQUEIOSPC = 12173
'Erro na exclus�o de Bloqueios de Pedido de Compras.
Public Const ERRO_USUARIO_NAO_PREENCHIDO2 = 12175 'Sem Parametros
'O preenchimento do Nome Reduzido do Usu�rio � obrigat�rio.
Public Const ERRO_LIMITE_MENSAL_MENOR_LIMITE_OPERACAO = 12177 'Sem Par�metros
'O Limite Mensal n�o deve ser menor que o Limite de Opera��o.
Public Const ERRO_DATAENVIODE_MAIOR_DATAENVIOATE = 12178  'Sem parametros
'Data de Envio De maior que Data de Envio At�.
Public Const ERRO_NUM_PEDIDOS_SELECIONADOS_SUPERIOR_MAXIMO = 12179 'Sem parametros
'O n�mero de pedidos selecionados � superior ao n�mero m�ximo poss�vel para
'baixa. Restrinja mais a sele��o para que o n�mero de pedidos lidos diminua.
Public Const ERRO_LEITURA_PEDIDOS_COMPRA_BAIXA_PC = 12180 'Sem parametros
'Erro na leitura de Pedidos de Compra para baixa.
Public Const ERRO_REQUISICAOCOMPRA_INEXISTENTE = 12181 'Sem parametros
'N�o foi encontrada nenhuma Requisi��o de Compras de acordo com a sele��o informada.
Public Const ERRO_AUSENCIA_REQUISICOES_BAIXAR = 12182 'Sem parametros
'Deve haver pelo menos uma Requisi��o marcada para ser baixada.
Public Const ERRO_REQUISICAO_INICIAL_MAIOR = 12183 'Sem parametros
'O n�mero da Requisi��o Inicial n�o pode ser maior que o da Requisi��o Final.
Public Const ERRO_REQUISITANTE_INICIAL_MAIOR = 12184 'Sem parametros
'O c�digo do requisitante inicial n�o pode ser maior que o do requisitante final.
Public Const ERRO_DATALIMITEDE_MAIOR = 12185 'Sem parametros
'A data limite inicial n�o pode ser maior que a data limite final.
Public Const ERRO_INSERCAO_ITENSPEDCOMPRABAIXADOS = 12186 'Par�metro tItemPedido.lNumINtDoc
'Erro na inser��o do item %l na tabela de Itens de Pedido de Compra Baixados
Public Const ERRO_RESIDUO_NAO_PREENCHIDO = 12187 'Sem parametros
'O preenchimento do Res�duo � obrigat�rio.
Public Const ERRO_FORNECEDORFILIALPRODUTO_NAO_CADASTRADA = 12188 'Par�metros: sProduto, iFilialForn, lFornecedor
'A associa��o do Produto %s com a Filial %i do Fornecedor %l n�o est� cadastrada
'no Banco de dados.
Public Const ERRO_EXCLUSAO_COTACAOITEMCONCORRENCA = 12189 'Sem par�metros
'Erro na exclus�o de registro na tabela Cota��o Item Concorr�ncia.
Public Const ERRO_INSERCAO_COTACAO = 12190 'Sem parametro
'Erro na tentativa de inser��o na tabela de Cota��o.
Public Const ERRO_GRID_FORN_LINHA_NAO_SELECIONADA = 12191 'Sem parametros
'Deve ser selecionada alguma linha do Grid de Fornecedores.
Public Const ERRO_PED_COTACAO_NAO_GERADO = 12192 'Sem parametros
'� necess�rio gerar Pedidos de Cota��o antes de imprimir.
Public Const ERRO_INSERCAO_COTACAOCONDPAGTO = 12193 'Sem parametros
'Erro na insercao de Condi��o de Pagamento na tabela CotacaoCondPagto.
Public Const ERRO_GRID_PRODUTOS_VAZIO = 12194 'Sem parametros
'O Grid de Produtos est� vazio
Public Const ERRO_USUARIO_INEXISTENTE = 12195 'Parametro: sNomeReduzido
'O Usu�rio com Nome Reduzido %s n�o existe.
Public Const ERRO_PRODUTO_REPETIDO_GRID_PRODUTOS = 12196 'Parametro: sCodigoProduto
'O Produto %s s� pode aparecer uma vez no Grid de Produtos.
Public Const ERRO_QUANTIDADE_COTAR_NAO_PREENCHIDA = 12197 'Parametro: sCodigo
'Produto %s n�o tem quantidade a cotar preenchida
Public Const ERRO_INSERCAO_PEDIDOCOTACAO = 12198 'Parametro: lNumIntPedCotacao
'Erro na inser��o do Pedido de Cota��o com n�mero interno %l.
Public Const ERRO_INSERCAO_ITEMPEDCOTACAO = 12199 'Parametro: lNumInt
'Erro na inser��o do Item de Pedido de Cota��o com n�mero interno %l.
Public Const ERRO_INSERCAO_COTACAOPRODUTO = 12200 'Sem parametros
'Erro na insercao de Cota��o Produto.
Public Const ERRO_PRODUTO_SEM_FORNECEDOR_ESCOLHIDO = 12201 'Parametro: sCodProduto
'Para o Produto %s n�o foi escolhido nenhum fornecedor.
Public Const ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR = 12202 'Parametros: sFornNomeReduzido,sCodProduto
'N�o existe Filial do Fornecedor %s cadastrada para o Produto %s.
Public Const ERRO_FILIALEMPRESA_NAO_CADASTRADA1 = 12203  'parametro sNome
'FilialEmpresa %s n�o est� cadastrada.
Public Const ERRO_GRID_REQUISICOES_VAZIO = 12204 'Sem parametros
'O Grid de Requisi��es est� vazio.
Public Const ERRO_REQUISICAOINICIAL_MAIOR_REQUISICAOFINAL = 12205 'Sem parametros
'Requisi��o Inicial n�o pode ser maior do que a Requisi��o Final.
Public Const ERRO_GRID_REQUISICAO_NAO_SELECIONADO = 12206 'Sem parametros
'N�o foi selecionado Requisi��o do Grid de Requisi��es.
Public Const ERRO_GRID_PRODUTOS_NAO_SELECIONADO = 12207 'Sem parametros
'N�o foi selecionado Produto do Grid de Produtos.
Public Const ERRO_GRID_ITENS_REQUISICAO_NAO_SELECIONADO = 12208 'Sem parametros
'N�o foi selecionado Item de Requisi��o no Grid de Itens de Requisi��es.
Public Const ERRO_QUANT_COTAR_ITEM_NAO_PREENCHIDA = 12209 'Parametro iLinhaGrid
'O Item de Requisi��o %i n�o est� com a quantidade a cotar preenchida.
Public Const ERRO_INSERCAO_COTACAOPRODUTOITEMRC = 12210 'Sem parametros
'Erro na tentativa de inser��o na tabela CotacaoProdutoItemRC.
Public Const ERRO_LEITURA_GERACAOPTOPEDIDO = 12211 'Sem Parametros
'Erro na Leitura das Tabelas Produtos, EstoqueProduto, ProdutoFilial e Almoxarifado.
Public Const ERRO_PEDIDOCOMPRA_NAO_ENVIADO = 12212 'Parametro lCodigo
'O pedido de compras com o c�digo %l n�o foi enviado.
Public Const ERRO_PEDIDOCOMPRA_NAO_GERADO = 12213 ' Parametro lCodigo
'O Pedido de Compra %l n�o � um pedido de compra gerado.
Public Const ERRO_CLIENTE_DESTINO_NAO_PREENCHIDO = 12214
'O cliente de destino n�o foi informado.
Public Const ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA = 12215
'A filial do fornecedor de destino n�o foi preenchida.
Public Const ERRO_CONDICAOPAGTO_NAO_DISPONIVEL = 12216 'Sem par�metros
'A Condi��o de Pagamento n�o pode ser � Vista. Escolha uma Condi��o de Pagamento A Prazo.
Public Const ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA1 = 12217 'Par�metros: lNumIntDoc
'A Requisi��o de Compras de n�mero interno %l n�o est� cadastrada no Banco de dados.
Public Const ERRO_TIPOTRIBUTACAO_NAO_CADASTRADA = 12218 'Par�metros: iTipo
'O Tipo de Tributa��o de c�digo %i n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_TIPOSTRIBUTACAOMOVTO = 12219 'Sem Par�metros
'Erro na leitura da Tabela TiposTributacaoMovto.
Public Const ERRO_LEITURA_ITEMRCITEMCONCORRENCIA1 = 12220 'Sem par�metros
'Erro na leitura da tabela ItemRCItemConcorrencia.
Public Const ERRO_FILIALCOMPRA_NAO_PREENCHIDA = 12221 'Sem par�metros
'A Filial de Compra deve ser informada.
Public Const ERRO_REQUISICAOCOMPRA_ENVIADA = 12222 'Par�metros: lCodigo
'A Requisi��o de Compras de c�digo %l j� foi enviada.
Public Const ERRO_REQUISICAO_COMPRA_BAIXADA = 12223 'Par�metros: lCodigo
'A Requisi��o de compras de c�digo %l j� est� baixada.
Public Const ERRO_CODIGO_REQUISICAO_COMPRA_EXISTENTE = 12224 'Par�metros: lCodigo
'J� existe no Banco de Dados a Requisi��o Compras com o c�digo %l.
Public Const ERRO_LEITURA_REQUISICAO_COMPRA_BAIXADA = 12225 'Par�metros: lCodigo
'Erro na leitura da Requisi��o de c�digo %l da tabela RequisicaoCompraBaixada.
Public Const ERRO_INSERCAO_REQUISICAO_COMPRA = 12226 'Par�metros: lCodigo
'Erro na inser��o da Requisi��o Compras de c�digo %l.
Public Const ERRO_INSERCAO_ITEMREQCOMPRA = 12227 'Par�metros: lNumIntDocItem
'Erro na tentativa de inserir o item de Requisi��o de Compras de n�mero interno %l.
Public Const ERRO_ITENSREQCOMPRA_NAO_CADASTRADO1 = 12228 'Par�metros: lCodReqCompras
'O item da Requisi��o de compras de c�digo de %l n�o est� cadastrado.
Public Const ERRO_REQUISICAO_COMPRA_ENVIADA = 12229 'Par�metros: lCodigo
'N�o � poss�vel excluir a Requisi��o de compras de c�digo %l porque ela j� foi enviada.
Public Const ERRO_ALTERACAO_REQUISICAOCOMPRA = 12230 'Par�metros: lCodigo
'Erro na atualiza��o da Requisi��o de Compras de C�digo %l.
Public Const ERRO_REQUISICAO_COMPRAS_AUSENCIA_ITENS = 12231 'Par�metros: lCodigo
'A Requisi��o de Compras de c�digo %l n�o possui itens.
Public Const ERRO_LOCK_ITEMREQCOMPRA = 12232 'Par�metros: lNumIntitem
'Erro na tentativa de fazer "lock" no item de n�mero interno %l de Requisi��o de Compras.
Public Const ERRO_EXCLUSAO_ITEMREQCOMPRA = 12233 'Par�metros: lNumIntitem
'Erro na exclus�o do Item de n�mero iterno %l de Requisi��o de Compras.
Public Const ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS = 12234 'Par�metros: iCodFilial, sFornNomeRed, sCodProduto
'A Filial %i do Fornecedor %s n�o est� associada ao Produto %s.
Public Const ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO = 12235 'Sem par�metros
'O Fornecedor Destino n�o foi preenchido.
Public Const ERRO_FILIALCLIENTE_DESTINO_NAO_PREENCHIDA = 12236 'Sem par�metros
'A Filial do Cliente destino n�o foi informada.
Public Const ERRO_LEITURA_ITENSREQMODELO = 12237 'Par�metros: lCodReqModelo
'Erro na leitura da tabela ItensReqModelo da Requis�o modelo de c�digo %l.
Public Const ERRO_REQUISICAO_MODELO_AUSENCIA_ITENS = 12239 'Par�metros: lReqModelo
'A Requisi��o Modelo %l n�o possui itens.
Public Const ERRO_REQUISITANTE_NAO_PREENCHIDO = 12240 'Sem par�metros
'O Requisitante deve ser preenchido.
Public Const ERRO_GRID_FORNECEDOR_NAO_PREENCHIDO = 12241 'Par�metros: iLinha
'O Campo Fornecedor da linha %i do Grid n�o foi preenchido.
Public Const ERRO_FILIALFORN_NAO_ENCONTRADA_ASSOCIADA = 12242 'Par�metros: sCodFornecedor, sCodProduto
'A Filial do Fornecedor %s n�o foi encontrada ou n�o est� associada ao Produto %s.
Public Const ERRO_GRIDITENS_VAZIO = 12243 'Sem par�metros
'N�o existem itens no Grid para gravar.
Public Const ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA = 12244 'Sem par�metros
'A Filial Empresa de destino n�o foi informada.
Public Const ERRO_GRID_QUANTIDADE_NAO_PREENCHIDA = 12245 'Par�metros: iLinha
'A Quantidade da linha %i do Grid n�o foi preenchida.
Public Const ERRO_LEITURA_REQUISICAOMODELO = 12246 'Par�metros: lCodReqModelo
'Erro na leitura da Requisi��o Modelo de C�digo %l.
Public Const ERRO_LEITURA_REQUISICAOMODELO2 = 12247 'Par�metros: lNumIntReqModelo
'Erro na leitura da Requisi��o Modelo de N�mero Interno %l.
Public Const ERRO_INSERCAO_REQUISICAOMODELO = 12248 'Par�metros: lCodReqModelo
'Erro na inser��o da Requisi��o Modelo de c�digo %l.
Public Const ERRO_LOCK_REQUISICAOMODELO = 12249 'Par�metros: lCodReqModelo
'Erro no "lock" da Requisi��o Modelo de c�digo %l.
Public Const ERRO_ALTERACAO_REQUISICAOMODELO = 12250 'Par�metros: lCodReqModelo
'Erro na atualiza��o da Requisi��o Modelo de c�digo %l.
Public Const ERRO_ITEMREQMODELO_NAO_CADASTRADO = 12251 'Par�metros: lCodReqModelo
'O item da Requisi��o Modelo de c�digo %l n�o est� cadastrado.
Public Const ERRO_LOCK_ITEMREQMODELO = 12252 'Par�metros: lNumIntReq
'Erro no lock no item de n�mero interno %l de Requisi��o Modelo.
Public Const ERRO_REQUISICAOMODELO_NAO_CADASTRADA = 12253 'Par�metros: lCodReqModelo
'A Requisi��o Modelo de c�digo %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_EXCLUSAO_REQUISICAOMODELO = 12254 'Par�metros: lCodReqModelo
'Erro na tentativa de excluir a Requisi��o Modelo de c�digo %l.
Public Const ERRO_EXCLUSAO_ITEMREQMODELO = 12255 'Par�metros: lNumIntItemReq
'Erro na tentativa de excluir o item de n�mero interno %l de Requisi��o Modelo.
Public Const ERRO_INSERCAO_ITEMREQMODELO = 12256 'Par�metros: lNumIntItemReq
'Erro na tentativa de gravar o Item de n�mero interno %l de Requisi��o Modelo.
Public Const ERRO_REQUISICAOMODELO_NAO_CADASTRADA1 = 12258 'Par�metros: lNumIntDoc
'A Requisi��o Modelo de n�mero interno %l n�o est� cadastrada.
Public Const ERRO_ATUALIZACAO_ITENSREQMODELO = 12259 'Par�metros: lNumItemReqModelo
'Erro na tentativa de atualizar o item de n�mero interno %l de Requisi��o Modelo.
Public Const ERRO_LEITURA_REQUISICOES_COMPRA_BAIXA_RC = 12260 'Sem parametros
'Erro na leitura de Requisi��es de Compra para baixa.
Public Const ERRO_NUM_REQUISICOES_SELECIONADAS_SUPERIOR_MAXIMO = 12261 'Sem parametros
'O n�mero de requisi��es selecionadas � superior ao n�mero m�ximo poss�vel para
'baixa. Restrinja mais a sele��o para que o n�mero de requisi��es lidas diminua.
Public Const ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO = 12262 'Parametros: lFornecedor, iFilialForn, sProduto,iFilialEmpresa
'O Fornecedor %l Filial %i n�o est� cadastrado na Tabela FornecedorProdutoFF para o Produto %s FilialEmpresa %i
Public Const ERRO_PEDIDOCOMPRA_NAO_ENCONTRADO = 12263 'Sem parametros
'N�o foi encontrado nenhum Pedido de Compra de acordo com a sele��o informada.
Public Const ERRO_COTACAO_VINCULADA_PEDIDOCOT_NAO_CADASTRADA = 12264 'Par�metros: lCodigo
'A Cota��o vinculada ao Pedido de Cota��o de c�digo %l n�o est� cadastrada no Banco de dados.
Public Const ERRO_ITEMCOTACAO_NAO_CADASTRADO = 12265 'Par�metros: lNumIntItem
'O Item de n�mero interno %l de Cota��o n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LOCK_ITEMCOTACAO = 12266 'Par�metros: lNumIntItem
'Erro na tentativa de fazer "lock" no Item de n�mero interno %l de Cota��o.
Public Const ERRO_EXCLUSAO_ITEMCOTACAO = 12267 'Par�metros: lNumIntItem
'Erro na exlus�o do Item de n�mero interno %l de Cota��o.
Public Const ERRO_INSERCAO_ITEMCOTACAOBAIXADO = 12268  'Par�metros: lNumIntItem
'Erro na inser��o do Item de n�mero interno %l de Cota��o na tabela ItensCotacaoBaixados.
Public Const ERRO_DATAVALIDADE_INICIAL_MAIOR = 12269 'Sem par�metros
'A data de validade inicial n�o pode ser maior que a data
'de validade final.
Public Const ERRO_LEITURA_PEDIDOCOTACAO1 = 12270 'Sem par�metros
'Erro na leitura da tabela PedidoCotacao.
Public Const ERRO_AUSENCIA_PEDIDOCOTACAO = 12271 'Sem par�metros
'N�o h� Pedidos de Cota��o para a Sele��o atual.
Public Const ERRO_NUM_PEDIDOS_SUPERIOR_MAXIMO = 12272 'Par�metros: NUM_MAX_PEDCOTACOES
'O n�mero de Pedidos de Cota��o da sele��o atual ultrapassou o limite que � %i.
Public Const ERRO_AUSENCIA_PEDCOTACAO_SELECIONADOS = 12273 'Sem par�metros
'N�o h� Pedidos de Cota��o selecionados para baixar.
Public Const ERRO_PEDCOTACAO_VINCULADO_PEDCOMPRA = 12274 'Par�metros: lCodPedCotacao, lCodPedCompra
'O Pedido de Cota��o de c�digo %l est� vinculado ao Pedido de Compra n�o Baixado de c�digo %l.
Public Const ERRO_PEDCOTACAO_VINCULADO_CONCORRENCIA = 12275 'Par�metros: lCodPedCotacao
'O Pedido de Cota��o de c�digo %l est� vinculado a uma Concorr�ncia n�o baixada.
Public Const ERRO_LEITURA_ITENSREQCOMPRATODOS = 12276 'Sem par�metros
'Erro na leitura da tabela ItensReqComprasTodos.
Public Const ERRO_LEITURA_ITENSCONCORRENCIATODOS = 12278  'Sem par�metros
'Erro na leitura da tabela ItensConcorrenciaTodos.
Public Const ERRO_LEITURA_COTACAOTODAS = 12279 'Sem par�metros
'Erro na leitura da tabela CotacaoTodas.
Public Const ERRO_MOTIVO_NAO_CADASTRADO1 = 12280 'Par�metros: sDescricao
'O Motivo de Descri��o %s n�o est� cadastrado no Banco de dados.
Public Const ERRO_LEITURA_QUANTIDADESSUPLEMENTARES = 12281 'Sem parametros
'Erro na leitura da tabela Quantidades Suplementares.
Public Const ERRO_LEITURA_MOTIVO = 12282 'Sem par�metros
'Erro na leitura da tabela Motivo.
Public Const ERRO_LEITURA_REQUISICAOCOMPRA2 = 12283 'Sem par�metros
'Erro na leitura da tabela Requisi��oCompra.
Public Const ERRO_QUANTIDADE_INFERIOR_INICIAL = 12284 'Sem parametros
'A quantidade atual � inferior � quantidade j� existente.
Public Const ERRO_COTACAO_NAO_CADASTRADA1 = 12285 'Sem Parametros
'A cota��o n�o est� cadastrada.
Public Const ERRO_MOTIVO_NAO_ENCONTRADO = 12286 'Parametro sMotivoEscolha
'O motivo de escolha %s n�o est� cadastrado.
Public Const ERRO_DATANECESSIDADE_ANTERIOR_DATAPEDIDO = 12287 'Sem Parametros
'A Data de Necessidade n�o pode ser anterior a Data do Pedido.
Public Const ERRO_FALTA_TIPO = 12288 'Sem Parametros
'Falta selecionar Tipo(s) de Produto.
Public Const ERRO_AUSENCIA_ITENS_GRID = 12289 'Sem Parametros
'Pelo menos uma linha do grid de itens deve ser preenchida.
Public Const ERRO_QUANTCOMPRAR_NAO_PREENCHIDA = 12290 'Sem parametros
'A quantidade a comprar do grid de Produtos deve ser preenchida.
Public Const ERRO_QUANTCOTACAO_DIFERENTE_QUANTCOMPRAR = 12291 'Parametro: sProduto
'A quantidade total selecionada nos itens de cota��o para o produto %s �
'diferente da quantidade a comprar informada.
Public Const ERRO_ATUALIZACAO_ITENSCONCORRENCIA = 12292 'Parametro: lCodConcorrencia
'Erro na tentativa de atualiza��o do Item de Concorr�ncia da Concorr�ncia %l.
Public Const ERRO_ATUALIZACAO_COTACAOITEMCONCORRENCIA = 12293 'Sem parametros
'Erro na tentativa de atualiza��o da tabela CotacaoItemConcorrencia.
Public Const ERRO_ATUALIZACAO_QUANTIDADESSUPLEMENTARES = 12294 'Sem parametros
'Erro na tentativa de atualiza��o da tabela QuantidadesSuplementares
Public Const ERRO_INSERCAO_QUANTIDADESSUPLEMENTARES = 12295 'Sem parametros
'Erro na tentativa de inser��o na tabela QuantidadesSuplementares
Public Const ERRO_INSERCAO_ITEMRCITEMCONCORRENCIA = 12296 'SEm parametros
'Erro na tentativa de inser��o na tabela ItemRCItemConcorrencia.
Public Const ERRO_ATUALIZACAO_ITEMRCITEMCONCORRENCIA = 12297 'Sem parametros
'Erro na tentativa de atualiza��o da tabela ItemRCItemConcorrencia.
Public Const ERRO_INSERCAO_COTACAOITEMCONCORRENCIA = 12298 'Sem parametros
'Erro na tentativa de inser��o na tabela CotacaoItemConcorrencia.
Public Const ERRO_REQUISICAO_NAO_SELECIONADA = 12299 'Sem par�metros
'Pelo menos uma requisi��o deve ser selecionada.
Public Const ERRO_QUANTCOTACAO_DIFERENTE_SOMAITENSREQ = 12300 'Sem par�metros
'A quantidade total selecionada nos itens de cota��o  � diferente da soma dos itens de requisi��o.
Public Const ERRO_ITEM_REQUISICAO_NAO_SELECIONADO = 12301 'Sem par�metros
'Pelo menos um item de requisi��o deve ser selecionado.
Public Const ERRO_NUM_REQUISICOES_SELECIONADAS_SUPERIOR_MAXIMO1 = 12302 'Sem parametros
'O n�mero de requisi��es selecionadas � superior ao n�mero m�ximo poss�vel.
'Restrinja mais a sele��o para que o n�mero de requisi��es lidas diminua.
Public Const ERRO_SELECAO_REQUISITANTE_AUTOMATICO = 12303
'N�o � poss�vel selecionar o Requisitante autom�tico.
Public Const ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_RC = 12304 'Sem par�metro
'Erro na leitura de pedidos de venda para a gera��o de requisi��es de compra.
Public Const ERRO_QUANTCOMPRAR_ITENSPED = 12305 'Sem parametros
'Ao menos um item de pedido de venda deve ter uma quantidade a comprar.
Public Const ERRO_PEDIDO_VENDA_NAO_ENCONTRADO = 12306 'Sem parametros
'N�o foi poss�vel encontrar nenhum Pedido de Venda de acordo com a sele��o informada.
Public Const ERRO_AUSENCIA_PEDVENDAS_GRID = 12307 'Sem parametros
'N�o h� pedidos de venda no grid Pedidos para selecionar.
Public Const ERRO_AUSENCIA_PEDVENDA_SELECIONADO = 12308 'Sem parametros
'Ao menos um pedido de venda deve ser selecionado.
Public Const ERRO_QUANTCOMPRAR_PRODUTOS_DIFERENTE = 12309 'Parametro sCodigoProduto
'A quantidade a comprar do grid de produtos � diferente da soma
'das quantidades a comprar do itens de pedido com o produto %s.
Public Const ERRO_INSERCAO_CONCORRENCIA = 12310 'Par�metros: lCodigo
'Erro na tentativa de inserir a Concorr�ncia de c�digo %l no Banco de Dados.
Public Const ERRO_LOCK_CONCORRENCIA = 12311 'Par�metros:lCodigo
'Erro no Lock da Concorr�ncia de c�digo %l.
Public Const ERRO_ATUALIZACAO_CONCORRENCIA = 12312 'Par�metros: lCodigo
'Erro na atualiza��o da Concorr�ncia de c�digo %l.
Public Const ERRO_LOCK_ITENSCONCORRENCIA = 12313 'Sem par�metros
'Erro na tentativa de fazer "lock" na tabela ItensConcorrencia.
Public Const ERRO_EXCLUSAO_ITENSCONCORRENCIA = 12314 'Par�metros: lNumIntItem
'Erro na tentativa de excluir o Item Concorr�ncia de n�mero interno %l.
Public Const ERRO_EXCLUSAO_COTACAOITEMCONCORRENCIA = 12315 'Par�metros: lNumIntItemConc
'Erro na exclus�o de Cota��oItemConcorr�ncia com Item Concorr�ncia de n�mero interno %l.
Public Const ERRO_LOCK_COTACAOITEMCONCORRENCIA = 12316 'Sem par�metros
'Erro na tentativa de fazer "lock" em Cota��oItemConcorr�ncia.
Public Const ERRO_LOCK_ITEMRCITEMCONCORRENCIA = 12317 'Sem par�metros
'Erro na tentativa de fazer "lock" em ItemRCItemConcorr�ncia.
Public Const ERRO_EXCLUSAO_ITEMRCITEMCONCORRENCIA = 12318 'Par�metros: lNumIntItemConc
'Erro na exclus�o de ItemRCItemConcorr�ncia com Item Concorr�ncia de n�mero interno %l.
Public Const ERRO_EXCLUSAO_QUANTIDADESSUPLEMENTARES = 12319 'Par�metros: lNumIntItemConc
'Erro na exclus�o de QuantidadesSuplementares com Item Concorr�ncia de n�mero interno %l
Public Const ERRO_INSERCAO_ITEMCONCONCORRENICA = 12320 'Par�metros: lNumIntItem
'Erro na tentativa de inserir o item de concorr�ncia de n�mero interno %l.
Public Const ERRO_QUANTCOMPRAR_MAIOR_RC = 12321 'Par�metros: dQuantComprar, dQuantidade
'A quantidade a comprar %d n�o pode ser maior que a quantidade do Item de Requisi��o de compras que � %d.
Public Const ERRO_INSERCAO_ITEMRCITEMPC = 12322 'Par�metros: lItemPedCompra, lItemRC
'Erro na inser��o de registros na tabela ItemRCItemPC com ItemPC %l e ItemRC %l.
Public Const ERRO_ITEMRCITEMCONCORRENCIA_NAO_CADASTRADO = 12323 'Par�metros: lItemRC, lItemConc
'O ItemRCItemConcorr�ncia com ItemRC = %l e Item Concorr�ncia %l n�o est� cadastrado.
Public Const ERRO_QUANTCOTACAO_DIFERENTE_QUANTITEMCONC = 12324 'Par�metros: sProduto, sFornecedor, iFilial
'A quantidade � comprar do item %s de Fornecedor %s e Filial de c�digo %i
'� diferente da quantide a comprar do mesmo item marcado no Grid de Cota��es.
Public Const ERRO_QUANTCOMPRAR_MENOR_QUANTCOMPRAR_RC = 12325 'Par�metros: dQuantidade, dQuantTotalRC
'A quantidade a Comprar %d tem que ser maior que a soma das quantidades a comprar dos Itens
'de Requisi��o que � %d.
Public Const ERRO_QUANTCOMPRAR_SUPERIOR_MAXIMA = 12326 'Par�metros: dQuantComprar, dQuantInicial
'A quantidade %d n�o pode ser superior a quantidade inicial %d.
Public Const ERRO_AUSENCIA_PEDCOTACAO_GRID = 12327 'Sem par�metros
'N�o h� gera��es de Pedidos de Cota��o no Grid de gera��o
'de Pedidos de Cota��o para selecionar.
Public Const ERRO_AUSENCIA_GERACAOPEDCOTACAO_SELECIONADA = 12328 'Sem par�metros
'Uma gera��o de pedido de cota��o deve ser selecionada.
Public Const ERRO_NENHUMA_COTACAO_SELECAO = 12329 'Sem par�metros
'N�o foi encontrada nenhuma cota��o na sele��o atual.
Public Const ERRO_NENHUM_TIPOPRODUTO_SELECIONADO = 12330 'Sem par�metros
'Nehum Tipo de Produto foi selecionado.
Public Const ERRO_CONSUMOMEDIO_NAO_PREENCHIDO = 12331 'Par�metros: iLinha
'O Consumo M�dio da linha %i do Grid n�o foi preenchido.
Public Const ERRO_INTERVALORESSUP_NAO_PREENCHIDO = 12332 'Par�metros: iLinha
'O Intervalo de Ressuprimento da linha %i do Grid n�o foi preenchido.
Public Const ERRO_NENHUM_PRODUTO_SELECIONADO = 12333 'Sem par�metros
'N�o foi encontrado nenhum Produto com a sele��o atual.
Public Const ERRO_QUANTNAOEXCLUSIVA_DIFERENTE_QUANTEXCLUSIVA = 12334 'Par�metros: sProduto
'A quantidade � comprar do Produto do %s do Grid de Produtos � diferente da
'quantidade � comprar do Grid de Cota��es.
Public Const ERRO_CONCORRENCIA_AUSENCIA_ITENS = 12335 'Par�metros: lCodConcorrencia
'A concorr�ncia de c�digo %l n�o possui Itens.
Public Const ERRO_INSERCAO_CONCORRENCIABAIXADA = 12336 'Par�metros: lCodConcorrencia
'Erro na inser��o da Concorr�ncia de c�digo %l na tabela ConcorrenciaBaixada.
Public Const ERRO_AUSENCIA_COTACOES_GRID = 12337 'Se par�metros
'N�o existem Cota��es no Grid de Cota��es.
Public Const ERRO_SELECAO_CONCORRENCIA = 12338 'Sem par�metros
'N�o h� Concorr�ncias cadastradas que obedecem sele��o atual.
Public Const ERRO_QUANTCOMPRAR_MAIOR_QUANTCOMPRARMAX = 12339 'Par�metros: dQuantComprar, dQuantMaxComprar
'A quantidade a comprar %d tem que ser menor que a quantidade m�xima a comprar que
'� %d.
Public Const ERRO_MOTIVO_NAO_CADASTRADO = 12340 'Par�metros: iCodigo
'O Motivo de C�digo %i n�o est� cadastrado no Banco de dados.
Public Const ERRO_LOCK_ITENSREQCOMPRA1 = 12341 'Par�metros:lNumIntDoc
'Erro no "lock" do Item de n�mero interno %l de Requisi��o de Compras.
Public Const ERRO_AUSENCIA_CONCORRENCIAS_GRID = 12342 'Sem par�metros
'N�o h� concorr�ncias no Grid concorr�ncias para selecionar.
Public Const ERRO_AUSENCIA_CONCORRENCIAS_SELECIONADAS = 12343 'Sem par�metros
'Uma concorr�ncia do Grid de concorr�ncias deve ser selecionada.
Public Const ERRO_ITEMREQ_NAO_SELECIONADO = 12344 'Sem par�metros
'Pelo menos um item de Requisi��o deve ser selecionado.
Public Const ERRO_CONCORRENCIA_NAO_CADASTRADA = 12345 'Par�metros: lCodigo
'A Concorr�ncia de c�digo %l n�o est� cadastrada no Banco de dados.
Public Const ERRO_LOCK_CONCORRECIA = 12346 'Par�metros: lCodigo
'Erro na tentativa de fazer "lock" na Concorr�ncia de c�digo %l.
Public Const ERRO_EXCLUSAO_CONCORRENCIA = 12347 'Par�metros: lCodigo
'Erro na exclus�o da Concorr�ncia de c�digo %l.
Public Const ERRO_CODIGO_INICIAL_MAIOR_FINAL = 12348 'Sem par�metros
'O c�digo inicial n�o pode ser maior que o final.
Public Const ERRO_ITEMREQCOMPRA_NAO_CADASTRADO = 12349 'Par�metros: sProduto, lReqCompra
'O Item com o Produto %s de Requisi��o de Compras de numero interno %l, n�o est�
'cadastrado no banco de dados.
Public Const ERRO_ITEMCONCORRENCIA_NAO_VINCULADO_ITEMCOTACAO = 12350 'Par�metros: iLinha
'O Item de Concorr�ncia da linha %i do Grid de Produtos n�o est� vinculado
'a nenhuma linha do Grid de Cota��es.
Public Const ERRO_COTACAOITEMCONCORRENCIA_NAO_CADASTRADA = 12351 'Par�metros: lNumIntConcorrencia, lItemCotacao)
'Cota��oItemConcorr�ncia com Item Concorr�ncia %l e Item Cota��o %l n�o cadastrada.
Public Const ERRO_NENHUM_ITEMCONC_SELECIONADO = 12352 'Sem par�metros
'Nenhuma linha do Grid de Produtos foi selecionada.
Public Const ERRO_ITEMREQCOMPRA_NAO_CADASTRADO1 = 12353 'Par�metros: lNumIntItem
'O Item de n�mero interno %l de Requisi��ode Compras n�o est� cadastrado.
Public Const ERRO_QUANTCOMPRAR_NAO_EXCLUSIVA_DIFERENTE = 12354 'Par�metros: dQuantProd1, dQuantSobraCot
'A quantidade a Comprar n�o exclusiva %d � diferente da quantidade
'que sobrou das Cota��es, que � %d.
Public Const ERRO_PEDCOMPRA_NAO_GERADO = 12355 'Parametro: lCodigoPedido
'O Pedido de Compras com c�digo %l n�o � um pedido gerado.
Public Const ERRO_LEITURA_ITEMPEDCOTACAO1 = 12356 'SEm parametros
'Erro na leitura da tabela ItemPedCotacao
Public Const ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO1 = 12357 'parametro:lCodigoPedCompra
'N�o existe Pedido de Cota��o que tenha gerado o Pedido de Compra %lCodigo.
Public Const ERRO_FORNECEDOR_ENTREGA_NAO_PREENCHIDO = 12358 'Sem parametros
'O campo Fornecedor de Entrega n�o foi preenchido.
Public Const ERRO_LEITURA_ITEMPEDCOTACAOTODOS = 12359 'Sem par�metros
'Erro na leitura da tabela ItemPedCotacaoTodos.
Public Const ERRO_LEITURA_ITENSCOTACAOTODOS = 12360 'Sem par�metros
'Erro na leitura da tabela ItensCotacaoTodos.
Public Const ERRO_EXCLUSAO_REQUISITANTE_AUTOMATICO = 12362  'Sem par�metros
'N�o � poss�vel excluir ou alterar o Requisitante Autom�tico.
Public Const ERRO_LEITURA_ITEMNFITEMRC1 = 12363 'Sem par�metros
'Erro na leitura da tabela ItemNFItemRC.
Public Const ERRO_EXCLUSAO_ITEMNFITEMRC = 12364 'Par�metros: lItemNF
'Erro na exclus�o de ItemNFItemRC com item de nota fiscal de n�mero interno %l.
Public Const ERRO_LEITURA_ITENSREQCOMPRASTODOS = 12365 'Sem par�metros
'Erro na leitura da tabela ItensReqComprasTodos.
Public Const ERRO_LEITURA_ITEMPEDCOMPRASTODOS = 12366 'Sem par�metros
'Erro na leitura da tabela ItemPedComprasTodos.
Public Const ERRO_PEDIDOCOMPRABAIXADO_NAO_CADASTRADO = 12367 'Par�metros: lNumIntDoc
'O Pedido de Compras baixado de n�mero interno %l n�o est� cadastrado no Banco de dados.
Public Const ERRO_EXCLUSAO_ITEMPEDCOMPRABAIXADO = 12368 'Par�metros: lNumIntItemPC
'Erro na tentativa de excluir o ItemPC Baixado de n�mero interno %l.
Public Const ERRO_ITENSREQCOMPRA_NAO_CADASTRADO2 = 12369 'Par�metros: lNumIntItem
'O Item com n�mero interno %l de Requisi��o de Compras n�o est� cadastrado no Banco de dados.
Public Const ERRO_LEITURA_REQUISICAOCOMPRABAIXADA = 12370 'Sem par�metros
'Erro na leitura da tabela RequisicaoCompraBaixada.
Public Const ERRO_EXCLUSAO_ITEMREQCOMPRABAIXADO = 12371 'Par�metros: lNumIntItem
'Erro na tentativa de excluir o Item de n�mero interno %l de Requisi��o de Compras Baixadas.
Public Const ERRO_EXCLUSAO_REQUISICAOCOMPRABAIXADA = 12372 'Par�metros: lNumIntDoc
'Erro na tentativa de excluir a Requisi��o de Compras Baixada de n�mero interno %l.
Public Const ERRO_EXCLUSAO_PEDIDOCOMPRABAIXADO = 12373 'Par�metros: lNumIntDoc
'Erro na tentativa de excluir o Pedido de Compras Baixado de n�mero interno %l.
Public Const ERRO_EXCLUSAO_ITEMNFITEMPC = 12374 'Par�metros: lNumItemNF
'Erro na tentativa de excluir registros na tabela ItemNFItemPC.
Public Const ERRO_TIPO_NOTA_FISCAL_DIFERENTE_COMPRAS = 12375 'Par�metros: iTipo
'O Tipo de documento %i n�o � de compras.
Public Const ERRO_ALIQUOTAIPINF_DIFERENTE_PC = 12376 'Par�metros: sProduto, dAliquotaIPINF, dAliquotaIPIPC
'O Produto %s com a al�quota IPI %d � diferente do valor da al�quota IPI no Item de Pedido de Compras que possui al�quotaIPI %d.
Public Const ERRO_ALIQUOTAICMSNF_DIFERENTE_PC = 12377 'Par�metros: sProduto, dAliquotaICMSNF, dAliquotaICMSPC
'O Produto %s com a al�quota ICMS %d � diferente do valor da al�quota ICMS no Item de Pedido de Compras que possui al�quotaICMS %d.
Public Const ERRO_ITEM_PC_PRECOUNITARIO_DIFERENTE = 12378 'Par�metros: sProduto
'O Produto %s de Pedido de Compra est� com pre�o unit�rio diferente
'dos outros pedidos j� selecionados.
Public Const ERRO_ITEM_PC_ALIQUOTAICMS_DIFERENTE = 12379 'Par�metros: sProduto
'O Produto %s de Pedido de Compra est� com Al�quota ICMS diferente
'dos outros pedidos j� selecionados.
Public Const ERRO_ITEM_PC_ALIQUOTAIPI_DIFERENTE = 12380 'Par�metros: sProduto
'O Produto %s de Pedido de Compra est� com Al�quota IPI diferente
'dos outros pedidos j� selecionados.
Public Const ERRO_EXCLUSAO_ITENSREQCOMPRA = 12381 'Par�metros: lNumIntItem
'Erro na exclus�o do Item de Requisi��o de Compras de n�mero interno %l.
Public Const ERRO_LEITURA_ITENSPEDCOMPRABAIXADOS = 12382 'Sem par�metros
'Erro na leitura da tabela ItensPedCompraBaixados.
Public Const ERRO_QUANTIDADE_DIFERENT_QUANTRECEBIDA = 12383 'Par�metros: sProduto
'A soma das quantidades recebidas do produto %s no Grid de intens de Pedido de Compras
'� diferente da Quantidade informada.
Public Const ERRO_QUANTIDADE_MAIOR_TOTALRECEBER = 12384 'Par�metros: dQuantidade, dQuantTotal
'A quantidade %d n�o pode ser maior que a soma das quantidades a receber de todos os
'itens de pedidos de compra que � %d.
Public Const ERRO_PEDIDOCOMPRA_BAIXADO = 12385 'Par�metros: lCodigo
'O Pedido de Compras  de c�digo %l est� baixado.
Public Const ERRO_QUANTRECEBIDA_MAIOR_QUANTRECEBER = 12386 'Sem par�metros
'A quantidade recebida � maior que a quantidade a receber.
Public Const ERRO_QUANTRECEBIDARC_MAIOR_QUANTRECEBIDAPC = 12387 'Sem par�metros
'A quantidade recebida � maior que a quantidade recebida do item de Pedido de Compras.
Public Const ERRO_QUANTRECEBIDARC_MAIOR_QUANTRECEBERRC = 12388 'Sem par�metros
'A quantidade recebida do item Requisi��o � maior que a quantidade a receber do item da Requisi��o.
Public Const ERRO_ITEMNFITEMPC_NAO_CADASTRADO = 12389 'Par�metros: lNumIntItemNF, lNumIntItemPedCompra
'ItemNFItemPC com ItemNF = %l e Item de Pedido de Compras = %l n�o est� cadastrado no Banco de dados.
Public Const ERRO_LOCK_ITEMNFITEMPC = 12390 'Par�metros: lNumIntItemNF, lNumIntItemPedCompra
'Erro no "lock" da tabela ItemNFItemPC com ItemNF = %l e Item de Pedido de Compras = %l.
Public Const ERRO_LEITURA_ITEMNFITEMRC = 12391 'Par�metros: lNumIntItemNF, lNumIntItemReqCompra
'Erro na leitura da tabela ItemNFItemRC com ItemNF = %l e Item de Requisi��o de Compras = %l.
Public Const ERRO_ITEMNFITEMRC_NAO_CADASTRADO = 12392 'Par�metros: lNumIntItemNF, lNumIntItemReqCompra
'ItemNFItemRC com ItemNF = %l e Item de Requisi��o de Compras = %l n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LOCK_ITEMNFITEMRC = 12393 'Par�metros: lNumIntItemNF, lNumIntItemReqCompra
'Erro na "lock" da tabela ItemNFItemRC com ItemNF = %l e Item de Requisi��o de Compras = %l.
Public Const ERRO_INSERCAO_ITEMNFITEMPC = 12394 'Par�metros: lItemNF, lItemPedCompra
'Erro na tentativa de inserir registros em ItemNFItemPC com ItemNF = %l e Item Pedido de Compras = %l.
Public Const ERRO_INSERCAO_ITEMNFITEMRC = 12395 'Par�metros: lItemNF, lItemReqCompra
'Erro na tentativa de inserir registros em ItemNFItemRC com ItemNF = %l e Item Requisi��o de Compras = %l.
Public Const ERRO_ITEMREQCOMPRA_INEXISTENTE = 12396 'Par�metros: lItemReqCompra
'O Item de Requisi��o de Compras de n�mero interno %l n�o est� cadastrado no Banco de dados.
Public Const ERRO_AUSENCIA_PEDIDOCOMPRAS = 12397 'Par�metros: sFornecedor, iFilial, iFilialCompra
'N�o existem Pedidos de Compras para o Fornecedor %s, Filial %i e FilialCompra %i.
Public Const ERRO_INSERCAO_PEDIDOCOMPRABAIXADO = 12398 'Par�metros: lCodigo
'Erro na inser��o do Pedido de Compras de c�digo %l na tabela PedidoCompraBaixado.
Public Const ERRO_LEITURA_ITEMNFITEMPC = 12399 'Sem par�metros
'Erro na leitura da tabela ItemNFItemPC.
Public Const ERRO_REQUISICAOCOMPRA_NAO_ENCONTRADA = 12400
'N�o existe Requisi��o de Compra de acordo com a sele��o informada
Public Const ERRO_BLOQUEIO_ALCADA_EXISTENTE = 12401 'sem parametros
'A op��o de Controle de Al�ada n�o pode ser desmarcada pois existem
'bloqueios de al�ada n�o liberados.
Public Const ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA2 = 12402
'Requisi��o de Compras n�o cadastrada.
Public Const ERRO_ITEMCOT_NAO_VINCULADO_ITEMCONC = 12403 'Par�metros: iLinha
'O Item de Cota��o da linha %i do Grid de Cota��es n�o est� vinculado
'a nenhum Item do Grid de Produtos.
Public Const ERRO_ITEMRC_NAO_VINCULADO_ITEMCONC = 12404 'Par�metros: iLinha
'O Item da linha %i do Grid de Produtos n�o est� vinculado com nenhum
'Item marcado no Grid de Itens de Requisi��o.
Public Const ERRO_GRID_PRECOUNITARIO_NAO_PREENCHIDO = 12405 'Par�metros: iLinha
'O Pre�o unit�rio do item da linha %i do Grid de Cota��es n�o foi preenchido.
Public Const ERRO_QUANTCOTACAO_MAIOR_QUANTREQUISITADA = 12406 'Par�metros: iLinha, dQuantFaltaCotar
'A quantidade a ser cotada da linha %i do Grid Itens de Requisi��o
'n�o pode ser maior que a quantidade que falta ser cotada %d.
Public Const ERRO_ITEMREQCOMPRA_NAO_CADASTRADO2 = 12407 'Par�metros: sProduto, lCodReq
'O Item com o Produto %s de Requisi��o de Compras de c�digo %l n�o est� cadastrado
'no Banco de dados.
Public Const ERRO_ITEM_NAO_VINCULADO_ITEMCOTACAO = 12408 'Par�metros: iLinha
'O item da linha %s do Grid de Produtos n�o est� vinculado a nenhum
'item do Grid de Cota��es.
Public Const ERRO_LEITURA_COTACAOITEMCONCORRENCIABAIXADA = 12409 'Sem parametros
'Erro na leitura da tabela Cotacao Item Concorr�ncia Baixada.
Public Const ERRO_LEITURA_CONCORRENCIABAIXADA = 12410 'Sem parametros
'Erro na leitura da tabela Concorr�ncia Baixada.
Public Const ERRO_LEITURA_ITENSCONCORRENCIABAIXADAS = 12411 'Sem parametros
'Erro na leitura da tabela Itens Concorr�ncia Baixados.
Public Const ERRO_COTACAOPRODUTOITEMRC_NAO_CADASTRADA = 12412 'Par�metros: lItemReqCompra
'A CotacaoProdutoItemRC com ItemRC de n�mero interno %l n�o est� cadastrado no Banco de dados.
Public Const ERRO_LEITURA_PEDIDOCOTACAOBAIXADO = 12413 'Sem par�metros
'Erro na leitura da tabela de Pedidos Cota��o Baixados.
Public Const ERRO_FILIAL_INICIAL_MAIOR = 12414
'A Filial inicial � maior que a final.
Public Const ERRO_COMPRADOR_INICIAL_MAIOR = 12415
'O Comprador inicial � maior que o final.
Public Const ERRO_USUARIO_NAO_COMPRADOR2 = 12416 'Parametro sCodUsuario
'O usu�rio %s n�o � um comprador.
Public Const ERRO_DATAENVIO_INICIAL_MAIOR = 12417
'A Data de Envio inicial � maior que a final.
Public Const ERRO_DATALIMITE_INICIAL_MAIOR = 12418
'A Data Limite inicial � maior que a final.
Public Const ERRO_PC_INICIAL_MAIOR = 12419
'O c�digo do Pedido de Compra inicial � maior que o final.
Public Const ERRO_DESCRICAO_INICIAL_MAIOR = 12420 'SEM PARAMETROS
'A Descri��o inicial � maior que a final.
Public Const ERRO_NUMNF_INICIAL_MAIOR = 12421 'sem parametros
'O N�mero da Nota Fiscal inicial � maior que o final.
Public Const ERRO_SERIE_INICIAL_MAIOR = 12422 'sem parametros
'A S�rie inicial � maior que a final.
Public Const ERRO_CODIGO_OP_INICIAL_MAIOR = 12423
'C�digo da Ordem de Produ��o inicial � maior que o final.
Public Const ERRO_PEDCOTACAO_INICIAL_MAIOR = 12424 'SEm parametros
'O Pedido de Cota��o inicial � maior que o final.
Public Const ERRO_PV_INICIAL_MAIOR = 12425
'Pedido de Venda inicial � maior que o final.
Public Const ERRO_NOMECLIENTE_INICIAL_MAIOR = 12426
'Nome do Cliente inicial � maior que o final.
Public Const ERRO_REQUISITANTE_INEXISTENTE = 12427
'O Requisitante informado n�o existe.
Public Const ERRO_CCL_INEXISTENTE = 12428
'O Centro de Custo informado n�o existe.
Public Const ERRO_NATUREZA_INICIAL_MAIOR = 12429
'A Natureza inicial � maior que o final.
Public Const ERRO_CODIGO_TIPO_PRODUTO_NAO_PREENCHIDO = 12430 'sem parametros
'Preenchimento do tipo de produto � obrigat�rio.
Public Const ERRO_TIPOPRODUTO_NAO_SELECIONADO = 12431 'Sem parametros
'Pelo menos um Tipo de Produto deve ser selecionado.
Public Const ERRO_ITENS_MESMO_LEQUE = 12432 'Par�metros: iItem, iItemCOmparado
'O item %s � igual ao item %s. Eles deve se tornar um �nico item.
Public Const ERRO_REQCOMPRAS_IMPRESSAO = 12433 'Sem parametros
'Selecione uma Requisi��o de Compra para executar a impress�o.
Public Const ERRO_PEDCOMPRA_IMPRESSAO = 12434 'Sem parametros
'Selecione um Pedido de Compra para executar a impress�o.
Public Const ERRO_PRODUTO_JA_EXISTENTE_PEDCOMPRA = 12435 'sProduto, iItem
'O produto %s j� participa deste Pedido de Compra no Item %i.
Public Const ERRO_PEDIDOCOMPRA_BLOQUEADO = 12436 'sParametro: lCodigo
'O Pedido de Compra com c�digo %l � bloqueado.
Public Const ERRO_PRECO_ITEM_NAO_PREENCHIDO = 12437 'sParametro: iItem
'O Pre�o Unit�rio da linha %i do Grid de Itens n�o est� preenchido.
Public Const ERRO_NUMERO_GERACAO_NAO_PREENCHIDO = 12438
'O n�mero da gera��o deve ser preenchido.
Public Const ERRO_PEDCOTACAO_IMPRESSAO = 12439 'Sem parametros
'Selecione um Pedido de Cota��o para executar a impress�o.
Public Const ERRO_DATALIMITE_MAIOR_DATAENVIO = 12440 'Par�metros: dtDataLimite, dtDataEnvio
'A data limite %s deve ser maior ou igual a data de envio, que � %s.
Public Const ERRO_DATALIMITE_INFERIOR_DATAREQ = 12441 'Sem Par�metros.
'A Data limite n�o pode ser menor que a data da requisi��o.
Public Const ERRO_PRODUTO_LEQUE_GRID = 12442 'Par�metros: sProduto
'O Produto %s s� poder� se repetir no grid se o Fornecedor e Filial informados forem diferentes dos j� informados.
Public Const ERRO_CODIGO_CONCORRENCIA_NAO_PREENCHIDO = 12443 'Sem Par�metros
'O c�digo da concorr�ncia deve ser preenchido
Public Const ERRO_MOTIVO_EXCLUSIVO = 12444 'Sem Par�metros.
'O motivo "Exclusividade" n�o pode ser escolhido pelo usu�rio.
Public Const ERRO_PRECOUNITARIO_ITEMCOTACAO_NAO_PREENCHIDO = 12445 'Par�metro : ItemConcorrencia
'Uma cota��o escolhida para o item %i do grid de produtos n�o est� com o pre�o unit�rio preenchido.
Public Const ERRO_LOCALENTREGA_DIFERENTE_FILIALEMPRESA = 12446 'Sem parametros
'O local de entrega � diferente de FilialEmpresa.
Public Const ERRO_QUANT_COTAR_MAIOR_QUANTIDADE_ITEMPEDCOTACAO = 12447 'Sem parametros
'N�o � poss�vel gerar Pedido de Compra. A Quantidade do Item do Pedido de Cota��o � diferente da Quantidade a Cotar dos Itens de Requisi��o.
Public Const ERRO_PRECOPRAZO_MENOR_PRECOVISTA = 12448 'sem parametros
'O Pre�o a Prazo n�o pode ser menor que o Pre�o � Vista.
Public Const ERRO_ATUALIZACAO_COTACAOPRODUTOITEMRC = 12449 'Sem parametros
'Erro na atualiza��o da tabela Cota��oProdutoItemRC
Public Const ERRO_DADOS_DESTINO_NAO_PREENCHIDOS = 12450 'Sem Par�metros
'O itens de cota��o dependem do preenchimento do Tipo de Destino.



'??? jones 28/10

Public Const ERRO_QUANTPRODUTO_MENOR_QUANTREQUISITADA = 12451 'Parametro: iLinha, dQuantCotarItens
'Quantidade a cotar do Produto da linha %i � menor que a quantidade a cotar de Requisi��es %d.
Public Const ERRO_COTACAOPRODUTO_NAO_ENCONTRADA = 12452 'sem parametros
'Cota��o Produto n�o foi encontrada.
Public Const ERRO_PEDIDOCOTACAO_STATUS_NAOATUALIZADO_PC = 12453 'sem parametros
'O Pedido de Cota��o n�o est� atualizado. N�o � poss�vel gerar Pedidos de Compra.
Public Const ERRO_NENHUM_MES_FECHADO = 12454 'Sem par�metros
'N�o existe nenhum m�s fechado para o c�lculo de Par�metros de Ponto de Pedido.







'C�digos de Avisos - Reservado de 15200 a 15299
Public Const AVISO_EXCLUSAO_ALCADA = 15200 'Par�metro: sCodUsuario
'Confirma a exclus�o da al�ada do usu�rio com c�digo %s?
Public Const AVISO_CONFIRMA_EXCLUSAO_COMPRADOR = 15201  'Parametro: sCodUsuario
'Confirma a exclus�o do comprador %s ?
Public Const AVISO_EXCLUIR_REQUISITANTE = 15202 ' Parametro: lCodigo
'Confirma exclus�o do Requisitante com c�digo %l?
Public Const AVISO_EXCLUIR_TIPODEBLOQUEIOPC = 15203 ' Parametro :iCodigo
'Confirma exclus�o do Tipo de Bloqueio de Pedido de Compra com codigo %i ?
Public Const AVISO_NUM_MAX_BLOQUEIOSPC_LIBERACAO = 15204
'O n�mero m�ximo de Bloqueios poss�veis de Pedido de Compras para exibi��o foi atingido . Ainda existem mais Bloqueios para Libera��o .
Public Const AVISO_CRIAR_REQUISITANTE = 15205 'Parametro: lCodigo
'Requisitante com c�digo %l n�o est� cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_REQUISITANTE1 = 15206 'Parametro: sNomeReduzido
'Requisitante com Nome Reduzido %s n�o est� cadastrado. Deseja cadastrar?
Public Const AVISO_EXCLUSAO_PEDIDOCOTACAO = 15207 'Parametros lCodigo
'Confirma a exclus�o do Pedido de Cota��o com o c�digo %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_PEDIDO_COMPRA = 15208 'PArametro lCodigo
'Confirma exclus�o do Pedido de Compra com o c�digo %l?
Public Const AVISO_BAIXA_PEDIDOCOMPRAS = 15209 'Parametro lCodigo
'Deseja realmente baixar o pedido de compras %l ?
Public Const AVISO_EXCLUIR_REQUISICAOCOMPRA = 15210 'Par�metros: lCodigo
'Confirma a exclus�o da Requisi��o de Compras de c�digo %l?
Public Const AVISO_EXCLUIR_REQUISICAOMODELO = 15211 'Par�metros: lCodReqModelo
'Confirma a exclus�o da Requisi��o Modelo de c�digo %l?
Public Const AVISO_REQUISICAOCOMPRA_GERADA = 15213  'Par�metros: lCodigo
'Foi gerada a Requisi��o de Compras de c�digo %l.
Public Const AVISO_EXCLUIR_CONCORRENCIA = 15214 'Par�metros: lCodigo
'Confirma exclus�o da Concorr�ncia de c�digo %l?
Public Const AVISO_RECEBMATERIALFCOM_MESMO_NUMERO = 15215 ' Parametros lFornecedor, iFilialForn, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEntrada
'No Banco de Dados existe Recebimento de Material com os dados C�digo do Fornecedor %l, C�digo da Filial %i, Tipo %i, S�rie NF %s, N�mero NF %l, DataEntrada %dt. Deseja prosseguir na inser��o de novo Recebimento de Material com o mesmo n�mero de Nota Fiscal?
Public Const AVISO_VALORFRETE_DIFERENTE_PC = 15216 'Par�metros: dValorFreteNF, dValorFretePC
'O Valor Frete %d da Nota Fiscal � diferente do Valor do Pedido de Compras que � %d. Deseja continuar com a grava��o?
Public Const AVISO_VALORSEGURO_DIFERENTE_PC = 15217 'Par�metros: dValorSeguroNF, dValorSeguroPC
'O Valor Seguro %d da Nota Fiscal � diferente do Valor do Pedido de Compras que � %d. Deseja continuar com a grava��o?
Public Const AVISO_VALORDESCONTO_DIFERENTE_PC = 15218 'Par�metros: dValorDescontoNF, dValorDescontoPC
'O Valor Desconto %d da Nota Fiscal � diferente do Valor do Pedido de Compras que � %d. Deseja continuar com a grava��o?
Public Const AVISO_VALORDESPESAS_DIFERENTE_PC = 15219 'Par�metros: dValorDespesasNF, dValorDespesasPC
'O Valor de Despesas %d da Nota Fiscal � diferente do Valor do Pedido de Compras que � %d. Deseja continuar com a grava��o?
Public Const AVISO_VALORUNITARIO_DIFERENTE_PC = 15220 'Par�metros: dValorUnitarioItemNF, iLinha, dValorUnitarioItemPC
'O Valor Unit�rio %d da linha %i do Grid de Itens de Notas Fiscais � diferente do Valor Unit�rio do Item de Pedido de Compras que � %d.
'Deseja continuar com a grava��o?
Public Const AVISO_DESCONTOITEM_DIFERENTE_PC = 15221 'Par�metros: dValorDescontoItemNF, iLinha, dValorDescontoItemPC
'O Valor Desconto %d da linha %i do Grid de Itens de Notas Fiscais � diferente da soma dos Valores de Desconto dos Itens de Pedido de Compras que � %d.
'Deseja continuar com a grava��o?
Public Const AVISO_ALIQUOTAIPI_DIFERENTE_PC = 15222 'Par�metros: sProduto, dAliquotaIPINF, dAliquotaIPIPC
'O Produto %s com a al�quota IPI %d � diferente do valor no Item de Pedido de Compras que possui al�quotaIPI %d. Deseja Continuar?
Public Const AVISO_ALIQUOTAICMS_DIFERENTE_PC = 15223 'Par�metros: sProduto, dAliquotaICMSNF, dAliquotaICMSPC
'O Produto %s com a al�quota ICMS %d � diferente do valor no Item de Pedido de Compras que possui al�quotaICMS %d. Deseja Continuar?
Public Const AVISO_BAIXA_REQCOMPRAS = 15224 'Par�metro: lCodReqCompra
'Deseja realmente baixar a requisi��o de compra %l ?
Public Const AVISO_LOTEMINIMO_MAIOR_QUANTCOMPRAR = 15225 'sProduto, dQuantidade, dLoteMinimo
'O Produto %s possui quantidade %d menor que o lote minimo relacionado ao Fornecedor Filial Produto que � %d. Deseja prosseguir com a grava��o?
Public Const AVISO_QUANTIDADE_PEDIDA_ESTOQUEMAXIMO = 15226  'Parametro: sProduto
'A soma da quantidade pedida mais a quantidade dispon�vel � maior que a quantidade de estoque m�ximo do produto %s. Confirma a quantidade informada?
Public Const AVISO_AUSENCIA_COTACOES_SELECAO = 15227 'Sem Par�metros.
'N�o existem cota��es com dados compat�veis para atender a esse item.
Public Const AVISO_CONCORRENCIA_GRAVADA = 15228 'Par�metros: lCodigo
'A concorr�ncia %s foi gravada com sucesso.
Public Const AVISO_PEDIDOCOMPRA_GERADO = 15229
'O Pedido de Compra foi gerado.
Public Const AVISO_CANCELAR_CALCULO_PARAMETROS_PTOPEDIDO = 15230 'Sem parametros
'Confirma o canselamento do Calculo dos Parametros para Ponto Pedido ?
Public Const AVISO_ITEMRC_PARTICIPA_CONCORRENCIA = 15231
'Existem concorr�ncia(s) anteriores envolvendo as requisi��es utilizadas nessa concorr�ncia. Deseja prosseguir com essa grava��o e excluir as concor�ncias anteriores?


