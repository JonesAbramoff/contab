Attribute VB_Name = "ErrosCOM"
Option Explicit

'Códigos de Erros - Reservado de 12000 a 12999
Public Const ERRO_PERCENT_MAIS_QUANTCOTACAO_ANTERIOR_NAO_PREENCHIDA = 12000
'O campo Percentagem a mais de Cotações Anteriores não foi preenchido.
Public Const ERRO_PERCENT_MENOS_QUANTCOTACAO_ANTERIOR_NAO_PREENCHIDA = 12001
'O campo Percentagem a menos de Cotações Anteriores não foi preenchido.
Public Const ERRO_PERCENT_MAIS_RECEB_NAO_PREENCHIDA = 12002
'O campo Percentagem a mais da Faixa de Recebimento não foi preenchido.
Public Const ERRO_PERCENT_MENOS_RECEB_NAO_PREENCHIDA = 12003
'O campo Percentagem a menos da Faixa de Recebimento não foi preenchido.
Public Const ERRO_LEITURA_ALCADA = 12009 'Parâmetro: sCodUsuario
'Erro na busca da alçada do usuário %s na tabela de Alçadas.
Public Const ERRO_LOCK_ALCADA = 12010 'Parâmetro: sCodUsuario
'Não conseguiu fazer o lock da alçada do usuário %s.
Public Const ERRO_LEITURA_PEDIDOCOMPRABAIXADO = 12012 'Sem parametro
'Erro na leitura da tabela de Pedido de Compra Baixado
Public Const ERRO_LEITURA_VALORPCLIBERADO = 12013 'Sem parametro
'Erro na leitura de Valor de Pedido de Compra Liberado
Public Const ERRO_ALCADA_VINCULADA_PEDIDOCOMPRA = 12014 'Parametro sCodUsuario
'A alçada não pode ser excluída pois o usuário %s está vinculado a um Pedido de Compra
Public Const ERRO_ALCADA_VINCULADA_PEDIDOCOMPRABAIXADO = 12015 'Parametro sCodUsuario
'A alçada não pode ser excluída pois o usuário %s está vinculado a um Pedido de Compra Baixado
Public Const ERRO_ALCADA_VINCULADA_VALORPCLIBERADO = 12016 'Parametros:sCodUsuario, iAno
'A alçada não pode ser excluída pois o usuário %s está vinculado a um registro
'na tabela de Valor de Pedido de Compra Liberado
Public Const ERRO_EXCLUSAO_ALCADA = 12017 'Parâmetro: sCodUsuario
'Erro na tentativa de exclusão da alçada do usuário %s da tabela de Alçadas.
Public Const ERRO_ATUALIZACAO_ALCADA = 12018 'Parâmetro: sCodUsuario
'Erro na tentativa de atualizar a alçada do usuário %s da tabela de Alçadas.
Public Const ERRO_INSERCAO_ALCADA = 12019 'Parametro: sCodUsuario
'Erro na tentiva de inserir alçada do usuario %s.
Public Const ERRO_LEITURA_REQUISITANTE = 12021 'Parametro:  lCodigo
'Erro na leitura do Requisitante %l na tabela de Requisitantes.
Public Const ERRO_LEITURA_REQUISITANTE1 = 12022 'Parâmetro: sNomeReduzido
'Erro na leitura do Requisitante %s na tabela de Requisitantes.
Public Const ERRO_LEITURA_REQUISICAOCOMPRAS = 12023 'Sem Parâmetros
'Erro na leitura da tabela de Requisição de Compras.
Public Const ERRO_LEITURA_ITENSCONCORRENCIA = 12025 'ParÂmetro: lCodConcorrencia
'Erro na leitura dos itens de concorrência da concorrência com o código %l da tabela de itens de concorrência.
Public Const ERRO_REQCOMPRA_VINCULADA_CONCORRENCIA_NAO_ENCONTRADA = 12027
'Uma requisição de compras vinculada a concorrência não foi encontrada.
Public Const ERRO_USUARIO_SEM_ALCADA = 12028 'parametro sCodUsuario
'O usuário %s não possui alçada.
Public Const ERRO_LOCK_VALORPCLIBERADO = 12029 ' Parametro sCodUsuario
'Erro na tentativa de fazer lock no Valor de Pedido de Compra do usuário %s.
Public Const ERRO_INSERCAO_VALORPCLIBERADO = 12030 'parametro sCodUsuario
'Erro na tentativa de inserir Valor de Pedido de Compra Liberado do usuário %s
Public Const ERRO_ATUALIZACAO_VALORPCLIBERADO = 12031 'Parametro sCodUsuario
'Erro na atualização do Valor de Pedido de Compra Liberado do usuário %s
Public Const ERRO_LIMITE_MENSAL_MENOR = 12032 'parametro dLimiteMensal
'O limite mensal %d é menor que o valor do pedido.
Public Const ERRO_LIMITE_OPERACAO_MENOR = 12033 'parametro dLimiteOperacao
'O limite de operação %d é menor que  o valor do pedido.
Public Const ERRO_AUSENCIA_BLOQUEIOS_LIBERAR = 12034 'Sem Parametros
'Não há bloqueios no Grid selecionados para a liberação
Public Const ERRO_NUM_BLOQUEIOS_SELECIONADOS_SUPERIOR_MAXIMO = 12035  'Sem parametros
'O número de bloqueios selecionados é superior ao número máximo possível para liberação. Restrinja mais a seleção para que o número de bloqueios lidos diminua.
Public Const ERRO_SEM_BLOQUEIOS_PC_SEL = 12036 'Sem parametros
'Não há bloqueios dentro dos critérios de seleção informados.
Public Const ERRO_LOCK_BLOQUEIOSPC = 12037 'Parametro: lPedCompras
'Não conseguiu fazer lock no Bloqueio de Pedido de Compras %l.
Public Const ERRO_ATUALIZACAO_BLOQUEIOSPC = 12038 'Parametro lPedCompras
'Erro na atualização do Bloqueio de Pedido de Compras %l.
Public Const ERRO_LEITURA_BLOQUEIOSPC1 = 12039 'Parametro: lPedCompras
'Erro na leitura do Bloqueio de Pedido de Compras %l.
Public Const ERRO_TIPOBLOQUEIOPC_NAO_MARCADO = 12040 'Sem parametros
'Pelo menos 1 Tipo de Bloqueio deve estar marcado para esta operação.
Public Const ERRO_ATUALIZACAO_PEDIDOCOMPRA = 12041 'Parametro lPedCompra
'Erro na atualização do Pedido de Compra com o código %l.
Public Const ERRO_BLOQUEIOPC_INEXISTENTE = 12042 'Parametro lPedCompra
'O Bloqueio de Pedido de Compras %l não existe.
Public Const ERRO_LOCK_PEDIDOCOMPRA = 12043 'Parametro lPedCompra
'Não conseguiu fazer o lock do Pedido de Compra com o código %l.
Public Const ERRO_LIMITE_MENSAL_ULTRAPASSADO = 12044 'Parametro dLimiteMensal, sCodUsuario
'O limite mensal %d do usuário %s foi ultrapassado.
Public Const ERRO_COMPRADOR_USUARIO = 12045 'Parametros: iCodComprador,sCodUsuario
'O código de Comprador %i correspondente ao Usuário %s no Bando de Dados não confere com o Comprador %i da Tela.
Public Const ERRO_COMPRADOR_NAO_CADASTRADO = 12046  'Parametro: iCodigo
'O comprador com o codigo %i nao esta cadastrado
Public Const USUARIO_COMPRADOR_NAO_ALTERAVEL = 12047 'Parametro: iCodigo
'O comprador com o codigo %i nao esta cadastrado
Public Const ERRO_LEITURA_COMPRADOR2 = 12048
'Erro de leitura na tabela de compradores.
Public Const ERRO_LEITURA_COMPRADOR = 12049 'Parametro: iCodigo
'Erro de leitura do Comprador com o código %i na tabela de compradores.
Public Const ERRO_LEITURA_COMPRADOR1 = 12050 'Parametro: sCodUsuario
'Erro de leitura do comprador com o código de usuário %s.
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
'O nome %s já está sendo utilizado por outro requisitante.
Public Const ERRO_LEITURA_REQUISICAOCOMPRABAIXADA1 = 12063 'Parametro: lCodRequisitante
'Erro na busca do Requisitante %l na tabela de Requisição de Compra Baixada.
Public Const ERRO_LEITURA_REQUISICAOMODELO1 = 12064 'Parametro lCodRequisitante
'Erro na busca do Requisitante %l na tabela de Requisição Modelo.
Public Const ERRO_REQUISITANTE_VINCULADO_REQCOMPRABAIXADA = 12065 'Parametro: lCodigo
'O Requisitante com código %l não pode ser excluído pois está relacionado a uma requisição de compra baixada.
Public Const ERRO_REQUISITANTE_VINCULADO_REQCOMPRAMODELO = 12066 'Parametro: lCodigo
'O Requisitante com código %l não pode ser excluído pois está relacionado a uma requisição de compra modelo.
Public Const ERRO_REQUISITANTE_NAO_CADASTRADO = 12067 'Parametro: lCodigo
'O Requisitante com código %l não está cadastrado.
Public Const ERRO_LEITURA_REQUISICAOCOMPRA1 = 12068 'Parametro: lCodRequisitante
'Erro na busca pelo Requisitante %l na tabela de RequisiçãoCompra
Public Const ERRO_LOCK_REQUISITANTE = 12069 'Parametro: lCodigo
'Não conseguiu fazer o lock do requisitante %l.
Public Const ERRO_EXCLUSAO_REQUISITANTE = 12070  'parametro: lCodigo
'Erro na exclusão do requisitante %l na tabela de Requisitantes.
Public Const ERRO_REQUISITANTE_NOMERED_DUPLICADO = 12071 'Parametro: sNomeReduzido
'O nome reduzido %s já está sendo utilizado por outro requisitante.
Public Const ERRO_ATUALIZACAO_REQUISITANTE = 12072 'Parâmetro : lCodigo
'Erro na atualização do Requisitante %l na tabela de Requisitantes.
Public Const ERRO_INSERCAO_REQUISITANTE = 12073 'Parâmetro: lCodigo
'Erro na Inserção do Requisitante %l na tabela de Requisitantes.
Public Const ERRO_REQUISITANTE_VINCULADO_REQUISICAOCOMPRA = 12074 'parametro: lCodigo
'O Requisitante %l não pode ser excluído pois está sendo utilizado em uma Requisição de Compra.
Public Const ERRO_EXCLUSAO_TIPO_BLOQUEIO_ALCADA = 12075  'Sem parametros
'Não é possível excluir o Tipo de Bloqueio de Alçada.
Public Const ERRO_ALTERACAO_TIPO_BLOQUEIO_ALCADA = 12076 'Sem parametro
'Não é possível alterar o Tipo de Bloqueio de Alçada.
Public Const ERRO_ATUALIZACAO_TIPODEBLOQUEIOPC = 12077 'Parametro: iCodigo
'Erro na atualização do Tipo de Bloqueio %i na tabela de Tipo de Bloqueio de Pedido de Compra.
Public Const ERRO_TIPODEBLOQUEIOPC_MESMA_DESCRICAO = 12078 ' Parametro sDescricao
'Já existe no Banco de Dados Tipo de Bloqueio com a descrição %s
Public Const ERRO_INSERCAO_TIPODEBLOQUEIOPC = 12079 'Parametro iCodigo
'Erro na inserção do Tipo de Bloqueio %i na tabela de Tipos de BloqueioPC
Public Const ERRO_LEITURA_TIPODEBLOQUEIOPC = 12080 'Parametro iCodigo
'Erro na leitura do Tipo de Bloqueio %i na tabela de Tipos de BloqueioPC
Public Const ERRO_TIPODEBLOQUEIOPC_MESMO_NOME = 12081 'Parametro sNomeReduzido
'Já existe no Banco de Dados Tipo de Bloqueio com nome reduzido %s
Public Const ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO = 12082 'Parametro iCodigo
'O Tipo de Bloqueio de Pedido de Compra %i não está cadastrado
Public Const ERRO_LOCK_TIPODEBLOQUEIOPC = 12083 'Parametro iCodigo
'Não conseguiu fazer o lock do Tipo de Bloqueio de Pedido de Compra %i
Public Const ERRO_LEITURA_BLOQUEIOSPC = 12084 'Parametro iTipoBloqueio
'Erro na leitura do Tipo de Bloqueio %i da tabela de BloqueiosPC
Public Const ERRO_EXCLUSAO_TIPODEBLOQUEIOPC = 12085 'Parametro:iCodigo
'Erro na exclusão do Tipo de Bloqueio de Pedido de Compra %i da tabela de TiposDeBloqueioPC
Public Const ERRO_TIPODEBLOQUEIOPC_VINCULADO_BLOQUEIOSPC = 12086 'Parametro iCodigoTipoBloqueioPC
'Não é possível excluir o Tipo de Bloqueio com o código %i pois ele está sendo utilizado por um Bloqueio de Pedido de Compra
Public Const ERRO_MESESMEDIATEMPORESSUP_NAO_PREENCHIDO = 12087 'Sem parametros
'Meses para tempo de ressuprimento não foi preenchido.
Public Const ERRO_MESESCONSUMOMEDIO_NAO_PREENCHIDO = 12088 'Sem parametros
'Meses de consumo médio não foi preenchido.
Public Const ERRO_LEITURA_BLOQUEIOS_PC_LIBERACAO = 12089
'Erro na leitura dos Bloqueios de Pedidos de Compras para liberar.
Public Const ERRO_PEDIDOCOMPRA_NAO_CADASTRADO = 12090 'Parametro lCodigo
'O Pedido de Compra com código %l não está cadastrado.
Public Const ERRO_USUARIO_NAO_COMPRADOR = 12091 'Parametro sCodUsuario
'O usuário %s não tem acesso a essa tela pois só é acessível à compradores
Public Const ERRO_PEDIDOCOTACAO_NAO_SELECIONADO = 12092 'Sem parametros
'Um Pedido de Cotação deve ser selecionado.
Public Const ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO = 12093 'Parametros lCodigo
'O Pedido de Cotação %l não está cadastrado no Banco de Dados.
Public Const ERRO_ITEM_PEDCOTACAO_VINCULADO_ITEM_PEDCOMPRA = 12094 'Parametros lCodigo
'Um item do Pedido de Cotação %l está vinculado à um item de um Pedido de Compra.
Public Const ERRO_ITEM_PEDCOTACAO_VINCULADO_CONCORRENCIA = 12095 'Parametros lCodigo
'Um item do Pedido de Cotação %l está vinculado a uma Concorrência.
Public Const ERRO_PEDIDO_COTACAO_AVULSO = 12096 'Sem parametros
'Cotações desvinculadas de Requisições não geram Pedidos de Compra.
Public Const ERRO_CONDICAO_PAGTO_NAO_PREENCHIDA = 12097 'Sem parametros
'A Condição de Pagamento deve ser informada.
Public Const ERRO_PRECOS_ITENS_CONDPAGTO_NAO_PREENCHIDOS = 12098 'Sem parametros
'Os preços dos itens para a Condição de Pagamento escolhida devem estar preenchidos.
Public Const ERRO_QUANTENTREGA_MAIOR_QUANTCOTACAO = 12099 'Sem parametros
'A quantidade disponível para entrega deve ser menor ou igual a quantidade de cotação.
Public Const ERRO_ITEMPEDCOTACAO_VINCULADO_CONCORRENCIA1 = 12100 'Parametros iCondPagtoPrazo, iIndice1
'O preço unitário para a condição de pagamento %i do item %i não pode ser apagado pois esse item já está relacionado com uma concorrência.
Public Const ERRO_LOCK_PEDIDO_COTACAO = 12107 'Parametro lCodigo.
'Não conseguiu fazer o lock do pedido de cotacao %l .
Public Const ERRO_LEITURA_OBSERVACAO = 12110 'Parametro sObservacao
'Erro na leitura da Observacao %s.
Public Const ERRO_INSERCAO_OBSERVACAO = 12111 'Parametros sObservacao
'Erro na inserção da observação %s.
Public Const ERRO_INSERCAO_ITENSCOTACAO = 12112 'Parametro lCodigo
'Erro na inserção dos itens de cotação do pedido de cotação %l.
Public Const ERRO_ATUALIZACAO_PEDIDOCOTACAO = 12113 'Parametros lCodigo
'Erro na atualização dos dados do pedido de cotação %l
Public Const ERRO_ATUALIZACAO_ITENSCOTACAO = 12114 'Parametros lCodigo
'Erro na atualização dos itens de cotação do pedido de cotação %l
Public Const ERRO_EXCLUSAO_PEDIDO_COTACAO = 12115 'Parametro lCodigo
'Erro na exclusão do pedido de cotação %l.
Public Const ERRO_AUSENCIA_ITENS_PEDIDOCOTACAO = 12118 'Sem parametros
'Não existem itens para o Pedido de Cotação.
Public Const ERRO_LEITURA_COTACAOPRODUTOITEMRC = 12119 'Sem parametros
'Erro na leitura da tabela CotacaoProdutoItemRC
Public Const ERRO_COTACAO_NAO_CADASTRADA = 12120 'Parametro lCotacao
'A cotação %l não está cadastrada.
Public Const ERRO_OBSERVACAO_NAO_CADASTRADA = 12121 'Parametros lObservacao
'A observação %l não está cadastrada.
Public Const ERRO_ITEMPEDCOTACAO_NAO_ENCONTRADO = 12122 'Parametro lCodigo
'Não foram encontrados itens para o pedido de cotação %l.
Public Const ERRO_ITENSCOTACAO_NAO_ENCONTRADOS = 12123 'Parametro lCodigo
'Não foram encontrados itens de cotação para o pedido de cotação %l.
Public Const ERRO_REQUISICAOCOMPRA_NAO_ENVIADA = 12124 'Parâmetros: lCodigo
'A Requisição de Compras de código %l não foi enviada.
Public Const ERRO_LOCK_REQUISICAOCOMPRA = 12126 'Parâmetros: lCodigo
'Erro na tentativa de fazer "lock" na tabela Requisição Compra com Requisição de código %l.
Public Const ERRO_QUANTCANCELAR_SUPERIOR_QUANTDISPONIVEL_CANCELAR = 12127 'Parâmetros: sProduto e dQuantMaximaCancelar
'A quantidade informada do produto %s é inválida. A quantidade máxima a cancelar é %d.
Public Const ERRO_REQUISICAO_NAO_CARREGADA = 12128 'Sem parâmetros
'Nenhuma requisição foi trazida para a tela.
Public Const ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA = 12129 'Parâmetros: lCodigo
'A Requisição de Compra de código %l não está cadastrada no Banco de Dados.
Public Const ERRO_REQUISITANTE_NAO_CADASTRADO1 = 12130 'Parâmetros: sNomeReduzido
'O Requisitante com o Nome Reduzido %s não está cadastrado no Banco de Dados.
Public Const ERRO_BAIXA_ITEMRC_VINCULADO_ITEMPC_NAO_BAIXADO = 12131 'Parâmetros: lCodigo, lNumIntDoc
'A Requisição de código %l não pode ser baixada pois o item de número interno %l está vinculado a um item
'de pedido de venda não baixado.
Public Const ERRO_BAIXA_ITEMRC_VINCULADO_ITEMCONCORRENCIA_NAO_BAIXADO = 12132 'Parâmetros: lCodigo, lItemConcorrencia
'A Requisição de código %l não pode ser baixada pois o Item de Concorrência %l está vinculado a um item de concorrência não baixado.
Public Const ERRO_LEITURA_ITEMRCITEMCONCORRENCIA = 12133 'Parâmetros: lItemReqCompra
'Erro na leitura da tabela Item Requisição Compra Item Concorrência com Item Requisição de Compra %l.
Public Const ERRO_LOCK_ITENSREQCOMPRA = 12134 'Parâmetros: lCodRequisicao
'Erro na tentativa de fazer lock na tabela Itens Requisição Compra
'com Requisição de código %l.
Public Const ERRO_EXCLUSAO_ITENSREQCOMPRAS = 12135 'Parâmetros: lCodRequisicao
'Erro na tentativa de excluir item da Requisição Compra com o código %l
Public Const ERRO_EXCLUSAO_COTACAOPRODUTOITEMRC = 12136 'Parâmetros: lItemReqCompra
'Erro na tentativa de excluir Cotação Produto Item Requisição Compras com Item Requisição de Compra %l.
Public Const ERRO_ATUALIZACAO_ITENSREQCOMPRA = 12137 'Parâmetros: lNumIntDoc
'Erro na atualização de Itens Requisição Compra de número interno %l.
Public Const ERRO_INSERCAO_ITENSREQCOMPRABAIXADOS = 12138 'Parâmetros: lNumIntDoc
'Erro na Inserção do item Requisição de Compra número interno %l na tabela Itens Requisição Compra Baixados.
Public Const ERRO_EXCLUSAO_REQUISICAOCOMPRA = 12139 'Parâmetros: lCOdigo
'Erro na tentativa de excluir a Requisição Compra de código %l da tabela Requisição Compra.
Public Const ERRO_INSERCAO_REQUISICAOCOMPRABAIXADA = 12140 'Parâmetros: lCodigo 'OK ??? passe o código da req como parâmetro
'Erro na tentativa de inserir a Requisição de código %l na tabela Requisição Compra Baixada.
Public Const ERRO_ITEMRC_DESVINCULADO_ITEMPC = 12141 'Parâmetros: lNumIntDocItemPC
'O Item %l de Requisição de Compras não possui nenhum vínculo com Pedidos de Compra.
Public Const ERRO_ITENSREQCOMPRA_NAO_CADASTRADO = 12142 'Parâmetros: lNumIntDocItemRC, lCodRequisicao
'O Item de número interno %l da Requisição de Compras de código %l não está cadastrado no
'Banco de dados.
Public Const ERRO_QUANTCANCELADA_MAIOR = 12143 'Sem parâmetros
'A quantidade a cancelar não pode ser maior que a quantidade a receber.
Public Const ERRO_COMPRADOR_NAO_CADASTRADO1 = 12144  'Parametro: sCodUsuario
'O comprador com o codigo de usuário %s não esta cadastrado
Public Const ERRO_PEDIDOCOMPRA_JA_CADASTRADO = 12145 'lCodigo
'O Pedido de Compra %l já está cadastrado no Banco de Dados.
Public Const ERRO_ATUALIZACAO_ITENSPEDCOMPRA = 12146 'Sem parametro
'Erro na atualização da tabela de Itens de Pedido de Compra.
Public Const ERRO_LEITURA_LOCALIZACAOITENSPC = 12147 'Sem parametro
'Erro na leitura da tabela de LocalizaçãoItensPC.
Public Const ERRO_PEDIDOCOMPRA_GERADO = 12148 'Parametro lCodigo
'O Pedido de Compra %l é um pedido de compra gerado.
Public Const ERRO_LOCK_ITENSPEDCOMPRA = 12149 'Sem parametro
'Erro na tentativa de lock em ItensPedCompra.
Public Const ERRO_EXCLUSAO_ITENSPEDCOMPRA = 12150 'Parametro lCodigo
'Erro na exclusão do Item com número interno %l da tabela de Itens de Pedido de Compra.
Public Const ERRO_INSERCAO_BLOQUEIOPC = 12151 'Parametro lCodigo
'Erro na inserção do Bloqueio de Pedido de Compra %l.
Public Const ERRO_INSERCAO_LOCALIZACAOITENSPC = 12152 'Parametros lNumIntDoc
'Erro na inserção do item com número interno %l na tabela de LocalizaçãoItensPC.
Public Const ERRO_INSERCAO_ITENSPEDCOMPRA = 12153 'Parametro lNumIntDoc
'Erro na inserção do item com número interno %l do Pedido de Compra.
Public Const ERRO_INSERCAO_PEDIDOCOMPRA = 12154 'Parametro lCodigo
'Erro na inserção do Pedido de Compra %l na tabela de PedidoCompra.
Public Const ERRO_OBSERVACAO_INEXISTENTE = 12155 'Parametro lNumInt
'A observação com o número interno %l não está cadastrada no Banco de Dados.
Public Const ERRO_DATAENVIO_INFERIOR_DATAPEDIDO = 12156 'Sem parametro
'A Data de Envio deve ser maior ou igual a Data do Pedido.
Public Const ERRO_PRODUTO_DESVINCULADO_ITEM = 12157 'Parametro sDescricao
'O produto %s não está presente em nenhum dos itens do Pedido de Compras
Public Const ERRO_PRODUTO_ITEM_DISTRIBUICAO_VAZIO = 12158 'Sem parametro
'Não é possível deixar um Item de Distribuição sem Produto
Public Const ERRO_PEDCOMPRA_BAIXADO = 12159 'Parametro lCodigo
'O Pedido de Compra com código %l está baixado.
Public Const ERRO_LOCK_OBSERVACAO = 12160 'Sem Parametros
'Erro na tentativa de lock na tabela de Observação.
Public Const ERRO_ALMOXARIFADO_ITEM_DISTRIBUICAO_VAZIO = 12161 'Sem parametro
'Não é possível deixar um Item de Distribuição sem Almoxarifado
Public Const ERRO_AUSENCIA_ITENS_PC = 12162 'Sem parametros
'Não existem itens para o Pedido de Compra.
Public Const ERRO_DATALIMITE_ITEM_INFERIOR_DATAPEDIDO = 12163 'Parametro iItem 'OK que é iIndice (se eu estiver em outra tela e precisar desse erro não vou entender)
'A Data Limite do item %i é menor que a Data do Pedido
Public Const ERRO_LEITURA_PEDIDOCOMPRA_BAIXADO = 12164 'Parametro lCodigo
'Erro na leitura do Pedido de Compra Baixado %l.
Public Const ERRO_PEDCOMPRA_BAIXADO_EXCLUSAO = 12165 'Parametro lCodigo
'O Pedido de Compra com código %l está baixado. Não pode ser excluído.
Public Const ERRO_DATALIMITE_INFERIOR_DATAPEDIDO = 12166 'Sem parametros
'A Data Limite deve ser maior ou igual a Data do Pedido
Public Const ERRO_ALIQUOTA_IGUAL_100 = 12167 'Sem parametros
'A alíquota deve ter um valor inferior a 100%
Public Const ERRO_VALORTOTAL_PC_NEGATIVO = 12168
'Valor Total do Pedido de Compra é negativo.
Public Const ERRO_PEDIDOCOMPRA_ENVIADO = 12169 'Parametro lcodigo
'O Pedido de Compra com código %l já foi enviado.
Public Const ERRO_ITEMPEDCOMPRA_INEXISTENTE = 12170 'Parametro lNumIntDoc
'Item do Pedido de Compra  com número interno %l não está cadastrado.
Public Const ERRO_PEDCOMPRA_AUSENCIA_ITENS = 12171 'Parametro lCodigo
'Pedido de Compra %l não possui itens.
Public Const ERRO_EXCLUSAO_PEDIDOCOMPRA = 12172 'Parametro lCodigo
'Erro na exclusão do Pedido de Compra %s.
Public Const ERRO_EXCLUSAO_BLOQUEIOSPC = 12173
'Erro na exclusão de Bloqueios de Pedido de Compras.
Public Const ERRO_USUARIO_NAO_PREENCHIDO2 = 12175 'Sem Parametros
'O preenchimento do Nome Reduzido do Usuário é obrigatório.
Public Const ERRO_LIMITE_MENSAL_MENOR_LIMITE_OPERACAO = 12177 'Sem Parâmetros
'O Limite Mensal não deve ser menor que o Limite de Operação.
Public Const ERRO_DATAENVIODE_MAIOR_DATAENVIOATE = 12178  'Sem parametros
'Data de Envio De maior que Data de Envio Até.
Public Const ERRO_NUM_PEDIDOS_SELECIONADOS_SUPERIOR_MAXIMO = 12179 'Sem parametros
'O número de pedidos selecionados é superior ao número máximo possível para
'baixa. Restrinja mais a seleção para que o número de pedidos lidos diminua.
Public Const ERRO_LEITURA_PEDIDOS_COMPRA_BAIXA_PC = 12180 'Sem parametros
'Erro na leitura de Pedidos de Compra para baixa.
Public Const ERRO_REQUISICAOCOMPRA_INEXISTENTE = 12181 'Sem parametros
'Não foi encontrada nenhuma Requisição de Compras de acordo com a seleção informada.
Public Const ERRO_AUSENCIA_REQUISICOES_BAIXAR = 12182 'Sem parametros
'Deve haver pelo menos uma Requisição marcada para ser baixada.
Public Const ERRO_REQUISICAO_INICIAL_MAIOR = 12183 'Sem parametros
'O número da Requisição Inicial não pode ser maior que o da Requisição Final.
Public Const ERRO_REQUISITANTE_INICIAL_MAIOR = 12184 'Sem parametros
'O código do requisitante inicial não pode ser maior que o do requisitante final.
Public Const ERRO_DATALIMITEDE_MAIOR = 12185 'Sem parametros
'A data limite inicial não pode ser maior que a data limite final.
Public Const ERRO_INSERCAO_ITENSPEDCOMPRABAIXADOS = 12186 'Parâmetro tItemPedido.lNumINtDoc
'Erro na inserção do item %l na tabela de Itens de Pedido de Compra Baixados
Public Const ERRO_RESIDUO_NAO_PREENCHIDO = 12187 'Sem parametros
'O preenchimento do Resíduo é obrigatório.
Public Const ERRO_FORNECEDORFILIALPRODUTO_NAO_CADASTRADA = 12188 'Parâmetros: sProduto, iFilialForn, lFornecedor
'A associação do Produto %s com a Filial %i do Fornecedor %l não está cadastrada
'no Banco de dados.
Public Const ERRO_EXCLUSAO_COTACAOITEMCONCORRENCA = 12189 'Sem parâmetros
'Erro na exclusão de registro na tabela Cotação Item Concorrência.
Public Const ERRO_INSERCAO_COTACAO = 12190 'Sem parametro
'Erro na tentativa de inserção na tabela de Cotação.
Public Const ERRO_GRID_FORN_LINHA_NAO_SELECIONADA = 12191 'Sem parametros
'Deve ser selecionada alguma linha do Grid de Fornecedores.
Public Const ERRO_PED_COTACAO_NAO_GERADO = 12192 'Sem parametros
'É necessário gerar Pedidos de Cotação antes de imprimir.
Public Const ERRO_INSERCAO_COTACAOCONDPAGTO = 12193 'Sem parametros
'Erro na insercao de Condição de Pagamento na tabela CotacaoCondPagto.
Public Const ERRO_GRID_PRODUTOS_VAZIO = 12194 'Sem parametros
'O Grid de Produtos está vazio
Public Const ERRO_USUARIO_INEXISTENTE = 12195 'Parametro: sNomeReduzido
'O Usuário com Nome Reduzido %s não existe.
Public Const ERRO_PRODUTO_REPETIDO_GRID_PRODUTOS = 12196 'Parametro: sCodigoProduto
'O Produto %s só pode aparecer uma vez no Grid de Produtos.
Public Const ERRO_QUANTIDADE_COTAR_NAO_PREENCHIDA = 12197 'Parametro: sCodigo
'Produto %s não tem quantidade a cotar preenchida
Public Const ERRO_INSERCAO_PEDIDOCOTACAO = 12198 'Parametro: lNumIntPedCotacao
'Erro na inserção do Pedido de Cotação com número interno %l.
Public Const ERRO_INSERCAO_ITEMPEDCOTACAO = 12199 'Parametro: lNumInt
'Erro na inserção do Item de Pedido de Cotação com número interno %l.
Public Const ERRO_INSERCAO_COTACAOPRODUTO = 12200 'Sem parametros
'Erro na insercao de Cotação Produto.
Public Const ERRO_PRODUTO_SEM_FORNECEDOR_ESCOLHIDO = 12201 'Parametro: sCodProduto
'Para o Produto %s não foi escolhido nenhum fornecedor.
Public Const ERRO_AUSENCIA_FILIAL_PRODUTO_FORNECEDOR = 12202 'Parametros: sFornNomeReduzido,sCodProduto
'Não existe Filial do Fornecedor %s cadastrada para o Produto %s.
Public Const ERRO_FILIALEMPRESA_NAO_CADASTRADA1 = 12203  'parametro sNome
'FilialEmpresa %s não está cadastrada.
Public Const ERRO_GRID_REQUISICOES_VAZIO = 12204 'Sem parametros
'O Grid de Requisições está vazio.
Public Const ERRO_REQUISICAOINICIAL_MAIOR_REQUISICAOFINAL = 12205 'Sem parametros
'Requisição Inicial não pode ser maior do que a Requisição Final.
Public Const ERRO_GRID_REQUISICAO_NAO_SELECIONADO = 12206 'Sem parametros
'Não foi selecionado Requisição do Grid de Requisições.
Public Const ERRO_GRID_PRODUTOS_NAO_SELECIONADO = 12207 'Sem parametros
'Não foi selecionado Produto do Grid de Produtos.
Public Const ERRO_GRID_ITENS_REQUISICAO_NAO_SELECIONADO = 12208 'Sem parametros
'Não foi selecionado Item de Requisição no Grid de Itens de Requisições.
Public Const ERRO_QUANT_COTAR_ITEM_NAO_PREENCHIDA = 12209 'Parametro iLinhaGrid
'O Item de Requisição %i não está com a quantidade a cotar preenchida.
Public Const ERRO_INSERCAO_COTACAOPRODUTOITEMRC = 12210 'Sem parametros
'Erro na tentativa de inserção na tabela CotacaoProdutoItemRC.
Public Const ERRO_LEITURA_GERACAOPTOPEDIDO = 12211 'Sem Parametros
'Erro na Leitura das Tabelas Produtos, EstoqueProduto, ProdutoFilial e Almoxarifado.
Public Const ERRO_PEDIDOCOMPRA_NAO_ENVIADO = 12212 'Parametro lCodigo
'O pedido de compras com o código %l não foi enviado.
Public Const ERRO_PEDIDOCOMPRA_NAO_GERADO = 12213 ' Parametro lCodigo
'O Pedido de Compra %l não é um pedido de compra gerado.
Public Const ERRO_CLIENTE_DESTINO_NAO_PREENCHIDO = 12214
'O cliente de destino não foi informado.
Public Const ERRO_FILIALFORN_DESTINO_NAO_PREENCHIDA = 12215
'A filial do fornecedor de destino não foi preenchida.
Public Const ERRO_CONDICAOPAGTO_NAO_DISPONIVEL = 12216 'Sem parâmetros
'A Condição de Pagamento não pode ser À Vista. Escolha uma Condição de Pagamento A Prazo.
Public Const ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA1 = 12217 'Parâmetros: lNumIntDoc
'A Requisição de Compras de número interno %l não está cadastrada no Banco de dados.
Public Const ERRO_TIPOTRIBUTACAO_NAO_CADASTRADA = 12218 'Parâmetros: iTipo
'O Tipo de Tributação de código %i não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_TIPOSTRIBUTACAOMOVTO = 12219 'Sem Parâmetros
'Erro na leitura da Tabela TiposTributacaoMovto.
Public Const ERRO_LEITURA_ITEMRCITEMCONCORRENCIA1 = 12220 'Sem parâmetros
'Erro na leitura da tabela ItemRCItemConcorrencia.
Public Const ERRO_FILIALCOMPRA_NAO_PREENCHIDA = 12221 'Sem parâmetros
'A Filial de Compra deve ser informada.
Public Const ERRO_REQUISICAOCOMPRA_ENVIADA = 12222 'Parâmetros: lCodigo
'A Requisição de Compras de código %l já foi enviada.
Public Const ERRO_REQUISICAO_COMPRA_BAIXADA = 12223 'Parâmetros: lCodigo
'A Requisição de compras de código %l já está baixada.
Public Const ERRO_CODIGO_REQUISICAO_COMPRA_EXISTENTE = 12224 'Parâmetros: lCodigo
'Já existe no Banco de Dados a Requisição Compras com o código %l.
Public Const ERRO_LEITURA_REQUISICAO_COMPRA_BAIXADA = 12225 'Parâmetros: lCodigo
'Erro na leitura da Requisição de código %l da tabela RequisicaoCompraBaixada.
Public Const ERRO_INSERCAO_REQUISICAO_COMPRA = 12226 'Parâmetros: lCodigo
'Erro na inserção da Requisição Compras de código %l.
Public Const ERRO_INSERCAO_ITEMREQCOMPRA = 12227 'Parâmetros: lNumIntDocItem
'Erro na tentativa de inserir o item de Requisição de Compras de número interno %l.
Public Const ERRO_ITENSREQCOMPRA_NAO_CADASTRADO1 = 12228 'Parâmetros: lCodReqCompras
'O item da Requisição de compras de código de %l não está cadastrado.
Public Const ERRO_REQUISICAO_COMPRA_ENVIADA = 12229 'Parâmetros: lCodigo
'Não é possível excluir a Requisição de compras de código %l porque ela já foi enviada.
Public Const ERRO_ALTERACAO_REQUISICAOCOMPRA = 12230 'Parâmetros: lCodigo
'Erro na atualização da Requisição de Compras de Código %l.
Public Const ERRO_REQUISICAO_COMPRAS_AUSENCIA_ITENS = 12231 'Parâmetros: lCodigo
'A Requisição de Compras de código %l não possui itens.
Public Const ERRO_LOCK_ITEMREQCOMPRA = 12232 'Parâmetros: lNumIntitem
'Erro na tentativa de fazer "lock" no item de número interno %l de Requisição de Compras.
Public Const ERRO_EXCLUSAO_ITEMREQCOMPRA = 12233 'Parâmetros: lNumIntitem
'Erro na exclusão do Item de número iterno %l de Requisição de Compras.
Public Const ERRO_FILIAL_FORN_PRODUTO_NAO_ASSOCIADOS = 12234 'Parâmetros: iCodFilial, sFornNomeRed, sCodProduto
'A Filial %i do Fornecedor %s não está associada ao Produto %s.
Public Const ERRO_FORNECEDOR_DESTINO_NAO_PREENCHIDO = 12235 'Sem parâmetros
'O Fornecedor Destino não foi preenchido.
Public Const ERRO_FILIALCLIENTE_DESTINO_NAO_PREENCHIDA = 12236 'Sem parâmetros
'A Filial do Cliente destino não foi informada.
Public Const ERRO_LEITURA_ITENSREQMODELO = 12237 'Parâmetros: lCodReqModelo
'Erro na leitura da tabela ItensReqModelo da Requisão modelo de código %l.
Public Const ERRO_REQUISICAO_MODELO_AUSENCIA_ITENS = 12239 'Parâmetros: lReqModelo
'A Requisição Modelo %l não possui itens.
Public Const ERRO_REQUISITANTE_NAO_PREENCHIDO = 12240 'Sem parâmetros
'O Requisitante deve ser preenchido.
Public Const ERRO_GRID_FORNECEDOR_NAO_PREENCHIDO = 12241 'Parâmetros: iLinha
'O Campo Fornecedor da linha %i do Grid não foi preenchido.
Public Const ERRO_FILIALFORN_NAO_ENCONTRADA_ASSOCIADA = 12242 'Parâmetros: sCodFornecedor, sCodProduto
'A Filial do Fornecedor %s não foi encontrada ou não está associada ao Produto %s.
Public Const ERRO_GRIDITENS_VAZIO = 12243 'Sem parâmetros
'Não existem itens no Grid para gravar.
Public Const ERRO_FILIALEMPRESA_DESTINO_NAO_PREENCHIDA = 12244 'Sem parâmetros
'A Filial Empresa de destino não foi informada.
Public Const ERRO_GRID_QUANTIDADE_NAO_PREENCHIDA = 12245 'Parâmetros: iLinha
'A Quantidade da linha %i do Grid não foi preenchida.
Public Const ERRO_LEITURA_REQUISICAOMODELO = 12246 'Parâmetros: lCodReqModelo
'Erro na leitura da Requisição Modelo de Código %l.
Public Const ERRO_LEITURA_REQUISICAOMODELO2 = 12247 'Parâmetros: lNumIntReqModelo
'Erro na leitura da Requisição Modelo de Número Interno %l.
Public Const ERRO_INSERCAO_REQUISICAOMODELO = 12248 'Parâmetros: lCodReqModelo
'Erro na inserção da Requisição Modelo de código %l.
Public Const ERRO_LOCK_REQUISICAOMODELO = 12249 'Parâmetros: lCodReqModelo
'Erro no "lock" da Requisição Modelo de código %l.
Public Const ERRO_ALTERACAO_REQUISICAOMODELO = 12250 'Parâmetros: lCodReqModelo
'Erro na atualização da Requisição Modelo de código %l.
Public Const ERRO_ITEMREQMODELO_NAO_CADASTRADO = 12251 'Parâmetros: lCodReqModelo
'O item da Requisição Modelo de código %l não está cadastrado.
Public Const ERRO_LOCK_ITEMREQMODELO = 12252 'Parâmetros: lNumIntReq
'Erro no lock no item de número interno %l de Requisição Modelo.
Public Const ERRO_REQUISICAOMODELO_NAO_CADASTRADA = 12253 'Parâmetros: lCodReqModelo
'A Requisição Modelo de código %l não está cadastrada no Banco de Dados.
Public Const ERRO_EXCLUSAO_REQUISICAOMODELO = 12254 'Parâmetros: lCodReqModelo
'Erro na tentativa de excluir a Requisição Modelo de código %l.
Public Const ERRO_EXCLUSAO_ITEMREQMODELO = 12255 'Parâmetros: lNumIntItemReq
'Erro na tentativa de excluir o item de número interno %l de Requisição Modelo.
Public Const ERRO_INSERCAO_ITEMREQMODELO = 12256 'Parâmetros: lNumIntItemReq
'Erro na tentativa de gravar o Item de número interno %l de Requisição Modelo.
Public Const ERRO_REQUISICAOMODELO_NAO_CADASTRADA1 = 12258 'Parâmetros: lNumIntDoc
'A Requisição Modelo de número interno %l não está cadastrada.
Public Const ERRO_ATUALIZACAO_ITENSREQMODELO = 12259 'Parâmetros: lNumItemReqModelo
'Erro na tentativa de atualizar o item de número interno %l de Requisição Modelo.
Public Const ERRO_LEITURA_REQUISICOES_COMPRA_BAIXA_RC = 12260 'Sem parametros
'Erro na leitura de Requisições de Compra para baixa.
Public Const ERRO_NUM_REQUISICOES_SELECIONADAS_SUPERIOR_MAXIMO = 12261 'Sem parametros
'O número de requisições selecionadas é superior ao número máximo possível para
'baixa. Restrinja mais a seleção para que o número de requisições lidas diminua.
Public Const ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO = 12262 'Parametros: lFornecedor, iFilialForn, sProduto,iFilialEmpresa
'O Fornecedor %l Filial %i não está cadastrado na Tabela FornecedorProdutoFF para o Produto %s FilialEmpresa %i
Public Const ERRO_PEDIDOCOMPRA_NAO_ENCONTRADO = 12263 'Sem parametros
'Não foi encontrado nenhum Pedido de Compra de acordo com a seleção informada.
Public Const ERRO_COTACAO_VINCULADA_PEDIDOCOT_NAO_CADASTRADA = 12264 'Parâmetros: lCodigo
'A Cotação vinculada ao Pedido de Cotação de código %l não está cadastrada no Banco de dados.
Public Const ERRO_ITEMCOTACAO_NAO_CADASTRADO = 12265 'Parâmetros: lNumIntItem
'O Item de número interno %l de Cotação não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_ITEMCOTACAO = 12266 'Parâmetros: lNumIntItem
'Erro na tentativa de fazer "lock" no Item de número interno %l de Cotação.
Public Const ERRO_EXCLUSAO_ITEMCOTACAO = 12267 'Parâmetros: lNumIntItem
'Erro na exlusão do Item de número interno %l de Cotação.
Public Const ERRO_INSERCAO_ITEMCOTACAOBAIXADO = 12268  'Parâmetros: lNumIntItem
'Erro na inserção do Item de número interno %l de Cotação na tabela ItensCotacaoBaixados.
Public Const ERRO_DATAVALIDADE_INICIAL_MAIOR = 12269 'Sem parâmetros
'A data de validade inicial não pode ser maior que a data
'de validade final.
Public Const ERRO_LEITURA_PEDIDOCOTACAO1 = 12270 'Sem parâmetros
'Erro na leitura da tabela PedidoCotacao.
Public Const ERRO_AUSENCIA_PEDIDOCOTACAO = 12271 'Sem parâmetros
'Não há Pedidos de Cotação para a Seleção atual.
Public Const ERRO_NUM_PEDIDOS_SUPERIOR_MAXIMO = 12272 'Parâmetros: NUM_MAX_PEDCOTACOES
'O número de Pedidos de Cotação da seleção atual ultrapassou o limite que é %i.
Public Const ERRO_AUSENCIA_PEDCOTACAO_SELECIONADOS = 12273 'Sem parâmetros
'Não há Pedidos de Cotação selecionados para baixar.
Public Const ERRO_PEDCOTACAO_VINCULADO_PEDCOMPRA = 12274 'Parâmetros: lCodPedCotacao, lCodPedCompra
'O Pedido de Cotação de código %l está vinculado ao Pedido de Compra não Baixado de código %l.
Public Const ERRO_PEDCOTACAO_VINCULADO_CONCORRENCIA = 12275 'Parâmetros: lCodPedCotacao
'O Pedido de Cotação de código %l está vinculado a uma Concorrência não baixada.
Public Const ERRO_LEITURA_ITENSREQCOMPRATODOS = 12276 'Sem parâmetros
'Erro na leitura da tabela ItensReqComprasTodos.
Public Const ERRO_LEITURA_ITENSCONCORRENCIATODOS = 12278  'Sem parâmetros
'Erro na leitura da tabela ItensConcorrenciaTodos.
Public Const ERRO_LEITURA_COTACAOTODAS = 12279 'Sem parâmetros
'Erro na leitura da tabela CotacaoTodas.
Public Const ERRO_MOTIVO_NAO_CADASTRADO1 = 12280 'Parâmetros: sDescricao
'O Motivo de Descrição %s não está cadastrado no Banco de dados.
Public Const ERRO_LEITURA_QUANTIDADESSUPLEMENTARES = 12281 'Sem parametros
'Erro na leitura da tabela Quantidades Suplementares.
Public Const ERRO_LEITURA_MOTIVO = 12282 'Sem parâmetros
'Erro na leitura da tabela Motivo.
Public Const ERRO_LEITURA_REQUISICAOCOMPRA2 = 12283 'Sem parâmetros
'Erro na leitura da tabela RequisiçãoCompra.
Public Const ERRO_QUANTIDADE_INFERIOR_INICIAL = 12284 'Sem parametros
'A quantidade atual é inferior à quantidade já existente.
Public Const ERRO_COTACAO_NAO_CADASTRADA1 = 12285 'Sem Parametros
'A cotação não está cadastrada.
Public Const ERRO_MOTIVO_NAO_ENCONTRADO = 12286 'Parametro sMotivoEscolha
'O motivo de escolha %s não está cadastrado.
Public Const ERRO_DATANECESSIDADE_ANTERIOR_DATAPEDIDO = 12287 'Sem Parametros
'A Data de Necessidade não pode ser anterior a Data do Pedido.
Public Const ERRO_FALTA_TIPO = 12288 'Sem Parametros
'Falta selecionar Tipo(s) de Produto.
Public Const ERRO_AUSENCIA_ITENS_GRID = 12289 'Sem Parametros
'Pelo menos uma linha do grid de itens deve ser preenchida.
Public Const ERRO_QUANTCOMPRAR_NAO_PREENCHIDA = 12290 'Sem parametros
'A quantidade a comprar do grid de Produtos deve ser preenchida.
Public Const ERRO_QUANTCOTACAO_DIFERENTE_QUANTCOMPRAR = 12291 'Parametro: sProduto
'A quantidade total selecionada nos itens de cotação para o produto %s é
'diferente da quantidade a comprar informada.
Public Const ERRO_ATUALIZACAO_ITENSCONCORRENCIA = 12292 'Parametro: lCodConcorrencia
'Erro na tentativa de atualização do Item de Concorrência da Concorrência %l.
Public Const ERRO_ATUALIZACAO_COTACAOITEMCONCORRENCIA = 12293 'Sem parametros
'Erro na tentativa de atualização da tabela CotacaoItemConcorrencia.
Public Const ERRO_ATUALIZACAO_QUANTIDADESSUPLEMENTARES = 12294 'Sem parametros
'Erro na tentativa de atualização da tabela QuantidadesSuplementares
Public Const ERRO_INSERCAO_QUANTIDADESSUPLEMENTARES = 12295 'Sem parametros
'Erro na tentativa de inserção na tabela QuantidadesSuplementares
Public Const ERRO_INSERCAO_ITEMRCITEMCONCORRENCIA = 12296 'SEm parametros
'Erro na tentativa de inserção na tabela ItemRCItemConcorrencia.
Public Const ERRO_ATUALIZACAO_ITEMRCITEMCONCORRENCIA = 12297 'Sem parametros
'Erro na tentativa de atualização da tabela ItemRCItemConcorrencia.
Public Const ERRO_INSERCAO_COTACAOITEMCONCORRENCIA = 12298 'Sem parametros
'Erro na tentativa de inserção na tabela CotacaoItemConcorrencia.
Public Const ERRO_REQUISICAO_NAO_SELECIONADA = 12299 'Sem parâmetros
'Pelo menos uma requisição deve ser selecionada.
Public Const ERRO_QUANTCOTACAO_DIFERENTE_SOMAITENSREQ = 12300 'Sem parâmetros
'A quantidade total selecionada nos itens de cotação  é diferente da soma dos itens de requisição.
Public Const ERRO_ITEM_REQUISICAO_NAO_SELECIONADO = 12301 'Sem parâmetros
'Pelo menos um item de requisição deve ser selecionado.
Public Const ERRO_NUM_REQUISICOES_SELECIONADAS_SUPERIOR_MAXIMO1 = 12302 'Sem parametros
'O número de requisições selecionadas é superior ao número máximo possível.
'Restrinja mais a seleção para que o número de requisições lidas diminua.
Public Const ERRO_SELECAO_REQUISITANTE_AUTOMATICO = 12303
'Não é possível selecionar o Requisitante automático.
Public Const ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_RC = 12304 'Sem parâmetro
'Erro na leitura de pedidos de venda para a geração de requisições de compra.
Public Const ERRO_QUANTCOMPRAR_ITENSPED = 12305 'Sem parametros
'Ao menos um item de pedido de venda deve ter uma quantidade a comprar.
Public Const ERRO_PEDIDO_VENDA_NAO_ENCONTRADO = 12306 'Sem parametros
'Não foi possível encontrar nenhum Pedido de Venda de acordo com a seleção informada.
Public Const ERRO_AUSENCIA_PEDVENDAS_GRID = 12307 'Sem parametros
'Não há pedidos de venda no grid Pedidos para selecionar.
Public Const ERRO_AUSENCIA_PEDVENDA_SELECIONADO = 12308 'Sem parametros
'Ao menos um pedido de venda deve ser selecionado.
Public Const ERRO_QUANTCOMPRAR_PRODUTOS_DIFERENTE = 12309 'Parametro sCodigoProduto
'A quantidade a comprar do grid de produtos é diferente da soma
'das quantidades a comprar do itens de pedido com o produto %s.
Public Const ERRO_INSERCAO_CONCORRENCIA = 12310 'Parâmetros: lCodigo
'Erro na tentativa de inserir a Concorrência de código %l no Banco de Dados.
Public Const ERRO_LOCK_CONCORRENCIA = 12311 'Parâmetros:lCodigo
'Erro no Lock da Concorrência de código %l.
Public Const ERRO_ATUALIZACAO_CONCORRENCIA = 12312 'Parâmetros: lCodigo
'Erro na atualização da Concorrência de código %l.
Public Const ERRO_LOCK_ITENSCONCORRENCIA = 12313 'Sem parâmetros
'Erro na tentativa de fazer "lock" na tabela ItensConcorrencia.
Public Const ERRO_EXCLUSAO_ITENSCONCORRENCIA = 12314 'Parâmetros: lNumIntItem
'Erro na tentativa de excluir o Item Concorrência de número interno %l.
Public Const ERRO_EXCLUSAO_COTACAOITEMCONCORRENCIA = 12315 'Parâmetros: lNumIntItemConc
'Erro na exclusão de CotaçãoItemConcorrência com Item Concorrência de número interno %l.
Public Const ERRO_LOCK_COTACAOITEMCONCORRENCIA = 12316 'Sem parâmetros
'Erro na tentativa de fazer "lock" em CotaçãoItemConcorrência.
Public Const ERRO_LOCK_ITEMRCITEMCONCORRENCIA = 12317 'Sem parâmetros
'Erro na tentativa de fazer "lock" em ItemRCItemConcorrência.
Public Const ERRO_EXCLUSAO_ITEMRCITEMCONCORRENCIA = 12318 'Parâmetros: lNumIntItemConc
'Erro na exclusão de ItemRCItemConcorrência com Item Concorrência de número interno %l.
Public Const ERRO_EXCLUSAO_QUANTIDADESSUPLEMENTARES = 12319 'Parâmetros: lNumIntItemConc
'Erro na exclusão de QuantidadesSuplementares com Item Concorrência de número interno %l
Public Const ERRO_INSERCAO_ITEMCONCONCORRENICA = 12320 'Parâmetros: lNumIntItem
'Erro na tentativa de inserir o item de concorrência de número interno %l.
Public Const ERRO_QUANTCOMPRAR_MAIOR_RC = 12321 'Parâmetros: dQuantComprar, dQuantidade
'A quantidade a comprar %d não pode ser maior que a quantidade do Item de Requisição de compras que é %d.
Public Const ERRO_INSERCAO_ITEMRCITEMPC = 12322 'Parâmetros: lItemPedCompra, lItemRC
'Erro na inserção de registros na tabela ItemRCItemPC com ItemPC %l e ItemRC %l.
Public Const ERRO_ITEMRCITEMCONCORRENCIA_NAO_CADASTRADO = 12323 'Parâmetros: lItemRC, lItemConc
'O ItemRCItemConcorrência com ItemRC = %l e Item Concorrência %l não está cadastrado.
Public Const ERRO_QUANTCOTACAO_DIFERENTE_QUANTITEMCONC = 12324 'Parâmetros: sProduto, sFornecedor, iFilial
'A quantidade à comprar do item %s de Fornecedor %s e Filial de código %i
'é diferente da quantide a comprar do mesmo item marcado no Grid de Cotações.
Public Const ERRO_QUANTCOMPRAR_MENOR_QUANTCOMPRAR_RC = 12325 'Parâmetros: dQuantidade, dQuantTotalRC
'A quantidade a Comprar %d tem que ser maior que a soma das quantidades a comprar dos Itens
'de Requisição que é %d.
Public Const ERRO_QUANTCOMPRAR_SUPERIOR_MAXIMA = 12326 'Parâmetros: dQuantComprar, dQuantInicial
'A quantidade %d não pode ser superior a quantidade inicial %d.
Public Const ERRO_AUSENCIA_PEDCOTACAO_GRID = 12327 'Sem parâmetros
'Não há gerações de Pedidos de Cotação no Grid de geração
'de Pedidos de Cotação para selecionar.
Public Const ERRO_AUSENCIA_GERACAOPEDCOTACAO_SELECIONADA = 12328 'Sem parâmetros
'Uma geração de pedido de cotação deve ser selecionada.
Public Const ERRO_NENHUMA_COTACAO_SELECAO = 12329 'Sem parâmetros
'Não foi encontrada nenhuma cotação na seleção atual.
Public Const ERRO_NENHUM_TIPOPRODUTO_SELECIONADO = 12330 'Sem parâmetros
'Nehum Tipo de Produto foi selecionado.
Public Const ERRO_CONSUMOMEDIO_NAO_PREENCHIDO = 12331 'Parâmetros: iLinha
'O Consumo Médio da linha %i do Grid não foi preenchido.
Public Const ERRO_INTERVALORESSUP_NAO_PREENCHIDO = 12332 'Parâmetros: iLinha
'O Intervalo de Ressuprimento da linha %i do Grid não foi preenchido.
Public Const ERRO_NENHUM_PRODUTO_SELECIONADO = 12333 'Sem parâmetros
'Não foi encontrado nenhum Produto com a seleção atual.
Public Const ERRO_QUANTNAOEXCLUSIVA_DIFERENTE_QUANTEXCLUSIVA = 12334 'Parâmetros: sProduto
'A quantidade à comprar do Produto do %s do Grid de Produtos é diferente da
'quantidade à comprar do Grid de Cotações.
Public Const ERRO_CONCORRENCIA_AUSENCIA_ITENS = 12335 'Parâmetros: lCodConcorrencia
'A concorrência de código %l não possui Itens.
Public Const ERRO_INSERCAO_CONCORRENCIABAIXADA = 12336 'Parâmetros: lCodConcorrencia
'Erro na inserção da Concorrência de código %l na tabela ConcorrenciaBaixada.
Public Const ERRO_AUSENCIA_COTACOES_GRID = 12337 'Se parâmetros
'Não existem Cotações no Grid de Cotações.
Public Const ERRO_SELECAO_CONCORRENCIA = 12338 'Sem parâmetros
'Não há Concorrências cadastradas que obedecem seleção atual.
Public Const ERRO_QUANTCOMPRAR_MAIOR_QUANTCOMPRARMAX = 12339 'Parâmetros: dQuantComprar, dQuantMaxComprar
'A quantidade a comprar %d tem que ser menor que a quantidade máxima a comprar que
'é %d.
Public Const ERRO_MOTIVO_NAO_CADASTRADO = 12340 'Parâmetros: iCodigo
'O Motivo de Código %i não está cadastrado no Banco de dados.
Public Const ERRO_LOCK_ITENSREQCOMPRA1 = 12341 'Parâmetros:lNumIntDoc
'Erro no "lock" do Item de número interno %l de Requisição de Compras.
Public Const ERRO_AUSENCIA_CONCORRENCIAS_GRID = 12342 'Sem parâmetros
'Não há concorrências no Grid concorrências para selecionar.
Public Const ERRO_AUSENCIA_CONCORRENCIAS_SELECIONADAS = 12343 'Sem parâmetros
'Uma concorrência do Grid de concorrências deve ser selecionada.
Public Const ERRO_ITEMREQ_NAO_SELECIONADO = 12344 'Sem parâmetros
'Pelo menos um item de Requisição deve ser selecionado.
Public Const ERRO_CONCORRENCIA_NAO_CADASTRADA = 12345 'Parâmetros: lCodigo
'A Concorrência de código %l não está cadastrada no Banco de dados.
Public Const ERRO_LOCK_CONCORRECIA = 12346 'Parâmetros: lCodigo
'Erro na tentativa de fazer "lock" na Concorrência de código %l.
Public Const ERRO_EXCLUSAO_CONCORRENCIA = 12347 'Parâmetros: lCodigo
'Erro na exclusão da Concorrência de código %l.
Public Const ERRO_CODIGO_INICIAL_MAIOR_FINAL = 12348 'Sem parâmetros
'O código inicial não pode ser maior que o final.
Public Const ERRO_ITEMREQCOMPRA_NAO_CADASTRADO = 12349 'Parâmetros: sProduto, lReqCompra
'O Item com o Produto %s de Requisição de Compras de numero interno %l, não está
'cadastrado no banco de dados.
Public Const ERRO_ITEMCONCORRENCIA_NAO_VINCULADO_ITEMCOTACAO = 12350 'Parâmetros: iLinha
'O Item de Concorrência da linha %i do Grid de Produtos não está vinculado
'a nenhuma linha do Grid de Cotações.
Public Const ERRO_COTACAOITEMCONCORRENCIA_NAO_CADASTRADA = 12351 'Parâmetros: lNumIntConcorrencia, lItemCotacao)
'CotaçãoItemConcorrência com Item Concorrência %l e Item Cotação %l não cadastrada.
Public Const ERRO_NENHUM_ITEMCONC_SELECIONADO = 12352 'Sem parâmetros
'Nenhuma linha do Grid de Produtos foi selecionada.
Public Const ERRO_ITEMREQCOMPRA_NAO_CADASTRADO1 = 12353 'Parâmetros: lNumIntItem
'O Item de número interno %l de Requisiçãode Compras não está cadastrado.
Public Const ERRO_QUANTCOMPRAR_NAO_EXCLUSIVA_DIFERENTE = 12354 'Parâmetros: dQuantProd1, dQuantSobraCot
'A quantidade a Comprar não exclusiva %d é diferente da quantidade
'que sobrou das Cotações, que é %d.
Public Const ERRO_PEDCOMPRA_NAO_GERADO = 12355 'Parametro: lCodigoPedido
'O Pedido de Compras com código %l não é um pedido gerado.
Public Const ERRO_LEITURA_ITEMPEDCOTACAO1 = 12356 'SEm parametros
'Erro na leitura da tabela ItemPedCotacao
Public Const ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO1 = 12357 'parametro:lCodigoPedCompra
'Não existe Pedido de Cotação que tenha gerado o Pedido de Compra %lCodigo.
Public Const ERRO_FORNECEDOR_ENTREGA_NAO_PREENCHIDO = 12358 'Sem parametros
'O campo Fornecedor de Entrega não foi preenchido.
Public Const ERRO_LEITURA_ITEMPEDCOTACAOTODOS = 12359 'Sem parâmetros
'Erro na leitura da tabela ItemPedCotacaoTodos.
Public Const ERRO_LEITURA_ITENSCOTACAOTODOS = 12360 'Sem parâmetros
'Erro na leitura da tabela ItensCotacaoTodos.
Public Const ERRO_EXCLUSAO_REQUISITANTE_AUTOMATICO = 12362  'Sem parâmetros
'Não é possível excluir ou alterar o Requisitante Automático.
Public Const ERRO_LEITURA_ITEMNFITEMRC1 = 12363 'Sem parâmetros
'Erro na leitura da tabela ItemNFItemRC.
Public Const ERRO_EXCLUSAO_ITEMNFITEMRC = 12364 'Parâmetros: lItemNF
'Erro na exclusão de ItemNFItemRC com item de nota fiscal de número interno %l.
Public Const ERRO_LEITURA_ITENSREQCOMPRASTODOS = 12365 'Sem parâmetros
'Erro na leitura da tabela ItensReqComprasTodos.
Public Const ERRO_LEITURA_ITEMPEDCOMPRASTODOS = 12366 'Sem parâmetros
'Erro na leitura da tabela ItemPedComprasTodos.
Public Const ERRO_PEDIDOCOMPRABAIXADO_NAO_CADASTRADO = 12367 'Parâmetros: lNumIntDoc
'O Pedido de Compras baixado de número interno %l não está cadastrado no Banco de dados.
Public Const ERRO_EXCLUSAO_ITEMPEDCOMPRABAIXADO = 12368 'Parâmetros: lNumIntItemPC
'Erro na tentativa de excluir o ItemPC Baixado de número interno %l.
Public Const ERRO_ITENSREQCOMPRA_NAO_CADASTRADO2 = 12369 'Parâmetros: lNumIntItem
'O Item com número interno %l de Requisição de Compras não está cadastrado no Banco de dados.
Public Const ERRO_LEITURA_REQUISICAOCOMPRABAIXADA = 12370 'Sem parâmetros
'Erro na leitura da tabela RequisicaoCompraBaixada.
Public Const ERRO_EXCLUSAO_ITEMREQCOMPRABAIXADO = 12371 'Parâmetros: lNumIntItem
'Erro na tentativa de excluir o Item de número interno %l de Requisição de Compras Baixadas.
Public Const ERRO_EXCLUSAO_REQUISICAOCOMPRABAIXADA = 12372 'Parâmetros: lNumIntDoc
'Erro na tentativa de excluir a Requisição de Compras Baixada de número interno %l.
Public Const ERRO_EXCLUSAO_PEDIDOCOMPRABAIXADO = 12373 'Parâmetros: lNumIntDoc
'Erro na tentativa de excluir o Pedido de Compras Baixado de número interno %l.
Public Const ERRO_EXCLUSAO_ITEMNFITEMPC = 12374 'Parâmetros: lNumItemNF
'Erro na tentativa de excluir registros na tabela ItemNFItemPC.
Public Const ERRO_TIPO_NOTA_FISCAL_DIFERENTE_COMPRAS = 12375 'Parâmetros: iTipo
'O Tipo de documento %i não é de compras.
Public Const ERRO_ALIQUOTAIPINF_DIFERENTE_PC = 12376 'Parâmetros: sProduto, dAliquotaIPINF, dAliquotaIPIPC
'O Produto %s com a alíquota IPI %d é diferente do valor da alíquota IPI no Item de Pedido de Compras que possui alíquotaIPI %d.
Public Const ERRO_ALIQUOTAICMSNF_DIFERENTE_PC = 12377 'Parâmetros: sProduto, dAliquotaICMSNF, dAliquotaICMSPC
'O Produto %s com a alíquota ICMS %d é diferente do valor da alíquota ICMS no Item de Pedido de Compras que possui alíquotaICMS %d.
Public Const ERRO_ITEM_PC_PRECOUNITARIO_DIFERENTE = 12378 'Parâmetros: sProduto
'O Produto %s de Pedido de Compra está com preço unitário diferente
'dos outros pedidos já selecionados.
Public Const ERRO_ITEM_PC_ALIQUOTAICMS_DIFERENTE = 12379 'Parâmetros: sProduto
'O Produto %s de Pedido de Compra está com Alíquota ICMS diferente
'dos outros pedidos já selecionados.
Public Const ERRO_ITEM_PC_ALIQUOTAIPI_DIFERENTE = 12380 'Parâmetros: sProduto
'O Produto %s de Pedido de Compra está com Alíquota IPI diferente
'dos outros pedidos já selecionados.
Public Const ERRO_EXCLUSAO_ITENSREQCOMPRA = 12381 'Parâmetros: lNumIntItem
'Erro na exclusão do Item de Requisição de Compras de número interno %l.
Public Const ERRO_LEITURA_ITENSPEDCOMPRABAIXADOS = 12382 'Sem parâmetros
'Erro na leitura da tabela ItensPedCompraBaixados.
Public Const ERRO_QUANTIDADE_DIFERENT_QUANTRECEBIDA = 12383 'Parâmetros: sProduto
'A soma das quantidades recebidas do produto %s no Grid de intens de Pedido de Compras
'é diferente da Quantidade informada.
Public Const ERRO_QUANTIDADE_MAIOR_TOTALRECEBER = 12384 'Parâmetros: dQuantidade, dQuantTotal
'A quantidade %d não pode ser maior que a soma das quantidades a receber de todos os
'itens de pedidos de compra que é %d.
Public Const ERRO_PEDIDOCOMPRA_BAIXADO = 12385 'Parâmetros: lCodigo
'O Pedido de Compras  de código %l está baixado.
Public Const ERRO_QUANTRECEBIDA_MAIOR_QUANTRECEBER = 12386 'Sem parâmetros
'A quantidade recebida é maior que a quantidade a receber.
Public Const ERRO_QUANTRECEBIDARC_MAIOR_QUANTRECEBIDAPC = 12387 'Sem parâmetros
'A quantidade recebida é maior que a quantidade recebida do item de Pedido de Compras.
Public Const ERRO_QUANTRECEBIDARC_MAIOR_QUANTRECEBERRC = 12388 'Sem parâmetros
'A quantidade recebida do item Requisição é maior que a quantidade a receber do item da Requisição.
Public Const ERRO_ITEMNFITEMPC_NAO_CADASTRADO = 12389 'Parâmetros: lNumIntItemNF, lNumIntItemPedCompra
'ItemNFItemPC com ItemNF = %l e Item de Pedido de Compras = %l não está cadastrado no Banco de dados.
Public Const ERRO_LOCK_ITEMNFITEMPC = 12390 'Parâmetros: lNumIntItemNF, lNumIntItemPedCompra
'Erro no "lock" da tabela ItemNFItemPC com ItemNF = %l e Item de Pedido de Compras = %l.
Public Const ERRO_LEITURA_ITEMNFITEMRC = 12391 'Parâmetros: lNumIntItemNF, lNumIntItemReqCompra
'Erro na leitura da tabela ItemNFItemRC com ItemNF = %l e Item de Requisição de Compras = %l.
Public Const ERRO_ITEMNFITEMRC_NAO_CADASTRADO = 12392 'Parâmetros: lNumIntItemNF, lNumIntItemReqCompra
'ItemNFItemRC com ItemNF = %l e Item de Requisição de Compras = %l não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_ITEMNFITEMRC = 12393 'Parâmetros: lNumIntItemNF, lNumIntItemReqCompra
'Erro na "lock" da tabela ItemNFItemRC com ItemNF = %l e Item de Requisição de Compras = %l.
Public Const ERRO_INSERCAO_ITEMNFITEMPC = 12394 'Parâmetros: lItemNF, lItemPedCompra
'Erro na tentativa de inserir registros em ItemNFItemPC com ItemNF = %l e Item Pedido de Compras = %l.
Public Const ERRO_INSERCAO_ITEMNFITEMRC = 12395 'Parâmetros: lItemNF, lItemReqCompra
'Erro na tentativa de inserir registros em ItemNFItemRC com ItemNF = %l e Item Requisição de Compras = %l.
Public Const ERRO_ITEMREQCOMPRA_INEXISTENTE = 12396 'Parâmetros: lItemReqCompra
'O Item de Requisição de Compras de número interno %l não está cadastrado no Banco de dados.
Public Const ERRO_AUSENCIA_PEDIDOCOMPRAS = 12397 'Parâmetros: sFornecedor, iFilial, iFilialCompra
'Não existem Pedidos de Compras para o Fornecedor %s, Filial %i e FilialCompra %i.
Public Const ERRO_INSERCAO_PEDIDOCOMPRABAIXADO = 12398 'Parâmetros: lCodigo
'Erro na inserção do Pedido de Compras de código %l na tabela PedidoCompraBaixado.
Public Const ERRO_LEITURA_ITEMNFITEMPC = 12399 'Sem parâmetros
'Erro na leitura da tabela ItemNFItemPC.
Public Const ERRO_REQUISICAOCOMPRA_NAO_ENCONTRADA = 12400
'Não existe Requisição de Compra de acordo com a seleção informada
Public Const ERRO_BLOQUEIO_ALCADA_EXISTENTE = 12401 'sem parametros
'A opção de Controle de Alçada não pode ser desmarcada pois existem
'bloqueios de alçada não liberados.
Public Const ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA2 = 12402
'Requisição de Compras não cadastrada.
Public Const ERRO_ITEMCOT_NAO_VINCULADO_ITEMCONC = 12403 'Parâmetros: iLinha
'O Item de Cotação da linha %i do Grid de Cotações não está vinculado
'a nenhum Item do Grid de Produtos.
Public Const ERRO_ITEMRC_NAO_VINCULADO_ITEMCONC = 12404 'Parâmetros: iLinha
'O Item da linha %i do Grid de Produtos não está vinculado com nenhum
'Item marcado no Grid de Itens de Requisição.
Public Const ERRO_GRID_PRECOUNITARIO_NAO_PREENCHIDO = 12405 'Parâmetros: iLinha
'O Preço unitário do item da linha %i do Grid de Cotações não foi preenchido.
Public Const ERRO_QUANTCOTACAO_MAIOR_QUANTREQUISITADA = 12406 'Parâmetros: iLinha, dQuantFaltaCotar
'A quantidade a ser cotada da linha %i do Grid Itens de Requisição
'não pode ser maior que a quantidade que falta ser cotada %d.
Public Const ERRO_ITEMREQCOMPRA_NAO_CADASTRADO2 = 12407 'Parâmetros: sProduto, lCodReq
'O Item com o Produto %s de Requisição de Compras de código %l não está cadastrado
'no Banco de dados.
Public Const ERRO_ITEM_NAO_VINCULADO_ITEMCOTACAO = 12408 'Parâmetros: iLinha
'O item da linha %s do Grid de Produtos não está vinculado a nenhum
'item do Grid de Cotações.
Public Const ERRO_LEITURA_COTACAOITEMCONCORRENCIABAIXADA = 12409 'Sem parametros
'Erro na leitura da tabela Cotacao Item Concorrência Baixada.
Public Const ERRO_LEITURA_CONCORRENCIABAIXADA = 12410 'Sem parametros
'Erro na leitura da tabela Concorrência Baixada.
Public Const ERRO_LEITURA_ITENSCONCORRENCIABAIXADAS = 12411 'Sem parametros
'Erro na leitura da tabela Itens Concorrência Baixados.
Public Const ERRO_COTACAOPRODUTOITEMRC_NAO_CADASTRADA = 12412 'Parâmetros: lItemReqCompra
'A CotacaoProdutoItemRC com ItemRC de número interno %l não está cadastrado no Banco de dados.
Public Const ERRO_LEITURA_PEDIDOCOTACAOBAIXADO = 12413 'Sem parâmetros
'Erro na leitura da tabela de Pedidos Cotação Baixados.
Public Const ERRO_FILIAL_INICIAL_MAIOR = 12414
'A Filial inicial é maior que a final.
Public Const ERRO_COMPRADOR_INICIAL_MAIOR = 12415
'O Comprador inicial é maior que o final.
Public Const ERRO_USUARIO_NAO_COMPRADOR2 = 12416 'Parametro sCodUsuario
'O usuário %s não é um comprador.
Public Const ERRO_DATAENVIO_INICIAL_MAIOR = 12417
'A Data de Envio inicial é maior que a final.
Public Const ERRO_DATALIMITE_INICIAL_MAIOR = 12418
'A Data Limite inicial é maior que a final.
Public Const ERRO_PC_INICIAL_MAIOR = 12419
'O código do Pedido de Compra inicial é maior que o final.
Public Const ERRO_DESCRICAO_INICIAL_MAIOR = 12420 'SEM PARAMETROS
'A Descrição inicial é maior que a final.
Public Const ERRO_NUMNF_INICIAL_MAIOR = 12421 'sem parametros
'O Número da Nota Fiscal inicial é maior que o final.
Public Const ERRO_SERIE_INICIAL_MAIOR = 12422 'sem parametros
'A Série inicial é maior que a final.
Public Const ERRO_CODIGO_OP_INICIAL_MAIOR = 12423
'Código da Ordem de Produção inicial é maior que o final.
Public Const ERRO_PEDCOTACAO_INICIAL_MAIOR = 12424 'SEm parametros
'O Pedido de Cotação inicial é maior que o final.
Public Const ERRO_PV_INICIAL_MAIOR = 12425
'Pedido de Venda inicial é maior que o final.
Public Const ERRO_NOMECLIENTE_INICIAL_MAIOR = 12426
'Nome do Cliente inicial é maior que o final.
Public Const ERRO_REQUISITANTE_INEXISTENTE = 12427
'O Requisitante informado não existe.
Public Const ERRO_CCL_INEXISTENTE = 12428
'O Centro de Custo informado não existe.
Public Const ERRO_NATUREZA_INICIAL_MAIOR = 12429
'A Natureza inicial é maior que o final.
Public Const ERRO_CODIGO_TIPO_PRODUTO_NAO_PREENCHIDO = 12430 'sem parametros
'Preenchimento do tipo de produto é obrigatório.
Public Const ERRO_TIPOPRODUTO_NAO_SELECIONADO = 12431 'Sem parametros
'Pelo menos um Tipo de Produto deve ser selecionado.
Public Const ERRO_ITENS_MESMO_LEQUE = 12432 'Parâmetros: iItem, iItemCOmparado
'O item %s é igual ao item %s. Eles deve se tornar um único item.
Public Const ERRO_REQCOMPRAS_IMPRESSAO = 12433 'Sem parametros
'Selecione uma Requisição de Compra para executar a impressão.
Public Const ERRO_PEDCOMPRA_IMPRESSAO = 12434 'Sem parametros
'Selecione um Pedido de Compra para executar a impressão.
Public Const ERRO_PRODUTO_JA_EXISTENTE_PEDCOMPRA = 12435 'sProduto, iItem
'O produto %s já participa deste Pedido de Compra no Item %i.
Public Const ERRO_PEDIDOCOMPRA_BLOQUEADO = 12436 'sParametro: lCodigo
'O Pedido de Compra com código %l é bloqueado.
Public Const ERRO_PRECO_ITEM_NAO_PREENCHIDO = 12437 'sParametro: iItem
'O Preço Unitário da linha %i do Grid de Itens não está preenchido.
Public Const ERRO_NUMERO_GERACAO_NAO_PREENCHIDO = 12438
'O número da geração deve ser preenchido.
Public Const ERRO_PEDCOTACAO_IMPRESSAO = 12439 'Sem parametros
'Selecione um Pedido de Cotação para executar a impressão.
Public Const ERRO_DATALIMITE_MAIOR_DATAENVIO = 12440 'Parâmetros: dtDataLimite, dtDataEnvio
'A data limite %s deve ser maior ou igual a data de envio, que é %s.
Public Const ERRO_DATALIMITE_INFERIOR_DATAREQ = 12441 'Sem Parâmetros.
'A Data limite não pode ser menor que a data da requisição.
Public Const ERRO_PRODUTO_LEQUE_GRID = 12442 'Parâmetros: sProduto
'O Produto %s só poderá se repetir no grid se o Fornecedor e Filial informados forem diferentes dos já informados.
Public Const ERRO_CODIGO_CONCORRENCIA_NAO_PREENCHIDO = 12443 'Sem Parâmetros
'O código da concorrência deve ser preenchido
Public Const ERRO_MOTIVO_EXCLUSIVO = 12444 'Sem Parâmetros.
'O motivo "Exclusividade" não pode ser escolhido pelo usuário.
Public Const ERRO_PRECOUNITARIO_ITEMCOTACAO_NAO_PREENCHIDO = 12445 'Parâmetro : ItemConcorrencia
'Uma cotação escolhida para o item %i do grid de produtos não está com o preço unitário preenchido.
Public Const ERRO_LOCALENTREGA_DIFERENTE_FILIALEMPRESA = 12446 'Sem parametros
'O local de entrega é diferente de FilialEmpresa.
Public Const ERRO_QUANT_COTAR_MAIOR_QUANTIDADE_ITEMPEDCOTACAO = 12447 'Sem parametros
'Não é possível gerar Pedido de Compra. A Quantidade do Item do Pedido de Cotação é diferente da Quantidade a Cotar dos Itens de Requisição.
Public Const ERRO_PRECOPRAZO_MENOR_PRECOVISTA = 12448 'sem parametros
'O Preço a Prazo não pode ser menor que o Preço à Vista.
Public Const ERRO_ATUALIZACAO_COTACAOPRODUTOITEMRC = 12449 'Sem parametros
'Erro na atualização da tabela CotaçãoProdutoItemRC
Public Const ERRO_DADOS_DESTINO_NAO_PREENCHIDOS = 12450 'Sem Parâmetros
'O itens de cotação dependem do preenchimento do Tipo de Destino.



'??? jones 28/10

Public Const ERRO_QUANTPRODUTO_MENOR_QUANTREQUISITADA = 12451 'Parametro: iLinha, dQuantCotarItens
'Quantidade a cotar do Produto da linha %i é menor que a quantidade a cotar de Requisições %d.
Public Const ERRO_COTACAOPRODUTO_NAO_ENCONTRADA = 12452 'sem parametros
'Cotação Produto não foi encontrada.
Public Const ERRO_PEDIDOCOTACAO_STATUS_NAOATUALIZADO_PC = 12453 'sem parametros
'O Pedido de Cotação não está atualizado. Não é possível gerar Pedidos de Compra.
Public Const ERRO_NENHUM_MES_FECHADO = 12454 'Sem parâmetros
'Não existe nenhum mês fechado para o cálculo de Parâmetros de Ponto de Pedido.







'Códigos de Avisos - Reservado de 15200 a 15299
Public Const AVISO_EXCLUSAO_ALCADA = 15200 'Parâmetro: sCodUsuario
'Confirma a exclusão da alçada do usuário com código %s?
Public Const AVISO_CONFIRMA_EXCLUSAO_COMPRADOR = 15201  'Parametro: sCodUsuario
'Confirma a exclusão do comprador %s ?
Public Const AVISO_EXCLUIR_REQUISITANTE = 15202 ' Parametro: lCodigo
'Confirma exclusão do Requisitante com código %l?
Public Const AVISO_EXCLUIR_TIPODEBLOQUEIOPC = 15203 ' Parametro :iCodigo
'Confirma exclusão do Tipo de Bloqueio de Pedido de Compra com codigo %i ?
Public Const AVISO_NUM_MAX_BLOQUEIOSPC_LIBERACAO = 15204
'O número máximo de Bloqueios possíveis de Pedido de Compras para exibição foi atingido . Ainda existem mais Bloqueios para Liberação .
Public Const AVISO_CRIAR_REQUISITANTE = 15205 'Parametro: lCodigo
'Requisitante com código %l não está cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_REQUISITANTE1 = 15206 'Parametro: sNomeReduzido
'Requisitante com Nome Reduzido %s não está cadastrado. Deseja cadastrar?
Public Const AVISO_EXCLUSAO_PEDIDOCOTACAO = 15207 'Parametros lCodigo
'Confirma a exclusão do Pedido de Cotação com o código %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_PEDIDO_COMPRA = 15208 'PArametro lCodigo
'Confirma exclusão do Pedido de Compra com o código %l?
Public Const AVISO_BAIXA_PEDIDOCOMPRAS = 15209 'Parametro lCodigo
'Deseja realmente baixar o pedido de compras %l ?
Public Const AVISO_EXCLUIR_REQUISICAOCOMPRA = 15210 'Parâmetros: lCodigo
'Confirma a exclusão da Requisição de Compras de código %l?
Public Const AVISO_EXCLUIR_REQUISICAOMODELO = 15211 'Parâmetros: lCodReqModelo
'Confirma a exclusão da Requisição Modelo de código %l?
Public Const AVISO_REQUISICAOCOMPRA_GERADA = 15213  'Parâmetros: lCodigo
'Foi gerada a Requisição de Compras de código %l.
Public Const AVISO_EXCLUIR_CONCORRENCIA = 15214 'Parâmetros: lCodigo
'Confirma exclusão da Concorrência de código %l?
Public Const AVISO_RECEBMATERIALFCOM_MESMO_NUMERO = 15215 ' Parametros lFornecedor, iFilialForn, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEntrada
'No Banco de Dados existe Recebimento de Material com os dados Código do Fornecedor %l, Código da Filial %i, Tipo %i, Série NF %s, Número NF %l, DataEntrada %dt. Deseja prosseguir na inserção de novo Recebimento de Material com o mesmo número de Nota Fiscal?
Public Const AVISO_VALORFRETE_DIFERENTE_PC = 15216 'Parâmetros: dValorFreteNF, dValorFretePC
'O Valor Frete %d da Nota Fiscal é diferente do Valor do Pedido de Compras que é %d. Deseja continuar com a gravação?
Public Const AVISO_VALORSEGURO_DIFERENTE_PC = 15217 'Parâmetros: dValorSeguroNF, dValorSeguroPC
'O Valor Seguro %d da Nota Fiscal é diferente do Valor do Pedido de Compras que é %d. Deseja continuar com a gravação?
Public Const AVISO_VALORDESCONTO_DIFERENTE_PC = 15218 'Parâmetros: dValorDescontoNF, dValorDescontoPC
'O Valor Desconto %d da Nota Fiscal é diferente do Valor do Pedido de Compras que é %d. Deseja continuar com a gravação?
Public Const AVISO_VALORDESPESAS_DIFERENTE_PC = 15219 'Parâmetros: dValorDespesasNF, dValorDespesasPC
'O Valor de Despesas %d da Nota Fiscal é diferente do Valor do Pedido de Compras que é %d. Deseja continuar com a gravação?
Public Const AVISO_VALORUNITARIO_DIFERENTE_PC = 15220 'Parâmetros: dValorUnitarioItemNF, iLinha, dValorUnitarioItemPC
'O Valor Unitário %d da linha %i do Grid de Itens de Notas Fiscais é diferente do Valor Unitário do Item de Pedido de Compras que é %d.
'Deseja continuar com a gravação?
Public Const AVISO_DESCONTOITEM_DIFERENTE_PC = 15221 'Parâmetros: dValorDescontoItemNF, iLinha, dValorDescontoItemPC
'O Valor Desconto %d da linha %i do Grid de Itens de Notas Fiscais é diferente da soma dos Valores de Desconto dos Itens de Pedido de Compras que é %d.
'Deseja continuar com a gravação?
Public Const AVISO_ALIQUOTAIPI_DIFERENTE_PC = 15222 'Parâmetros: sProduto, dAliquotaIPINF, dAliquotaIPIPC
'O Produto %s com a alíquota IPI %d é diferente do valor no Item de Pedido de Compras que possui alíquotaIPI %d. Deseja Continuar?
Public Const AVISO_ALIQUOTAICMS_DIFERENTE_PC = 15223 'Parâmetros: sProduto, dAliquotaICMSNF, dAliquotaICMSPC
'O Produto %s com a alíquota ICMS %d é diferente do valor no Item de Pedido de Compras que possui alíquotaICMS %d. Deseja Continuar?
Public Const AVISO_BAIXA_REQCOMPRAS = 15224 'Parâmetro: lCodReqCompra
'Deseja realmente baixar a requisição de compra %l ?
Public Const AVISO_LOTEMINIMO_MAIOR_QUANTCOMPRAR = 15225 'sProduto, dQuantidade, dLoteMinimo
'O Produto %s possui quantidade %d menor que o lote minimo relacionado ao Fornecedor Filial Produto que é %d. Deseja prosseguir com a gravação?
Public Const AVISO_QUANTIDADE_PEDIDA_ESTOQUEMAXIMO = 15226  'Parametro: sProduto
'A soma da quantidade pedida mais a quantidade disponível é maior que a quantidade de estoque máximo do produto %s. Confirma a quantidade informada?
Public Const AVISO_AUSENCIA_COTACOES_SELECAO = 15227 'Sem Parâmetros.
'Não existem cotações com dados compatíveis para atender a esse item.
Public Const AVISO_CONCORRENCIA_GRAVADA = 15228 'Parâmetros: lCodigo
'A concorrência %s foi gravada com sucesso.
Public Const AVISO_PEDIDOCOMPRA_GERADO = 15229
'O Pedido de Compra foi gerado.
Public Const AVISO_CANCELAR_CALCULO_PARAMETROS_PTOPEDIDO = 15230 'Sem parametros
'Confirma o canselamento do Calculo dos Parametros para Ponto Pedido ?
Public Const AVISO_ITEMRC_PARTICIPA_CONCORRENCIA = 15231
'Existem concorrência(s) anteriores envolvendo as requisições utilizadas nessa concorrência. Deseja prosseguir com essa gravação e excluir as concorências anteriores?


