Attribute VB_Name = "ErrosMAT"
Option Explicit

'Códigos de Erro  RESERVADO de 11200 a 11399 - ERROS MAT2
Public Const ERRO_PRODUTO_SEM_TIPO = 11200 'Parametros sCodigo
'Produto %s não tem Tipo de Produto associado.
Public Const ERRO_PRODUTO_MESMA_DESCRICAO = 11201 'sDescricaoProduto
'Já existe um Produto cadastrado com a Descrição = %s
Public Const ERRO_LEITURA_FORNECEDORPRODUTOFF = 11202 'Sem Parametro
'Erro na Leitura da Tabela FornecedorProdutoFF.
Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_PEDCOMPRA = 11204 'Parâmetros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'Não é possível excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'filial %i estão sendo utilizadas no Pedido de Compra de código %l.
Public Const ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF = 11205 'Parâmetros: lFornecedor, iFilial, sProduto
'Erro na tentativa de atualizar registro na tabela FornecedorProdutoFF com Fornecedor %l, Filial %i e Produto %s.
Public Const ERRO_INSERCAO_FORNECEDORPRODUTOFF = 11206 'Parâmetros: lFornecedor, sProduto
'Erro na tentativa de inserir registro na tabela FornecedorProdutoFF com Fornecedor %l e Produto %s.
Public Const ERRO_LOCK_FORNECEDORPRODUTOFF = 11207 'Parâmetros: lFornecedor, sProduto
'Erro na tentativa de "lock" na tabela FornecedorProdutoFF com Fornecedor %l e Produto %s.
Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_REQUISICAOCOMPRA = 11208 'Parâmetros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'Não é possível excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'filial %i estão sendo utilizadas na Requisicão de Compra de código %l.
Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_CONCORRENCIA = 11209 'Parâmetros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'Não é possível excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'filial %i estão sendo utilizadas na Concorrência de código %l.
Public Const ERRO_PRODUTO_SEM_FORNECEDOR = 11210 'Parametro: sCodigo
'O Produto %s não tem Fornecedores cadastrados nessa Filial da Empresa.
Public Const ERRO_FORNECEDORPRODUTOFF_NAO_ENCONTRADO = 11212 'Parâmetros: lFornecedor, sProduto
'O Fornecedor %l do Produto %s não está cadastrado no Banco de Dados.
Public Const ERRO_INSERCAO_ITEMPEDCOTACAOBAIXADO = 11213 'Sem parâmetros
'Erro na tentativa de inserir registros na tabela Item Pedido Cotação Baixado.
Public Const ERRO_ITEMCOTACAO_VINCULADO_COTACAOITEMCONCORRENCIA = 11214 'Parâmetros: lFornecedor, iFilialForn, sProduto
'Não é possível excluir o registro Fornecedor %l, Filial %i e Produto %s pois eles estão vinculados com
'Cotação Item Concorência.
Public Const ERRO_INSERCAO_ITENSCOTACAOBAIXADOS = 11215 'Sem parâmetros
'Erro na tentativa de inserir registros na tabela Itens Cotação Baixados.
Public Const ERRO_LOCK_PEDIDOCOTACAO = 11216 'Sem parâmetros
'Erro na tentativa de fazer "lock" na tabela Pedido Cotação.
Public Const ERRO_INSERCAO_PEDIDOCOTACAOBAIXADO = 11217 'Sem parâmetros
'Erro na tentativa de inserir registros na tabela Pedido Cotação Baixado.
Public Const ERRO_EXCLUSAO_PEDIDOCOTACAO = 11218 'Sem parâmetros
'Erro na exclusão de registro na tabela Pedido Cotação.
Public Const ERRO_LEITURA_COTACAOPRODUTO = 11219 'Sem parâmetros
'Erro na leitura da tabela Cotação Produto.
Public Const ERRO_LOCK_COTACAO = 11220 'Sem parâmetros
'Erro na tantativa de fazer "lock" na tabela Cotação.
Public Const ERRO_INSERCAO_COTACAOPRODUTOBAIXADO = 11221 'Sem parâmetros
'Erro na tentativa de inserir registros na tabelaCotação Produto Baixado.
Public Const ERRO_LOCK_COTACAOPRODUTO = 11222 'Sem parâmetros
'Erro na tentativa de fazer "lock" na tabela Cotação Produto.
Public Const ERRO_EXCLUSAO_COTACAOPRODUTO = 11223 'Sem parâmetros
'Erro na tentativa de excluir registros da tabela Cotação Produto.
Public Const ERRO_INSERCAO_COTACAOBAIXADA = 11224 'Sem parâmetros
'Erro na tentativa de inserir registros na tabela Cotação Baixada.
Public Const ERRO_EXCLUSAO_COTACAO = 11225 'Sem parâmetros
'Erro na exclusão de registro na tabela Cotação.
Public Const ERRO_EXCLUSAO_FORNECEDORPRODUTOFF = 11226 'Parâmetros: sProduto
'Erro na tentativa de excluir registros de FornecedorProdutoFF com Produto = %s.
Public Const ERRO_PRODUTO_FORNECEDORPRODUTOFF = 11227 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado a Fornecedor ProdutoFF.
Public Const ERRO_LEITURA_SLDDIAFORN = 11228 'Sem parametros
'Ocorreu um erro na leitura da tabela de saldos diários de fornecedor.
Public Const ERRO_LEITURA_SLDMESFORN = 11229 'Sem parametros
'Ocorreu um erro na leitura da tabela de saldos mensais de fornecedor.
Public Const ERRO_PRODUTO_SLDMESFORN = 11230 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado a um Saldo Mensal de Fornecedor.
Public Const ERRO_PRODUTO_SLDDIAFORN = 11231 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado a um Saldo Diário de Fornecedor.
Public Const ERRO_PRODUTO_ITENSPEDCOMPRA = 11232 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado a um Item de Pedido de Compra.
Public Const ERRO_PRODUTO_ITENSREQCOMPRA = 11233 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado a um Item de Requisição de Compra.
Public Const ERRO_LEITURA_ITENSCONCORRENCIA1 = 11234 'sem parametros
'Erro na leitura da tabela ItensConcorrencia
Public Const ERRO_PRODUTO_ITENSCONCORRENCIA = 11235 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado a um Item de Concorrência.
Public Const ERRO_PRODUTO_COTACAOPRODUTO = 11236 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado a Cotação Produto.
Public Const ERRO_LEITURA_SLDMESEST1_2 = 11237  'Parametros:iAno, iFilialEmpresa
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i.
Public Const ERRO_LEITURA_SLDMESEST2_2 = 11238  'Parametros:iAno, iFilialEmpresa
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i.
Public Const ERRO_LOCK_SLDMESEST_1 = 11239 'Parametros iAno, iFilialEmpresa
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst). Ano=%i, FilialEmpresa=%i
Public Const ERRO_LOCK_SLDMESEST1_1 = 11240 'Parametros iAno, iFilialEmpresa
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i
Public Const ERRO_LOCK_SLDMESEST2_1 = 11241 'Parametros iAno, iFilialEmpresa
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i
Public Const ERRO_ABERTURA_NOVOMES_SLDMESEST1 = 11242 'Parametros: iAno, iFilialEmpresa, iMes
'Não foi possivel abrir um novo mês (SldMesEst1) com os dados a seguir. Ano: i%, FilialEmpresa: %i, Mes: %i
Public Const ERRO_ABERTURA_NOVOMES_SLDMESEST2 = 11243 'Parametros: iAno, iFilialEmpresa, iMes
'Não foi possivel abrir um novo mês (SldMesEst2) com os dados a seguir. Ano: i%, FilialEmpresa: %i, Mes: %i
Public Const ERRO_PRODUTO_ALTERACAO_RASTRO = 11244 'Parametro= sProduto
'Não é permitida a alteração do Tipo de Rastreamento do Produto %s pois já existem movimentos de estoque.
Public Const ERRO_FILIAL_OP_NAO_PREENCHIDA = 11245 'iLinha
'A Filial da O.P. não foi preenchida na linha %s.
Public Const ERRO_LOTE_GRID_NAO_PREENCHIDO = 11246 'iLinha
'O Lote não foi preenchido na linha %s.
Public Const AVISO_LOTE_PRODUTO_INEXISTENTE = 11247 'Parâmetros: sLote, sProduto
'Não existe lote %s para o produto %s. Deseja cadastrá-la?
Public Const ERRO_QUANTTOTAL_LOTE_MAIOR_ALMOXARIFADO = 11248 'Parâmetros: dQuantTotalLote, dQuantAlmoxarifado
'A quantidade total dos lotes %s não pode ser maior que a quantidade do almoxarifado que é %s.
Public Const ERRO_LEITURA_RASTREAMENTOLOTE = 11249 'Sem parâmetros
'Erro na leitura da tabela RastreamentoLote.
Public Const ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO = 11250 'Sem parâmetros
Public Const ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO1 = 11251 'Parâmetros: sProduto, lLote
'O RastreamentoLote com Produto %s, Lote de código %s e FilialOP = %s não está cadastrado.
Public Const ERRO_LEITURA_RASTREAMENTOLOTE2 = 11252 'Parametros sProduto, sLote, iFilialOP
'Ocorreu um erro na leitura da tabela de RastreamentoLote. Produto=%s, Lote=%s, Filial = %i.
Public Const ERRO_LOCK_RASTREAMENTOLOTE = 11253  'Parametros sProduto, sLote, iFilialOP
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de RastreamentoLote. Produto=%s, Lote=%s, Filial = %i.
Public Const ERRO_LOTE_RASTREAMENTOLOTE_NAO_ABERTO = 11254  'Parametros sProduto, sLote, iFilialOP
'O Lote de Rastreamento não está com status aberto. Produto=%s, Lote=%s, Filial = %i.
Public Const ERRO_RASTREAMENTOLOTE_MOV_NAO_PRODUCAO = 11255 'Parametros sProduto, sLote, iFilialOP
'O Lote de Rastreamento não está cadastrado e o movimento de estoque não é uma produção de material. Produto=%s, Lote=%s, Filial = %i.
Public Const ERRO_INSERCAO_RASTREAMENTOLOTE = 11256 'Parametros sProduto, sLote, iFilialOP
'Ocorreu um erro na inserção de um registro natabela de RastreamentoLote. Produto=%s, Lote=%s, Filial = %i.
Public Const ERRO_LEITURA_RASTREAMENTOLOTESALDO = 11257 'Parametros sProduto, iAlmoxarifado, Lote
'Ocorreu um erro na leitura da tabela de RastreamentoLoteSaldo. Produto=%s, Almoxarifado=%i, Lote = %s.
Public Const ERRO_LOCK_RASTREAMENTOLOTESALDO = 11258  'Parametros sProduto, iAlmoxarifado, Lote
'Ocorreu um erro na tentativa de fazer 'lock' tabela de RastreamentoLoteSaldo. Produto=%s, Almoxarifado=%i, Lote = %s.
Public Const ERRO_SALDO_MAT_CONSERTO_RASTRO = 11259 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto em conserto no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s,  Saldo = %d.
Public Const ERRO_SALDO_MAT_CONSERTO3_RASTRO = 11260 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto de terceiros em conserto no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_CONSIG_RASTRO = 11261 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto em consignação no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_CONSIG3_RASTRO = 11262 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto de terceiros em consignação no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_DEFEITUOSO_RASTRO = 11263 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto defeituoso no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_DEMO_RASTRO = 11264 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto em demonstração no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_DEMO3_RASTRO = 11265 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto de terceiros no lote de rastreamento em demonstração é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_DISPONIVEL_RASTRO = 11266 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto disponível no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_INDISPONIVEL_RASTRO = 11267 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto indisponível no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_OUTRAS_RASTRO = 11268 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto em poder de terceiros no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_OUTRAS3_RASTRO = 11269 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto de terceiros no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_DISP_CONSIG3_RASTRO = 11270 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto disponível + o saldo de terceiros em consignação no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_BENEF_RASTRO = 11271 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto em poder de terceiros para beneficiamento no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_BENEF3_RASTRO = 11272 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto de terceiros em beneficiamento no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_SALDO_MAT_OP_RASTRO = 11273 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto em Ordem de Produção no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_ATUALIZACAO_RASTREAMENTOLOTESALDO = 11274 'Parametros sProduto, iAlmoxarifado, sLote
'Ocorreu um erro na atualização da tabela de RastreamentoLoteSaldo. Produto=%s, Almoxarifado=%i, Lote = %s.
Public Const ERRO_INSERCAO_RASTREAMENTOLOTESALDO = 11275 'Parametros sProduto, iAlmoxarifado, sLote
'Ocorreu um erro na inserção de um registro na tabela de RastreamentoLoteSaldo. Produto=%s, Almoxarifado=%i, Lote = %s.
Public Const ERRO_SALDO_MAT_RESDISP_RESCONSIG3_RASTRO = 11276 'Parametros sProduto, iAlmoxarifado, Lote, Saldo
'O Saldo do Produto Reservado disponível + o saldo de terceiros reservado em consignação no lote de rastreamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Lote=%s, Saldo = %d.
Public Const ERRO_INSERCAO_RASTREAMENTOMOVTO = 11277 'Parametros sProduto, lNumIntDocLote
'Ocorreu um erro na inserção de um registro na tabela de RastreamentoMovto. Produto=%s, NumIntDocLote = %l.
Public Const ERRO_INSERCAO_APROPRIACAOINSUMOSPROD = 11278 'Parametros sProduto
'Ocorreu um erro na inserção de um registro na tabela de ApropriacaoInsumosProd. Produto=%s
Public Const ERRO_LEITURA_RASTREAMENTOMOVTO = 11279 'Parametros lNumIntDocOrigem, iTipoDocOrigem
'Ocorreu um erro na leitura da tabela de RastreamentoMovto. Número Interno do Documento de Origem do Rastreamento = %l, Tipo do Documento de Origem = %i.
Public Const ERRO_EXCLUSAO_RASTREAMENTOMOVTO = 11280 'Parametros sProduto, lNumIntDocLote
'Ocorreu um erro na exclusão de um registro da tabela de RastreamentoMovto. Produto=%s, NumIntDocLote = %l.
Public Const ERRO_LOTE_NAO_CADASTRADO_RASTREAMENTO = 11281 'Parametros sLote, sProduto, iFilialOP
'O lote do produto não está cadastrado. Lote = %s, Produto=%s, FilialOP = %i.
Public Const ERRO_LEITURA_APROPRIACAOINSUMOSPROD = 11282 'Parametros lNumIntDocOrigem
'Ocorreu um erro na leitura da tabela de ApropriacaoInsumosProd. Número Interno do Mov. Estoque Origem da Apropriação dos Insumos = %l.
Public Const ERRO_EXCLUSAO_APROPRIACAOINSUMOSPROD = 11283 'Parametros lNumIntDoc
'Ocorreu um erro na exclusão de um registro da tabela de ApropriacaoInsumosProd. Numero Interno do Documento = %l.
Public Const ERRO_MOVIMENTOESTOQUE_SEM_ALTERACAO = 11284  'Parametros iFilialEmpresa, lCodigo
'Ocorreu um erro. Esta operação não alterou a movimentação de estoque. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_QUANT_INSUF_RASTREAMENTO = 11285 ' Parametros sProduto, iAlmoxarifado, iEscaninho
'Não há quantidade suficiente nos lotes para atender o movimento de estoque em questão - Rastreamento. Produto = %s, Almoxarifado = %i, Escaninho = %i.
Public Const ERRO_LEITURA_RASTREAMENTOLOTESALDO2 = 11286 'Parametros sProduto, iAlmoxarifado
'Ocorreu um erro na leitura da tabela de RastreamentoLoteSaldo. Produto=%s, Almoxarifado=%i.
Public Const ERRO_MOVESTOQUE_NAO_CADASTRADO1 = 11287 'Parametro lNumIntDoc
'O Movimento de Estoque com NumIntDoc = %l não está cadastrado no Banco de Dados.
Public Const ERRO_PRODUTO_LIVREGCADPROD = 11288 'sem parametros
'O Produto em questão não pode ser excluído pois está relacionado com Livros Fiscais.
Public Const ERRO_LEITURA_IMPORTPROD = 11289 'Sem parametros
'Erro na leitura da tabela ImportProd.
Public Const ERRO_LEITURA_IMPORTPRODAUX = 11290 'Sem parametros
'Erro na leitura da tabela ImportProdAux.
Public Const ERRO_LEITURA_IMPORTPRODDESC = 11291 'Sem parametros
'Erro na leitura da tabela ImportProdDesc.
Public Const ERRO_PRODUTOINSUMO_NAO_PREENCHIDO = 11292 'Sem parâmetros
'O Produto Insumo deve estar preenchido.
Public Const ERRO_REQUISICAO_SEM_PRODUTO = 11293  'Parametros sProduto, lCodigo, iFilialEmpresa
'O Produto %s não está na Requisição %s da Filial Empresa %s.
Public Const ERRO_REQUISICAO_PRODUTO_SEM_QUANTIDADE = 11294  'Parametros sProduto, lCodigo, dQuantidade
'O produto %s foi requisitado na Requisição %s a quantidade %s, não é possivel utilizar uma quantidade maior que a requerida.
Public Const ERRO_REQUISICAO_NAO_CADASTRADA = 11295  'Parametros lCodigo, iFilialEmpresa
'A Requisição %l para Filial Empresa %i não está cadastrada.
Public Const ERRO_PRODUTOINSUMO_NAO_ENCONTRADO_GRIDMOVIMENTOS = 11297 'Parametros: sProdutoInsumo, iLinha.
'O Produto Insumo %s da linha %i não foi encontrado no Grid de Movimentos.
Public Const ERRO_LEITURA_TABELA_APROPRIACAOINSUMOPROD = 11299 'Sem parametros
'Erro de leitura na Tabela ApropriacaoInsumoProd.
Public Const ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO = 11300 'Sem parametros
'O preenchimento do Lote na linha %s .
Public Const ERRO_OP_RASTREAMENTO_INEXISTENTE = 11301 'Parâmetros: CodigoOP, iFilialOP, sProduto
'Não existe Rastreamento para O.P. %s, Filial %s e Produto %s.
Public Const ERRO_DATAVALIDADE_MENOR_DATAFABRICACAO = 11302 'Sem parâmetros
'A data de validade não pode ser menor que a data de fabricação.
Public Const ERRO_LOTE_BAIXADO = 11303 'Parâmetros: sLote
'O Lote de código %s está baixado.
Public Const ERRO_EXCLUSAO_RASTREAMENTOLOTE = 11305 'Parâmetros: sLote
'Erro na tentativa de excluir RastreamentoLote com Lote de código %s.
Public Const ERRO_ATUALIZACAO_RASTREAMENTOLOTE = 11306 'Parâmetros: sLote
'Erro na atualização RastreamentoLote com Lote de código %s.
Public Const ERRO_RASTREAMENTOOP_NAO_CADASTRADO = 11308 'Parâmetros: sProduto, sLote
'O Rastreamento de OP com Produto = %s, Lote = %s e FilialOP = %s não está cadastrado.
Public Const ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO = 11309 'Parâmetros: sProduto, sLote
'O Rastreamento de Lote com Produto %s e Lote de código %s não está cadastrado.
Public Const AVISO_EXCLUSAO_RASTREAMENTOLOTE = 11310 'Parâmetros: sProduto, lLote
'Confirma exclusão de RastreamentoLote com Produto %s e Lote de código %l?
Public Const ERRO_LEITURA_RASTREAMENTOLOTEMOVTO = 11312 'Sem parâmetros
'Erro na leitura da tabela RastreamentoLoteMotvto.
Public Const ERRO_LOTE_VINCULADO_RASTREAMENTOLOTESALDO = 11313 'Parâmetros: sProduto, lLote
'O RastreamentoLote com Produto %s e código do Lote %s está vinculado a RastreamentoLoteSaldo.
Public Const ERRO_LOTE_VINCULADO_RASTREAMENTOLOTEMOVTO = 11314 'Parâmetros: sProduto, lLote
'O RastreamentoLote com Produto %s e código do Lote %s está vinculado a RastreamentoLoteMovto.
Public Const ERRO_EXCLUSAO_RASTREAMENTOOP = 11315 'Sem parâmetros
'Não é possível excluir um Rastreamento vinculado a uma Ordem de Produção.
Public Const ERRO_INSERCAO_RASTROLOTEOP = 11316 'Sem parâmetros
'Não é possível criar Lotes para Rastreamento vinculado à Ordem de Produção.
Public Const ERRO_GRID_LOTE_NAO_PREENCHIDO = 11317 'Parâmetros: iLinha
'O lote da linha %i não foi preenchido.
Public Const ERRO_GRID_DATA_NAO_PREENCHIDA = 11318 'Parâmetros: iLinha
'A Data da linha %i não foi preenchida.
Public Const ERRO_QUANTLOTE_MAIOR_QUANTALM = 11319 'Sem parâmetros
'A quantidade do lote não pode ser maior que a quantidade do almoxarifado.
Public Const ERRO_QUANTALOCLOTE_MAIOR_QUUANTALOCALM = 11320 'Parâmetros: dQuantLote, dQuantAlm
'A quantidade alocada nos lotes %s tem que ser menor do que a quantidade alocada no Almoxarifado, que é %s.
Public Const ERRO_QUANTALOCADALOTE_MAIOR_QUANTALOCADAALM = 11321 'Parâmetros: iLinha
'A quantidade alocada do lote não pode ser maior que a quantidade alocada no almoxarifado. Rastreamento do produto, Linha = %i.
Public Const ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO = 11322
'Não foi encontrada transação na tabela TransacaoCTB correspondente ao Tipo de Nota Fiscal %s.
Public Const ERRO_LEITURA_RASTROESTINI = 11323 'Parametros sProduto, iAlmoxarifado
'Ocorreu um erro na leitura da tabela de RastroEstIni. Produto=%s, Almoxarifado=%i.
Public Const ERRO_EXCLUSAO_RASTROESTINI = 11324 'Parametros sProduto, iAlmoxarifado, iEscaninho, NumIntDocLote
'Ocorreu um erro na exclusão de um registro da tabela de RastroEstIni. Produto=%s, Almoxarifado=%i, Escaninho = %i, NumIntDocLote = %l.
Public Const ERRO_INCLUSAO_RASTROESTINI = 11325 'Parametros sProduto, iAlmoxarifado, iEscaninho, NumIntDocLote
'Ocorreu um erro na inclusão de um registro na tabela de RastroEstIni. Produto=%s, Almoxarifado=%i, Escaninho = %i, NumIntDocLote = %l.
Public Const ERRO_PRODUTO_ALTERACAO_RASTRO1 = 11326 'Parametro= sProduto
'Não é permitida a alteração do Tipo de Rastreamento do Produto %s pois já existem lotes de rastreamento.
Public Const ERRO_LEITURA_RASTREAMENTOLOTE1 = 11327 'Parametor = sProduto
'Erro na leitura da tabela RastreamentoLote. Produto = %s.
Public Const ERRO_PRODUTO_NAO_RASTRO = 11328 'Parâmetros: sProduto
'O produto %s não está indicado para rastreamento. Verifique o cadastro de produtos.
Public Const ERRO_QUANTESTINI_MENOR_RASTREAMENTO = 11329 'Parametros dQuantEstIni, dQuantRastreamento
'A quantidade do estoque inicial não pode ser menor do que o total de lotes de rastreamento. Quantidade Estoque Inicial = %d, Total Rastreamento = %d.
Public Const ERRO_QUANTTOTAL_LOTE_MAIOR_ESCANINHO = 11330 'Parâmetros: dQuantTotalLote, dQuantAlmoxarifado
'A quantidade total dos lotes %s não pode ser maior que a quantidade do escaninho que é %s.
Public Const ERRO_ESCANINHO_NAO_SELECIONADO = 11331 'Sem Parametros
'Nenhum escaninho foi selecionado.
Public Const ERRO_GRID_QUANTLOTE_NAO_PREENCHIDA = 11332 'Parametros: iLinhaGrid
'A quantidade alocada do grid não foi preenchida. Linha = %i.
Public Const ERRO_GRID_QUANTLOTE_ZERADA = 11333 'Parametros: iLinhaGrid
'A quantidade alocada do lote não pode ser zero. Linha = %i.
Public Const ERRO_LOTE_JA_UTILIZADO_GRID = 11334 'Parametros: sLote
'O Lote %s já foi utilizado no grid.
Public Const ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID = 11335 'Parametros: sLote, iFilialOP
'O Lote %s, FilialOP %i já foi utilizado no grid.
Public Const ERRO_LEITURA_ESCANINHOS = 11336 'Sem Parametros
'Ocorreu um erro na leitura da tabela de Escaninhos.
Public Const ERRO_PRODUTO_NAO_RASTROOP = 11337 'Parametro= sProduto
'O Produto %s não é rastreado por O.P. e portanto não pode ter uma filial de O.P. associada.
Public Const ERRO_PRODUTO_RASTROOP_FILIAL_ZERADA = 11338 'Parametro = sProduto
'O Produto %s é rastreado por O.P. e portanto a FilialOP deve ser preenchida.
Public Const ERRO_PRODUTO_AUSENTE_GRID_ITENS = 11339 'Parametro: sProduto
'O Produto %s não está presente no grid de Itens.
Public Const ERRO_PRODUTO_ALMOX_AUSENTE_GRID_ITENS = 11340 'Parametro: sProduto, sAlmoxarifado
'O Par Produto = %s, Almoxarifado = %s não está presente no grid de Itens.
Public Const ERRO_ALMOXARIFADO_AUSENTE_GRID_ALOCACAO = 11341 'Parametro: sAlmoxarifado, iItem
'O Almoxarifado %s não está presente no grid de Alocações para o item %i.
Public Const ERRO_PRODUTO_NAO_PREENCHIDO_GRID = 11342 'Parametro: iLinhaGrid
'O Campo de Produto não foi preenchido nesta linha do grid. Linha = %i.
Public Const ERRO_ALMOXARIFADO_NAO_PREENCHIDO_GRID = 11343 'Parametro: iLinhaGrid
'O Campo Almoxarifado não foi preenchido nesta linha do grid. Linha = %i.
Public Const ERRO_LOTE_PROD_ALMOX_JA_UTILIZADO_GRID = 11344 'Parametros: sLote, sProduto, sAlmoxarifado
'O Lote %s do Produto = %s para o Almoxarifado %s já foi utilizado no grid.
Public Const ERRO_LOTEOP_PROD_ALMOX_JA_UTILIZADO_GRID = 11345 'Parametros: sLote, sProduto, iFilialOP, sAlmoxarifado
'O Lote %s do Produto = %s, da FilialOP = %i para o Almoxarifado %s já foi utilizado no grid.
Public Const ERRO_ITEMRASTRO_NAO_ITEMNF = 11346 'Parametro: iLinhaGrid
'Este item não corresponde a nenhum dos itens da nota fiscal. Item = %i.
Public Const ERRO_GRIDRASTRO_ITEM_NAO_PREENCHIDO = 11347 'Parâmetros: iLinha
'No grid de Rastro o item da linha %i não foi preenchido.
Public Const ERRO_GRIDRASTRO_LOTE_NAO_PREENCHIDO = 11348 'Parâmetros: iLinha
'No grid de Rastro o lote da linha %i não foi preenchido.
Public Const ERRO_GRIDRASTRO_PRODUTO_NAO_PREENCHIDO = 11349 'Parâmetros: iLinha
'No grid de Rastro o produto da linha %i não foi preenchido.
Public Const ERRO_GRIDRASTRO_ALMOX_NAO_PREENCHIDO = 11350 'Parâmetros: iLinha
'No grid de Rastro o almoxarifado da linha %i não foi preenchido.
Public Const ERRO_GRIDRASTRO_ESCANINHO_NAO_PREENCHIDO = 11351 'Parâmetros: iLinha
'No grid de Rastro o escaninho da linha %i não foi preenchido.
Public Const ERRO_GRIDRASTRO_UM_NAO_PREENCHIDO = 11352 'Parâmetros: iLinha
'No grid de Rastro a Unidade de Medida da linha %i não foi preenchida.
Public Const ERRO_GRIDRASTRO_PRODUTO_INEXISTENTE = 11353 'Parâmetros: sProduto, iLinha
'No grid de Rastro o Produto %s da linha %i não está cadastrado.
Public Const ERRO_GRIDRASTRO_FILIALOP_NAO_PREENCHIDA = 11354 'Parâmetros: iLinha
'No grid de Rastro a Filial da O.P. da linha %i não foi preenchida.
Public Const ERRO_GRIDRASTRO_QUANT_NAO_PREENCHIDA = 11355 'Parâmetros: iLinha
'No grid de Rastro a Quantidade da linha %i não foi preenchida.
Public Const ERRO_GRIDRASTRO_QUANT_ZERADA = 11356 'Parâmetros: iLinha
'No grid de Rastro a Quantidade da linha %i está zerada.
Public Const ERRO_GRIDRASTRO_QUANT_MAIOR_ITEM = 11357 'Parametros: iItem, dQuantTotalRastro, dQuantItem
'A quantidade total rastreada para o item %i ultrapassou a quantidade do item. Quant. Rastreada = %d, Quant. Item = %d.
Public Const ERRO_PRODUTO_NAO_PREENCHIDO_GRID_ITENS = 11358 'Parametro: iItem
'O Produto não foi preenchido na linha %i do grid de itens.
Public Const ERRO_PRODUTO_RASTRO_DIF_ITEMNF = 11359 'Parametro: sProduto, iItem
'O Produto %s referente ao Item %i do rastreamento não corresponde ao produto correspondente do item da nota fiscal gravada = %s.
Public Const ERRO_ALMOX_RASTRO_DIF_ITEMNF = 11360 'Parametro: iAlmoxRastro, iItemRastro, iAlmoxItemNF
'O Almoxarifado %i referente ao Item %i do rastreamento não corresponde ao almoxarifado do item da nota fiscal gravada = %i.
Public Const ERRO_SIGLAUM_RASTRO_DIF_ITEMNF = 11361 'Parametro: sSiglaUMRastro, iItemRastro, sSiglaUMItemNF
'A Unidade de Medida %s referente ao Item %i do rastreamento não corresponde a unidade de medida do item da nota fiscal gravada = %s.
Public Const ERRO_QUANT_RASTRO_DIF_ITEMNF = 11362 'Parametro: dQuantRastro, iItemRastro, dQuantItemNF
'A Quantidade %d referente ao Item %i do rastreamento não corresponde a quantidade do item da nota fiscal gravada = %s.
Public Const ERRO_QUANT_RASTRO_MAIOR_ITEMNF = 11363 'Parametro: iItemNF, dQuantTotalRastro, dQuantItemNF
'A Quantidade total do rastreamento para o item %i da Nota Fiscal ultrapassa a quantidade gravada do item. Quant.Total Rastro = %d, Quant. ItemNF = %d.
Public Const ERRO_NUM_ITEMNF_DIF_TELA = 11364 'Parametro: iQuantItemNFGravado, iQuantItemNFTela
'O número de itens da nota fiscal gravada difere do número de itens da tela. Total Gravado = %i, Total Tela = %i.
Public Const ERRO_ITEMNF_PRODUTO_DIF_TELA = 11365 'Parametro: sItemNFProdutoTela, iItemTela, sItemNFProdutoGravado
'O Produto %s referente ao Item da Nota Fiscal %i da tela não corresponde ao produto correspondente do item da nota fiscal gravada = %s.
Public Const ERRO_ITEMNF_ALMOX_DIF_TELA = 11366 'Parametro: iItemNFAlmoxTela, iItemTela, iItemNFAlmoxGravado
'O Almoxarifado %i referente ao Item da Nota Fiscal %i da tela não corresponde ao almoxarifado correspondente do item da nota fiscal gravada = %i.
Public Const ERRO_ITEMNF_UM_DIF_TELA = 11367 'Parametro: sItemNFSiglaUMTela, iItemTela, sItemNFSiglaUMGravado
'A Unidade de Medida %s referente ao Item da Nota Fiscal %i da tela não corresponde à unidade de medida correspondente do item da nota fiscal gravada = %s.
Public Const ERRO_ITEMNF_QUANT_DIF_TELA = 11368 'Parametro: dItemNFQuantUMTela, iItemTela, dItemNFQuantGravado
'A Quantidade %d referente ao Item da Nota Fiscal %i da tela não corresponde à quantidade correspondente do item da nota fiscal gravada = %d.
Public Const ERRO_GRIDRASTRO_QUANT_MAIOR_ALOC = 11369 'Parametros: sProduto, sAlmoxarifado, iLinhaAlocacao, dQuantTotalRastro, dQuantAloc
'A quantidade total rastreada para o produto %s e almoxarifado %s ultrapassou a quantidade alocada. Quant. Rastreada = %d, Quant. Alocada = %d.
Public Const ERRO_RASTRO_NAO_UTILIZADO = 11370 'Parametros: iItemRastro
'O Item %i do rastreamento não está sendo utilizado pois não há uma alocação de material para o produto, almoxarifado e escaninho escolhidos no banco de dados.
Public Const ERRO_LEITURA_NFISCAL_ITENSNF_MOVESTOQUE = 11371 'Parametros: giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, iStatus, iTipoNF, iItem, iTipoNumIntDocOrigem, iTipoMovEstoque)
'Ocorreu um erro na leitura do relacionamento das tabelas NFiscal, ItensNFiscal e MovimentoEstoque. Filial = %i, Série da NF = %s, Numero NF = %l, Data Emissão NF = %dt, Status da Nota <> %i, Tipo da NF = %i, Item da NF = %i, TipoNumIntDocOrigem(MovEstoque) = %i, Tipo Mov. Estoque = %i.
Public Const ERRO_ITEMRASTRO_NAO_ALOCACAO = 11372 'Parametro: iLinhaGrid
'Este item não corresponde a nenhum dos itens da localização (alocação). Item = %i.
Public Const ERRO_NFISCAL_MOVESTOQUE_INEXISTENTE = 11373 'Parametros: iItemNF, objAlmoxarifado.sNomeReduzido
'Não existe movimento de estoque cadastrado para este item de nota fiscal neste almoxarifado. Item N.F. = %i, Almoxarifado = %s.
Public Const ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO = 11374
'O Número de Nota Fiscal não foi preenchido.
Public Const ERRO_NF_MOVESTOQUE_DISP_INEXISTENTE = 11375 'Parametros: iItemNF, sAlmoxarifado
'Não existe movimento de estoque de material disponível para este item de nota fiscal neste almoxarifado. Item N.F. = %i, Almoxarifado = %s.
Public Const ERRO_NF_MOVESTOQUE_CONSIG_INEXISTENTE = 11376 'Parametros: iItemNF, sAlmoxarifado
'Não existe movimento de estoque de material consignado para este item de nota fiscal neste almoxarifado. Item N.F. = %i, Almoxarifado = %s.
Public Const ERRO_TIPODOCINFO_NAO_RASTREAVEL = 11377 'Sem Parametros
'Este tipo de documento não pode ter rastreamento associado.
Public Const ERRO_QUANT_RASTRO_DIF_ALOCADO = 11378 'Parametro sProduto, sAlmoxarifadoNomeRed, dQuantidadeItemMovEst, dQuantidadeTotalRastro
'A quantidade alocada para o produto/almoxarifado é diferente da quantidade rastreada. Produto = %s, Almoxarifado = %s, Quantidade Alocada = %d, Quant. Total Rastreada = %d.
Public Const ERRO_QUANT_RASTRO_MAIOR_ALOC = 11379 'Parametros: sProduto, iAlmoxarifado, dQuantItemMovEstoque, dQuantTotalRastro
'A quantidade alocada para o produto/almoxarifado é maior do que a quantidade rastreada. Produto = %s, Almoxarifado = %s, Quantidade Alocada = %d, Quant. Total Rastreada = %d.
Public Const ERRO_QUANT_RASTRO_MENOR_ALOC = 11380 'Parametros: sProduto, iAlmoxarifado, dQuantItemMovEstoque, dQuantTotalRastro
'A quantidade alocada para o produto/almoxarifado é menor do que a quantidade rastreada. Produto = %s, Almoxarifado = %s, Quantidade Alocada = %d, Quant. Total Rastreada = %d.
Public Const ERRO_ARQUIVO_INVALIDO = 11381 'Parametro: NomeFigura.text
' %s não é o nome de um arquivo.
Public Const ERRO_EMBALAGEM_NAO_ENCONTRADA = 11382 'Parametro: objEmbalagem.iCodigo
'A embalagem com o código %s não está cadastrada.
Public Const ERRO_PREENCH_CAMPOS_IDENTIFICACAO = 11383 'Sem Parametros
'Em produtos que podem ser vendidos, é obrigatório o preenchimentos de pelo menos um dos campos: Código de barras ou Referência.
Public Const ERRO_PREENCH_ICMS = 11384 'Sem parametros
'Em produtos que podem ser vendidos e que o tipo de situação tributária  for Tributado ou Substituição Tributária  é obrigatório o preenchimento da alíquota ICMS.
Public Const ERRO_CODBARRAS_OU_REFERENCIA_PREENCH_OBRIGATORIOS = 11385 'sem parametros
'É obrigatório o preenchimento de um dos campos para produtos gerenciais e de vendas: Código de barras ou Referência.
Public Const ERRO_LEITURA_EMBALAGEM = 11386
'Erro na leitura da tabela embalagens
Public Const ERRO_CAPACIDADE_NAO_PREENCHIDA = 11387 'sem parametros
'O preenchimento da capacidade é obrigatório.
Public Const ERRO_EMBALAGEM_VINCULADA_PRODUTO = 11388 'parametro:iCodigo
'A embalagem com código %i não pode ser excluída pois ela é
'embalagem padrão de produto.
Public Const ERRO_EXCLUSAO_EMBALAGEM = 11389 'parametro:icodigo
'Erro na exclusão da embalagem com código %i.
Public Const ERRO_PESO_NAO_PREENCHIDO = 11390 'sem parametros
'O preenchimento do peso é obrigatório.
Public Const ERRO_LOCK_EMBALAGEM = 11391 'parametro:icodigo
'Erro na tentativa de lock na embalagem com código %i.
Public Const ERRO_ATUALIZACAO_EMBALAGEM = 11392 'parametro:icodigo
'Erro na atualização da Embalagem com código %i.
Public Const ERRO_INSERCAO_EMBALAGEM = 11393 'parametro:icodigo
'Erro na inserção da Embalagem com código %i.
Public Const ERRO_DESC_EMBALAGEM_IGUAL = 11394 'parametro: icodigo
'A descrição dessa embalagem é a mesma usada pela embalagem com código %i.
Public Const ERRO_EMBALAGEM_NAO_CADASTRADA = 11395 'parametro:icodigo
'A embalagem com código %i não está cadastrada.
Public Const ERRO_SIGLA_NAO_PREENCHIDA = 11396 'sem parametros
'O preenchimento da sigla é obrigatório.
Public Const ERRO_OPBAIXADA_NAO_REATIVADA = 11397 'sem parametros
'A ordem de produção está baixada e nenhum item está com Situação= Normal no grid.
Public Const ERRO_ATUALIZACAO_ITENSORDEMPRODUCAOBAIXADAS = 11398 'sem parametros
'Erro na tentativa de atualização na tabela de Itens de Ordem de Produção Baixadas
Public Const ERRO_EXCLUSAO_ITEMOPBAIXADA = 11399 'sem parametros
'Erro na tentativa de exclusão na tabela de Itens de Ordem de Produção Baixadas


'Códigos de Erros - Reservado de 7300 até 7999 ; 8500 até 8999
Public Const ERRO_LEITURA_CATEGORIAPRODUTOITEM2 = 7300 'Parâmetro: sCategoria, sItem
'Erro na leitura do registro da categoria %s, do item %s, da tabela itens das categorias de Produto.
Public Const ERRO_LOCK_CATEGORIAPRODUTOITEM = 7301  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de itens das Categorias de Produtos.
Public Const ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE = 7302 'Parametro: sItem, sCategoria
'O Item %s da Categoria %s de Produto não existe.
Public Const ERRO_CATEGORIAPRODUTO1_INEXISTENTE = 7303  'Parametro: sCategoria
'A Categoria %s de Produto não existe.
Public Const ERRO_CATEGORIAPRODUTO_UTILIZADA = 7308 'parametros categoria e produto
'A categoria %s não pode ser excluída pois é utilizada por produtos como %s.
Public Const ERRO_CATEGORIAPRODUTO_NAO_INFORMADA = 7309 'Sem parametro
'A Categoria deve ser informada.
Public Const ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO = 7310 'Sem parametro
'Se a opção 'Todas' não estiver selecionada, deve ser preenchido a Categoria e os Itens a serem selecionados.
Public Const ERRO_FALTA_ITEM_CATEGORIAPRODUTO = 7311 'Sem parametro
'Somente a Descrição do Item foi informada. O Item não foi informado.
Public Const ERRO_LEITURA_CATEGORIAPRODUTO = 7312 'Sem parametro
'Erro na leitura da tabela de Categorias de Produtos.
Public Const ERRO_LEITURA_CATEGORIAPRODUTOITEM = 7313 'Sem parametro
'Erro na leitura da tabela de Itens de Categorias de Produtos.
Public Const ERRO_INSERCAO_CATEGORIAPRODUTOITEM = 7314 'Sem parametro
'Erro na inserção de registro na tabela de Itens de Categorias de Produtos.
Public Const ERRO_MODIFICACAO_CATEGORIAPRODUTO = 7315 'Sem parametro
'Erro na modificação da tabela de Categorias de Produtos.
Public Const ERRO_MODIFICACAO_CATEGORIAPRODUTOITEM = 7316 'Sem parametro
'Erro na modificação da tabela de Itens de Categorias de Produtos.
Public Const ERRO_EXCLUSAO_CATEGORIAPRODUTOITEM = 7317 'Sem parametro
'Erro na exclusão de registro na tabela de Itens de Categorias de Produtos.
Public Const ERRO_LEITURA_CATEGORIAPRODUTOITENS_CATEGORIA = 7318 'Sem parametro
'Erro na leitura de itens de uma Categoria.
Public Const ERRO_INSERCAO_CATEGORIAPRODUTO = 7319 'Sem parametro
'Erro na inserção de uma Categoria de Produto.
Public Const ERRO_CATEGORIAPRODUTO_INEXISTENTE = 7320 'Sem parametro
'A Categoria de Produto não existe.
Public Const ERRO_LOCK_CATEGORIAPRODUTO = 7321  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de Categorias de Produtos.
Public Const ERRO_EXCLUSAO_CATEGORIAPRODUTO = 7322  'Sem parametro
'Erro na exclusão de uma Categoria de Produto.
Public Const ERRO_LEITURA_PRODUTOS_CATEGORIA = 7323  'Parametro: Categoria
'Erro na leitura dos produtos da Categoria %s.
Public Const ERRO_LEITURA_PRODUTOS_CATEGORIA_ITEM = 7324  'Parametro: Categoria e Item
'Erro na leitura de produtos da Categoria %s com valor %s.
Public Const ERRO_CATEGORIAPRODUTOITEM_UTILIZADO = 7325  'Parametro: Produto, Categoria e Item
'O Produto %s está associado ao ítem %s da Categoria %s.
Public Const ERRO_LEITURA_ALMOXARIFADO = 7326 'Parametro: iCodigo
'Erro na leitura do Almoxarifado com código %i.
Public Const ERRO_LEITURA_ALMOXARIFADO1 = 7327 'Parametro: sNomeReduzido
'Erro na leitura do Almoxarifado %s.
Public Const ERRO_ALMOXARIFADO_NAO_NORMAL = 7330 'Parametros: iCodigo, iTipo
'Almoxarifado %i não é do tipo normal. Tipo do Almoxarifado: %i.
Public Const ERRO_LEITURA_CATEGS_PROD = 7331  'Parametro: sProduto
'Erro na leitura do Produto %s na tabela de Categorias de Produtos.
Public Const ERRO_LEITURA_ESTOQUEPRODUTO1 = 7332 'Parâmetros: sProduto, iAlmoxarifado
'Erro na leitua da tabela EstoqueProduto com Produto %s e Almoxarifado %i.
Public Const ERRO_LEITURA_ESTOQUESPRODUTO = 7333 'Sem parametros
'Erro na leitura da tabela EstoqueProduto
Public Const ERRO_NAO_EXISTE_ESTOQUE = 7334 'Parametro sCodProduto
'Não existe estoque do Produto %s.
Public Const ERRO_NAO_EXISTE_ALMOXARIFADO = 7335 'Parametro: sCodProduto
'O produto %s não está associado a nenhum almoxarifado.
Public Const ERRO_ALMOXARIFADO_NAO_TEM_PRODUTO = 7336 'Parametros: iCodAlmoxarifado, sCodProduto
'Almoxarifado %i não trabalha com o Produto %s.
Public Const ERRO_PRIMEIRA_LINHA_NAO_PODE_SER_EXCLUIDA = 7337 'Sem parametro
'A primeira linha do Grid não pode ser excluída.
Public Const ERRO_CLASSEUM_INEXISTENTE = 7338 'Parametro: iClasse
'A ClasseUM, com código %i não foi encontrada.
Public Const ERRO_LOCK_CLASSEUM = 7339 'Sem parametro
'Erro na tentativa de fazer "lock" na tabela ClasseUM.
Public Const ERRO_INSERCAO_CLASSEUM = 7340 'Parametro: código da Classe
'Erro na tentativa de inserir a ClasseUM com código %i.
Public Const ERRO_EXCLUSAO_CLASSEUM = 7341 'Parametro: código da Classe
'Erro na tentativa de excluir a ClasseUM com código %i.
Public Const ERRO_INSERCAO_UNIDADESDEMEDIDA = 7342 'Parametro: código da Classe
'Erro na tentativa de inserir a ClasseUM com código %i, na tabela UnidadesDeMedida.
Public Const ERRO_EXCLUSAO_UNIDADESDEMEDIDA = 7343 'Parametro: código da Classe
'Erro na tentativa de excluir a UM da ClasseUM com código %i.
Public Const ERRO_UM_REPETIDA = 7344 'Sem parametro
'Não pode haver repetição de unidades no Grid.
Public Const ERRO_LEITURA_CLASSEUM = 7345 'Sem parametro
'Erro na leitura da tabela ClasseUM.
Public Const ERRO_MODIFICACAO_CLASSEUM = 7346 'Sem parametro
'Erro na modificação da tabela ClasseUM.
Public Const ERRO_LEITURA_UNIDADESDEMEDIDA = 7347 'Sem parametro
'Erro na leitura da tabela UnidadesDeMedidas.
Public Const ERRO_LEITURA_ITENSPEDIDODEVENDA = 7349 'Sem parametro
'Erro na leitura da tabela ItensPedidoDeVenda.
Public Const ERRO_CLASSEUM_UTILIZADA_PRODUTOS = 7350 'Parametro: iClasse
'A Classe está sendo utilizada em Produtos .
Public Const ERRO_CLASSEUM_E_SIGLAUM_UTILIZADAS_PRODUTOS = 7351 'Parametros: iClasse e sSigla
'A Classe e a Sigla estão sendo utilizadas em Produtos .
Public Const ERRO_CLASSEUM_E_SIGLAUM_UTILIZADAS_ITENSPEDIDODEVENDA = 7352 'Parametros: iClasse e sSigla
'A Classe e a Sigla estão sendo utilizadas em ItensPedidoDeVenda .
Public Const ERRO_CODIGO_CLASSEUM_NAO_PREENCHIDO = 7353 'Sem parametro
'O Código da Classe UnidadeDeMedida deve ser preenchido.
Public Const ERRO_DESCRICAO_CLASSEUM_NAO_PREENCHIDA = 7354 'Sem parametro
'A Descrição da Classe UnidadeDeMedidadeve ser preenchida.
Public Const ERRO_SIGLA_CLASSEUM_NAO_PREENCHIDA = 7355 'Sem parametro
'A Sigla da Classe UnidadeDeMedidadeve ser preenchida.
Public Const ERRO_QUANTIDADE_CLASSEUM_NAO_PREENCHIDA = 7356 'Sem parametro
'A Quantidade da Classe UnidadeDeMedidadeve ser preenchida.
Public Const ERRO_SIGLAUMBASE_NAO_PREENCHIDA = 7357 'Sem parametro
'A SiglaBASE da Classe UnidadeDeMedidadeve ser preenchida.
Public Const ERRO_NOMEUMBASE_NAO_PREENCHIDO = 7358 'Sem parametro
'O Nome da Classe UnidadeDeMedidadeve ser preenchido.
Public Const ERRO_LEITURA_TIPODEPRODUTOCATEGORIAS1 = 7359 'Parâmetro: sCategoria
'Erro de leitura da tabela TipoDeProdutoCategorias Com Categoria %s.
Public Const ERRO_EXCLUSAO_CATEGORIAPRODUTO_UTILIZADA = 7360 'Parâmetro: sCategoria, iTipo
'Não é permitido excluir a Categoria %s porque está sendo utilizada pelo Tipo de Produto %i.
Public Const ERRO_CATEGORIAPRODUTOITEM_NAO_CADASTRADA = 7361 'Parametro: sItem e sCategoria
'O Item %s da Categoria %s não está cadastrado.
Public Const ERRO_UM_NAO_CADASTRADA = 7362 'Parametro: Classe de UM
'A Unidade de Medida, cuja Classe é %i, não está cadastrada.
Public Const ERRO_LOCK_UNIDADESDEMEDIDA = 7363 'Sem parametro
'Erro na tentativa de fazer "lock" na Tabela UnidadesDeMedida.
Public Const ERRO_INSERCAO_TIPOPRODUTO = 7364 'Parametro: iTipo
'Erro na tentativa de inserir o Tipo de Produto %i.
Public Const ERRO_INSERCAO_TIPOPRODUTOCATEGORIA = 7365 'Parametro: iTipo
'Erro na tentativa de inserir o Tipo de Produto %i.
Public Const ERRO_EXCLUSAO_TIPODEPRODUTOCATEGORIAS = 7366 'Parametro: Tipo do Produto
'Erro na tentativa de excluir a Categoria do Tipo de Produto %i.
Public Const ERRO_TIPOPRODUTO_UTILIZADO_PRODUTOS = 7367 'Parametro: iTipo
'O Tipo de Produto está sendo utilizado em Produtos .
Public Const ERRO_EXCLUSAO_TIPOPRODUTO = 7368 'Parametro: iTipo
'Erro na tentativa de excluir o Tipo de Produto com código %i.
Public Const ERRO_TIPOPRODUTO_INEXISTENTE = 7369 'Parametro: iTipo
'O Tipo de Produto, com código %i não foi encontrado.
Public Const ERRO_LOCK_TIPOSDEPRODUTO = 7370 'Sem parametro
'Erro na tentativa de fazer "lock" na Tabela TiposDeProduto.
Public Const ERRO_TIPOPRODUTO_NAO_CADASTRADO = 7371 'Parametro: Código do Tipo de Produto
'O código do Tipo de Produto %s não está cadastrado.
Public Const ERRO_CODIGO_TIPOPRODUTO_NAO_PREENCHIDO = 7372 'Sem parametro
'O código do Tipo de Produto deve ser informado.
Public Const ERRO_DESCRICAO_TIPOPRODUTO_NAO_PREENCHIDA = 7373 'Sem parametro
'A descrição do Tipo de Produto deve ser informada.
Public Const ERRO_CATEGORIAPRODUTO_REPETIDA_NO_GRID = 7374 'Sem parametro
'Não pode haver repetição de Categorias de Produto no Grid.
Public Const ERRO_LEITURA_TIPOSDEPRODUTO = 7375 'Parâmetro: iTipo
'Erro na leitura do Tipo %i na tabela de Tipos de Produto.
Public Const ERRO_MODIFICACAO_TIPOSDEPRODUTO = 7376 'Parâmetro: iTipo
'Erro na modificação da tabela TiposDeProduto do Tipo de Produto %i.
Public Const ERRO_LEITURA_TIPODEPRODUTOCATEGORIAS = 7377 'Parâmetro: iTipo
'Erro na leitura da tabela TipoDeProdutoCategorias com o Tipo de Produto = %i.
Public Const ERRO_MODIFICACAO_TIPODEPRODUTOCATEGORIAS = 7378 'Sem parametro
'Erro na modificação da tabela TipoDeProdutoCategorias.
Public Const ERRO_LEITURA_CLASSEUM1 = 7379 'Sem parâmetros
'Erro na leitura da tabela ClasseUM.
Public Const ERRO_UNIDADES_MEDIDAS_NAO_CADASTRADAS = 7380 'Sem Parâmetros
'As Unidades de Medida de Estoque, Compra e Venda não estão cadastradas.
Public Const ERRO_LEITURA_UNIDADESDEMEDIDA1 = 7381 'Parâmetro: iClasse
'Erro na leitura da Tabela de Unidades de Medidas onde a Classe = %i.
Public Const ERRO_CLASSEUM_SIGLAUM_INEXISTENTE = 7382 'Parametro: iClasse, sSigla
'A unidade de medida da classe %i e sigla %s não foi encontrada.
Public Const ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO = 7383 'Sem parametros
'Preenchimento do código do produto raiz do kit é obrigatório.
Public Const ERRO_CODIGO_PRODUTOKIT_NAO_PREENCHIDO = 7384 'Sem parametros
'Preenchimento do código do produto constituinte do kit é obrigatório.
Public Const ERRO_VERSAO_KIT_NAO_PREENCHIDA = 7385 'Sem parametros
'Preenchimento da versão do kit é obrigatório.
Public Const ERRO_LEITURA_PRODUTOKIT = 7386 'Sem parametros
'Erro na leitura da tabela de itens de um kit de produtos.
Public Const ERRO_EXCLUSAO_PRODUTOKIT = 7387 'Sem parametros
'Erro na exclusão de registro da tabela de itens de um kit de produtos.
Public Const ERRO_ATUALIZACAO_PRODUTOKIT = 7388 'Sem parametros
'Erro na atualização de registro da tabela de itens de um kit de produtos.
Public Const ERRO_ATUALIZACAO_KIT = 7389 'Sem parametros
'Erro na atualização de registro da tabela de kits de produtos.
Public Const ERRO_INSERCAO_PRODUTOKIT = 7390 'Sem parametros
'Erro na inclusão de registro da tabela de itens de um kit de produtos.
Public Const ERRO_INSERCAO_KIT = 7391 'Sem parametros
'Erro na inclusão de registro da tabela de kits de produtos.
Public Const ERRO_KIT_INEXISTENTE = 7392 'Parametros: sVersao, sCodigo
'Kit com versão %s do produto de código %s não cadastrado na tabela de kits.
Public Const ERRO_PRODUTOKIT_INEXISTENTE = 7393 'Parametros: sVersao, sCodigo
'Kit com versão %s do produto de código %s não cadastrado na tabela de itens constituintes de um kit.
Public Const ERRO_PRODUTO_JA_FAZ_PARTE_CAMINHO_KIT = 7394 'Sem parametros
'Não é possível incluir produto nesta posição pois este já se encontra no caminho especificado.
Public Const ERRO_INCLUIR_PRODUTO_NAOBASICO_NAOINTERMEDIARIO = 7395 'Sem parametros
'Não é possível incluir produtos que não sejam nem básicos nem intermediários.
Public Const ERRO_PRODUTOPAI_TEM_KIT = 7396 'Parametros: sCodigo1, sCodigo2
'Não é possível incluir produto %s, pois o produto %s (seu antecessor na árvore) já tem kit.
Public Const ERRO_ATINGIU_NIVEL_MAXIMO = 7397 'Parametros: iNivelMaximo
'Não é possível incluir produto neste ponto, pois a árvore atingiu seu número máximo de níveis(%i).
Public Const ERRO_SLDMESEST_STATUS_FECHADO = 7398 'Parametros: iMes, iAno
'O mes %i/%i está fechado. Não é possível incluir um movimento de estoque.
Public Const ERRO_LEITURA_TIPOSMOVEST = 7399 'Parametro iCodigo
'Ocorreu um erro na leitura da tabela de tipos de movimento de estoque. Codgio=%i.
Public Const ERRO_LEITURA_ESTOQUEPRODUTO = 7400 'Parametros sProduto, iAlmoxarifado
'Ocorreu um erro na leitura da tabela de EstoqueProduto. Produto=%s, Almoxarifado=%i.
Public Const ERRO_LOCK_ESTOQUEPRODUTO = 7401  'Parametros sProduto, iAlmoxarifado
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de EstoqueProduto. Produto=%s, Almoxarifado=%i.
Public Const ERRO_ATUALIZACAO_ESTOQUEPRODUTO = 7402  'Parametros sProduto, iAlmoxarifado
'Ocorreu um erro na tentativa de atualizar o Produto %s, Almoxarifado %i na tabela de EstoqueProduto.
Public Const ERRO_LEITURA_SLDDIAEST = 7403 'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na leitura da tabela de saldos diários de estoque. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_INSERCAO_SLDDIAEST = 7404 'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na inclusão de registro na tabela de saldos diários de estoque. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_LOCK_SLDDIAEST = 7405  'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos diários de estoque. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_ATUALIZACAO_SLDDIAEST = 7406  'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos diários de estoque. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_ATUALIZACAO_SLDMESEST = 7407  'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque. Ano=%i, FilialEmpresa=%i, Produto=%s.
Public Const ERRO_INSERCAO_ESTOQUEMOVIMENTO = 7408 'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na inclusão de registro na tabela de saldos diários de estoque. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_LEITURA_MATCONFIG = 7409 'Parametro sCodigo
'Ocorreu um erro na leitura da tabela de Configuração de Materiais (MATConfig). Codigo=%s.
Public Const ERRO_LOCK_MATCONFIG = 7410  'Parametro sCodigo
'Ocorreu um erro na tentativa de fazer 'lock' na tabela Configuração de Materiais (MATConfig). Codigo=%s.
Public Const ERRO_ATUALIZACAO_MATCONFIG = 7411  'Parametro sCodigo
'Ocorreu um erro na tentativa de atualizar um registro na tabela de Configuração de Materiais (MATConfig). Codigo=%s.
Public Const ERRO_INSERCAO_MOVIMENTOESTOQUE = 7412 'Parametros iFilialEmpresa, lCodigo
'Ocorreu um erro na inclusão de registro na tabela de movimentos de estoque. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_LOCK_MOVIMENTOESTOQUE = 7413  'Parametros iFilialEmpresa, lCodigo
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de movimentos de estoque. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_LEITURA_MOVIMENTOESTOQUE1 = 7414 'Parametros iFilialEmpresa, lCodigo
'Ocorreu um erro na leitura de um registro da tabela de movimentos de estoque. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_ATUALIZACAO_MOVIMENTOESTOQUE = 7415  'Parametros iFilialEmpresa, lCodigo
'Ocorreu um erro na tentativa de atualizar um registro da tabela de Movimentos de Estoque. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_EXCLUSAO_MOVIMENTOESTOQUE = 7416  'Parametros iFilialEmpresa, lCodigo
'Ocorreu um erro na tentativa de excluir um registro da tabela de Movimentos de Estoque. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_EXCLUSAO_ESTOQUEPRODUTO = 7417 'Parametros: Produto e Almoxarifado
'Ocorreu um erro na tentativa de excluir um registro da tabela de Estoque de Produtos. Produto=%i, Almoxarifado=%l.
Public Const ERRO_MOVIMENTOESTOQUE_NAO_CADASTRADO = 7418  'Parametros iFilialEmpresa, lCodigo
'O Movimento de Estoque não está cadastrado. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_MOVIMENTOESTOQUE_ESTORNADO = 7419  'Parametros iFilialEmpresa, lCodigo
'O Movimento de Estoque está estornado. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_SALDO_MAT_CONSERTO = 7420 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto em conserto é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %, Saldo = %d.
Public Const ERRO_SALDO_MAT_CONSERTO3 = 7421 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto de terceiros em conserto é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_CONSIG = 7422 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto em consignação é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_CONSIG3 = 7423 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto de terceiros em consignação é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_DEFEITUOSO = 7424 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto defeituoso é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, , Saldo = %d.
Public Const ERRO_SALDO_MAT_DEMO = 7425 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto em demonstração é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_DEMO3 = 7426 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto de terceiros em demonstração é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_DISPONIVEL = 7427 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto disponível é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_INDISPONIVEL = 7428 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto indisponível é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_OUTRAS = 7429 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto em poder de terceiros é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_OUTRAS3 = 7430 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto de terceiros é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_MOVIMENTOESTOQUE_JA_CADASTRADO = 7431  'Parametros iFilialEmpresa, lCodigo
'A Filial/Codigo deste(s) Movimento(s) de Estoque já está(ão) cadastrados. FilialEmpresa=%i, Codigo=%l.
Public Const ERRO_LEITURA_MOVIMENTOESTOQUE2 = 7432 'Parametros iFilialEmpresa
'Ocorreu um erro na leitura de um registro da tabela de movimentos de estoque. FilialEmpresa=%i.
Public Const ERRO_MOVIMENTOESTOQUE_DATA = 7433 'Parametros sUltima_Data, sData_do_Movimento
'Não é possível cadastrar um movimento cuja data seja menor que a data do ultimo movimento cadastrado. Data do Ultimo Movimento Cadastrado = %s, Data do Movimento = %s.
Public Const ERRO_PRODUTO_SEM_ESTOQUE = 7434 'Parametro: sProduto
'O Produto %s não é um produto de Estoque.
Public Const ERRO_KIT_APENAS_COM_PRODUTORAIZ = 7435 'Sem Parametro
'Um Kit não pode conter apenas o Produto Raiz.
Public Const ERRO_ALMOXARIFADO_OUTRA_FILIAL = 7436 'Parametro: sAlmoxarifado
'O Almoxarifado %s está localizado em outra Filial da Empresa.
Public Const ERRO_SALDO_MAT_DISP_CONSIG3 = 7437 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto disponível + o saldo de terceiros em consignação é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_RESDISP_RESCONSIG3 = 7438 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto Reservado disponível + o saldo de terceiros reservado em consignação é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const Erro_EstoqueMes_Fechado = 7440 'Parametros iFilialEmpresa, iAno, iMes
'O Estoque está fechado para o mês em questão. FilialEmpresa = %i, Ano = %i, Mes = %i.
Public Const ERRO_LEITURA_ESTOQUEMES1 = 7442 'Sem parametros
'Erro na leitura da tabela EstoqueMes.
Public Const ERRO_ESTOQUEMES_ABERTO = 7443 'Parametros iFilialEmpresa, iAno, iMes.
'O Estoque está aberto para o mês em questão. FilialEmpresa = %i, Ano = %i, Mes = %i.
Public Const ERRO_LOCK_ESTOQUEMES = 7444 'Parametros iFilialEmpresa, iAno, iMes.
'Erro na tentativa de lock na tabela EstoqueMes. FilialEmpresa = %i, Ano = %i, Mes = %i.
Public Const ERRO_PRODUTO_APROPR_CUSTO_REAL_INEXISTENTE = 7445 'Sem parametros
'Não está cadastrado nenhum Produto com apropriação custo real de produção.
Public Const ERRO_NENHUM_ITEMOP_INFORMADO = 7446 'Sem parametros
'A ordem de produção está vazia. Preencha-a com pelo menos 1 item.
Public Const ERRO_ALMOXARIFADO_NAO_PREENCHIDO = 7448 ' Parametro iLinhaGrid
'O Almoxarifado do ítem %i do Grid não foi preenchido.
Public Const ERRO_ALMOXARIFADO_INEXISTENTE2 = 7449 'Parametro: iCodigo
'O Almoxarifado %i não está cadastrado.
Public Const ERRO_DATAFIMOP_ANTERIOR_DATAINICIOOP = 7450 'Parâmetros: sDataFim, sDataInicio, iItem
'A Data de Previsão de Término da Produção = %s é anterior a Data de Previsão de Início = %s. Item = %i.
Public Const ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA = 7451 'Parametro: sCodigo
'Já existe uma ordem de produção cadastrada com o codigo %s.
Public Const ERRO_MOVIMENTO_NAO_REQCONSUMO = 7452 'Parametro lCodMovEstoque
'Movimento %l não é Requisição de Material para Consumo.
Public Const ERRO_ALMOXARIFADO_INEXISTENTE1 = 7453 'Parametro sAlmoxarifadoNomeReduzido
'Almoxarifado %s não está cadastrado.
Public Const ERRO_CUSTO_STANDARD_NAO_CADASTRADO = 7454 'Parametros sCodProduto e iFilialEmpresa
'O Custo Standard para o Produto %s e Filial %i não foi cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_TABELA_PRODUTOSFILIAL = 7455 'Sem parametros
'Erro de leitura na tabela ProdutosFilial
Public Const ERRO_QUANTIDADE_REQ_MAIOR = 7456 'Sem parametros
'Quantidade requisitada é maior do que a Quantidade Disponível.
Public Const ERRO_LEITURA_TIPOMOVEST = 7457 'Sem Parametros
'Erro de Leitura na Tabela TiposMovimentoEstoque.
Public Const ERRO_TIPOMOVEST_INEXISTENTE = 7458 'Parâmetro iCodigoTipoMovEst
'Tipo de Movimento de Estoque com Código %i não está cadastrado no Banco de Dados.
Public Const ERRO_EXCLUSAO_ITEM_GRAVADO = 7459 'Sem Parametros
'Não é possível excluir ítem gravado no Banco de Dados.
Public Const ERRO_MOVIMENTOESTOQUE_INEXISTENTE = 7460 'Parametros iFilialEmpresa,lCodigoMovEstoque, lNumIntDocEst
'Não existe cadastrado no Banco de Dados Movimento de Estoque na FilialEmpresa %i com Código %l e Documento de Estorno %l.
Public Const ERRO_LOCK_ALMOXARIFADO = 7461  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de Almoxarifados.
Public Const ERRO_SIGLA_UM_NAO_EDITAVEL = 7462 'Parâmetros: iClasse, sSiglaUMBase
'Não é permitida a alteração da Sigla %s da Classe %i. Esta Sigla foi utilizada em algum movimento de estoque ou é uma Sigla Padrão desta Classe.
Public Const ERRO_LOCK_ITEMOP = 7463 'Parâmetro sOPCodigo, sProdutoOP
'Erro na tentativa de fazer o "lock" do Item de Ordem de Produção com Código %s e Código do Produto %s.
Public Const ERRO_ITEMOP_NAO_CADASTRADO = 7464 'Parametros sOPCodigo,sCodProduto
'Ítem de Ordem de Produção não está cadastrado no Banco de Dados. Código da Ordem de Produção: %s, Código do Produto :%s.
Public Const ERRO_ORDEMPRODUCAO_SEM_ITENS = 7465 ' Parametros sOPCodigo, iFilialEmpresa
'Ordem de Produção com código %s da FilialEmpresa %i não possui itens relacionados.
Public Const ERRO_LEITURA_ITENSOP = 7466 'Sem parametros
'Erro de Leitura na Tabela de ItensOrdemProducao.
Public Const ERRO_TIPOMOVEST_NAO_CADASTRADO = 7467 'Parametro iCodTipoMovEstoque
'Tipo de Movimento de Estoque %i não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_ORDEMPRODUCAO = 7468 'Sem Parametros
'Erro de Leitura na Tabela OrdensDeProdução.
Public Const ERRO_CUSTO_MEDIO_NAO_CADASTRADO = 7469 'Parametros sCodProduto e iFilialEmpresa
'O Custo Médio para o Produto %s e Filial %i não foi cadastrado no Banco de Dados.
Public Const ERRO_PRODUTO_FILIAL_NAO_CADASTRADO = 7470 'Parametros sCodProduto e iFilialEmpresa
'O registro na Tabela ProdutoFilial para Produto %s e Filial %i não está cadastrado no Banco de Dados.
Public Const ERRO_MOVIMENTO_NAO_INTERNO = 7471 'Paramatro lCodMovEstoque
'Movimento %l não é do Tipo Interno.
Public Const ERRO_TIPOMOVINT_NAO_CADASTRADO = 7472 'Sem parametros
'Tipo de Movimento Interno não está cadastrado ou está Inativo.
Public Const ERRO_TIPOMOV_NAO_PREENCHIDO = 7473 'Parametro iLinhaGrid
'Tipo de Movimento do ítem %i do Grid Movimentos não foi preenchido.
Public Const ERRO_OP_NAO_PREENCHIDO = 7474 'Parametro iLinhaGrid
'Ordem de Produção do ítem %i do Grid Movimentos não foi preenchido.
Public Const ERRO_PRODUTOOP_NAO_PREENCHIDO = 7475 'Parametro iLinhaGrid
'Produto da Ordem de Produção do ítem %i do Grid Movimentos não foi preenchido.
Public Const ERRO_PRODUTO_FORA_OP = 7476 ' Parametros sProdutoOP, iItem, sOrdemProducao
'O Produto %s do Ítem %i do GridMovimentos não faz parte da ordem de Produção com código %s.
Public Const ERRO_PRODUTO_NAO_PARTICIPA_OP = 7477 'Parametros sCodProduto,sCodOP
'O Produto %s não faz parte da Ordem de Produção %s.
Public Const ERRO_MOV_EST_NAO_PRODUCAO = 7478 'Parametro lCodMovEstoque
'Movimento %l não é Entrada de Material Produzido.
Public Const ERRO_PRODUTO_NAO_PCP = 7479 'Parametro sCodProduto
'O Produto %s não pode participar da Produção.
Public Const ERRO_ESTORNO_MOVTO_ESTOQUE_NAO_CADASTRADO = 7480 'Parametro lCodigo
'O Movimento de Estoque com código %l não pode ser estornado pois não está cadastrado no Banco de Dados
Public Const ERRO_MOVTO_ESTOQUE_CADASTRADO = 7481 'Parametro lCodigo
'O Movimento de Estoque com código %l está cadastrado no Banco de Dados. Não é possível alterar.
Public Const ERRO_ESTORNO_ITEM_NAO_CADASTRADO = 7482 'Parametro iIndice
'O Ítem %i do Movimento não pode ser estornado pois não está cadastrado no Banco de Dados.
Public Const ERRO_ITEM_OP_NAO_CADASTRADO = 7483 'Parametros sCodigo , sProduto
'Ítem de Ordem de Producao não está cadastrado no Banco de Dados . Código da Ordem de Produção: %s, Código do Produto: %s
Public Const ERRO_INSERCAO_ITEMOPBAIXADA = 7484 'Sem Parametros
'Erro na inclusão de registro na tabela de itens de ordens de produção baixadas.
Public Const ERRO_LOCK_ITENSORDENSDEPRODUCAO = 7485 'Sem Parametros
'Erro na tentativa de fazer "lock" na tabela ItensOrdemProducao.
Public Const ERRO_INSERCAO_ORDENSDEPRODUCAO = 7486 'Sem Parametros
'Erro na inclusão de registro na tabela OrdensDeProducao.
Public Const ERRO_ATUALIZACAO_ITENSORDENSDEPRODUCAO = 7487 'Parametros iItem , sCodigo, iFilialEmpresa
'Erro na atualizacao de registro na tabela ItensOrdemProducao para o ítem %s da Produção %s com Empresa %i.
Public Const ERRO_ATUALIZACAO_ORDENSDEPRODUCAO = 7488 'Parametros sCodigo , iFilialEmpresa
'Erro na atualizacao de registro na tabela OrdensDeProducao para a Produção %s com Empresa %i.
Public Const ERRO_LOCK_ORDENSDEPRODUCAO = 7489 'Sem parametros
'Erro na tentativa de fazer "lock" na tabela OrdensDeProducao.
Public Const ERRO_LEITURA_ITENSORDENSPRODUCAO = 7490 'Sem parametros
'Erro na leitura da tabela de itens de ordens de produção.
Public Const ERRO_LEITURA_ORDENSDEPRODUCAO = 7491 'Sem parametros
'Erro na leitura da tabela de ordens de produção.
Public Const ERRO_EXCLUSAO_ITENSORDENSDEPRODUCAO = 7492 'Sem parametros
'Erro na exclusão de registro da tabela de itens de ordens de produção.
Public Const ERRO_EXCLUSAO_ORDENSDEPRODUCAO = 7493 'Sem parametros
'Erro na exclusão de registro da tabela de ordens de produção.
Public Const ERRO_LEITURA_ITENSORDEMPRODUCAO1 = 7494 'Parâmetros: sCodigo , iFilialEmpresa, sProduto
'Erro na leitura da tabela de ItensOrdemProducao com OP %s , Produto %s e Filial %i
Public Const ERRO_ALMOXARIFADO_INEXISTENTE = 7495 'Parametro: sNomeReduzido
'O Almoxarifado %s não está cadastrado.
Public Const ERRO_TIPOORIGEM_NAO_INFORMADO = 7496 'Parametro iLinhaGrid
'Tipo de Origem do ítem %i do Grid Movimentos não foi preenchido.
Public Const ERRO_TIPODESTINO_NAO_INFORMADO = 7497 'Parametro iLinhaGrid
'Tipo de Destino do ítem %i do Grid Movimentos não foi preenchido.
Public Const ERRO_ALMOXORIGEM_NAO_INFORMADO = 7498 'Parametro iLinhaGrid
'Almoxarifado de Origem do ítem %i do Grid Movimentos não foi preenchido.
Public Const ERRO_ALMOXDESTINO_NAO_INFORMADO = 7499 'Parametro iLinhaGrid
'Almoxarifado de Destino do ítem %i do Grid Movimentos não foi preenchido.
Public Const ERRO_MOVESTOQUE_NAO_INFORMADO = 7500 'Sem Parametros
'Não há Ítens de Movimento de Estoque no Grid
Public Const ERRO_MOVESTOQUE_NAO_TRANSFERENCIA = 7501 'Parametro lCodMovEstoque
'Movimento de Estoque %l não é uma Transferência.
Public Const ERRO_LEITURA_EMPENHO = 7502 'Parâmetros: iFilialEmpresa, lCodigo
'O Empenho da Filial %i e Código %l não está cadastrado no Banco de Dados.
Public Const ERRO_CODIGOOP_NAO_PREENCHIDO = 7503 'Sem parâmetros
'O preenchimento do código da Ordem de Produção é obrigatório.
Public Const ERRO_ITEM_NAO_PREENCHIDO = 7504 'Sem parâmetros
'O preenchimento do Item do Empenho é obrigatório.
Public Const ERRO_QUANT_EMPENHADA_NAO_PRRENCHIDA = 7505 'Sem parâmetros
'O preenchimento da quantidade à ser empenhada é obrigatório.
Public Const ERRO_EMPENHO_NAO_CADASTRADO = 7506 'Parâmetro: lCodigo
'O Empenho com Código %l não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_ITENSORDEMPRODUCAO = 7507 'Parâmetros: iFilialEmpresa, sCodigo, iItem
'Erro na leitura da tabela de ItensOrdemProducao com Filial %i, OP %s e Item %i.
Public Const ERRO_LEITURA_ITENSORDEMPRODUCAO2 = 7508 'Parâmetros: lNumIntDoc, iFilialEmpresa
'Erro na leitura da tabela de ItensOrdemProducao com número interno %l e Filial %i.
Public Const ERRO_ITEM_ORDEMPRODUCAO_NAO_CADASTRADO = 7509 'Parâmetros: iItem, sCodigo, iFilialEmpresa
'O Item %i da Ordem de Produção %s da Filial %i não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_ITENSORDEMPRODUCAO = 7510 'Parâmetros: iFilialEmpresa, sCodigoOP, iItemOP
'Erro na tentativa de fazer "lock" num registro da tabela ItensOrdemProducao com Filial %i, Item %i e OP %s.
Public Const ERRO_LOCK_EMPENHO = 7511 'Parâmetros: iFilialEmpresa, lCodigo
'Erro na tentativa de fazer "lock" na tabela de Empenho com Código %l e Filial %i.
Public Const ERRO_MODIFICACAO_EMPENHO = 7512 'Parâmetros: lCodigo, iFilialEmpresa
'Só é permitida a alteração de Quantidade e Emissão de um Empenho.
Public Const ERRO_EXCLUSAO_EMPENHO_REQUISICAO = 7513 'Parâmetros: lCodigo, iFilialEmpresa
'Não é permitido excluir o Empenho %l da Filial %i porque tem requisição.
Public Const ERRO_ATUALIZACAO_EMPENHO = 7514 'Parâmetros: lCodigo, iFilialEmpresa
'Erro na tentativa de atualizar um registro na tabela de Empenho com Empenho %l e Filial %i.
Public Const ERRO_LEITURA_EMPENHO1 = 7515 'Parâmetros: sProduto, iAlmoxarifado, lNumIntDoc
'Erro na leitura da tabela de Empenho com Produto %s, Almoxarifado %i e Número interno do item da OP %l.
Public Const ERRO_EMPENHO_REPETIDO = 7516 'Parâmetros: sProduto, iAlmoxarifado, lNumIntDoc
'Não pode haver mais de um empenho de um mesmo par (Produto %s e Almoxarifado %i) p/um item %l de uma OP.
Public Const ERRO_INSERCAO_EMPENHO = 7517 'Parâmetros: lCodigo, iFilialEmpresa
'Erro na tentatica de inserir um registro na tabela de Empenho com Empenho %l e Filial %i.
Public Const ERRO_EXCLUSAO_EMPENHO = 7518 'Parâmetros: lCodigo, iFilialEmpresa
'Erro na tentativa de excluir o Empenho %l da Filial %i da tabela de empenho.
Public Const ERRO_PRODUTO_IGUAL_ITEM_OP = 7519 'Sem parâmetros
'O Produto não pode ser igual ao Item da OP que se vai produzir.
Public Const ERRO_ITEM_EMPENHO_NAO_CADASTRADO = 7520 'Parâmetros: lNumIntDoc, iFilialEmpresa
'O Item da OP com número interno %l da Filial %i não está cadastrado no Banco de Dados.
Public Const ERRO_MODIFICACAO_ITEMOP_EMPENHO = 7521 'Parâmetros: lNumIntDoc, lNumIntDocItemOP
'Não é permitido modificar o Item da OP com número interno %l porque não pertence ao Empenho com número interno %l.
Public Const ERRO_ALMOXARIFADO_NAO_CADASTRADO = 7522 'Parâmetro: iCodigo
'O Almoxarifado com código %i não está cadastrado no Banco de Dados.
Public Const ERRO_ALMOXARIFADO_NAO_PREENCHIDO1 = 7524 'Sem parâmetros
'O Almoxarifado deve estar preenchido.
Public Const ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO = 7525 'Parâmetros: sProduto, iAlmoxarifado
'A associação entre o Produto %s e o Almoxarifado %i não está cadastrada.
Public Const ERRO_LOCK_PRODUTOS1 = 7526 'Patarmetro: sProduto
'Erro na tentativa de fazer lock no Produto de Código %s na Tabela de Produtos.
Public Const ERRO_LEITURA_CATEGORIACLIENTEITEM2 = 7527 'Parâmetro: sCategoria, sItem
'Erro na leitura do registro da categoria %s, do item %s, da tabela de itens das categoria de Cliente.
Public Const ERRO_INSERCAO_ICMSEXCECOES = 7528 'Parâmetros: sEstadoDestino, sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de inserção de registro da tabela de exceções de ICMS.Destino: %s, Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_ATUALIZACAO_ICMSEXCECOES = 7529 'Parâmetros: sEstadoDestino, sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de atualização de registro da tabela de exceções de ICMS.Destino: %s, Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_ICMSEXCECOES_INEXISTENTE = 7530 'Parâmetros: sEstadoDestino, sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'O registro não existe na tabela de exceções de ICMS.Destino: %s, Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_EXCLUSAO_ICMSEXCECOES = 7531 'Parâmetros: sEstadoDestino, sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de exclusão de registro da tabela de exceções de ICMS.Destino: %s, Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_CATEGORIAPRODUTO_TAMMAX = 7532 'parametros: tam max da categoria
'A categoria deve ter no máximo %i caracteres.
Public Const ERRO_VALOR_ALIQUOTA_INVALIDO = 7533
'O valor da Alioquota não pode sem de 100%.
Public Const ERRO_APROPRIACAO_NAO_DEFINIDA = 7534 'Sem parâmetros
'É obrigatória a definição da Apropriação.
Public Const ERRO_SUBSTITUTO_IGUAL_PRODUTO = 7535 'Parametro: sProduto
'O Código %s não pode pertencer ao Produto e ao Substituto ao mesmo tempo.
Public Const ERRO_SUBSTITUTO1_IGUAL_SUBSTITUTO2 = 7536 'Parâmetro: sProduto
'O Código dos Produtos substitutos não podem ser iguais.
Public Const ERRO_LEITURA_PRODUTOS2 = 7537 'Parametro: sNomeReduzido
'Erro na leitura do Produto com Nome Reduzido %s na Tabela de Produtos.
Public Const ERRO_LEITURA_PRODUTOCATEGORIA = 7538 'Parâmetro: sProduto
'Erro na leitura dos registros relacionados ao Produto %s na tabela Produtocategoria
Public Const ERRO_PRODUTO_PAI_NAO_CADASTRADO = 7539
'O Produto em questão não tem um Produto 'pai' dentro da hierarquia dos Produtos.
Public Const ERRO_PRODUTO_PAI_FINAL = 7540
'O Produto em questão possui um Produto "pai" Final. Produto Final não pode conter Produtos embaixo dele.
Public Const ERRO_MASCARA_PRODUTO_OBTERNIVEL = 7541 'Parâmetro: sProduto
'Erro na obtençao do nível do Produto. Produto = %s.
Public Const ERRO_MASCARA_RETORNAPRODUTONONIVEL = 7542 'Parâmetros: sProduto, iNivel
'Erro na obtençao do Produto %s no nível %i.
Public Const ERRO_DATA_PRECO_NAO_PREENCHIDA = 7543 'iTabela
'A Data do Preço para a Tabela %i não foi preenchida.
Public Const ERRO_PRODUTO_FINAL_COM_FILHOS = 7544 'ScODIGO
'O Produto %s não pode ser final pois possui Produtos embaixo dele.
Public Const ERRO_PRODUTO_SUBSTITUTO = 7545
'O Produto em questão não pode ser excluído pois é usado como Produto Substituto de outro Produto.
Public Const ERRO_LEITURA_MOVIMENTOESTOQUE = 7546
'Erro na leitura da Tabela MovimentoEstoque.
Public Const ERRO_PRODUTO_MOVIMENTOESTOQUE = 7547
'O Produto em questão não pode ser excluído pois faz parte de um Movimento de Estoque.
Public Const ERRO_PRODUTO_ITENSORDEMPRODUCAO = 7548
'O Produto em questão não pode ser excluído pois está sendo utilizado por um Item de Oredem de Produção.
Public Const ERRO_LEITURA_ITENSORDEMPRODUCAOBAIXADAS = 7549
'Erro na leitura da Tabela ItensOrdemProducaoBaixadas.
Public Const ERRO_PRODUTO_EMPENHO = 7550
'O Produto em questão não pode ser excluído pois está sendo utilizado em um Empenho.
Public Const ERRO_PRODUTO_KIT_EXCLUSAO = 7551
'O Produto em questão não pode ser excluído pois ele faz parte de um Kit.
Public Const ERRO_LEITURA_ITENSPEDIDODEVENDA1 = 7552
'Erro na leitura da tabela de Itens de Pedido de Venda.
Public Const ERRO_PRODUTO_ITEMPV = 7553
'O Produto em questão não pode ser excluído pois ele participa em um Item de Pedido de Venda.
Public Const ERRO_LEITURA_ITENSPEDIDODEVENDABAIXADOS1 = 7554
'Erro na leitura da tabela de Itens de Pedido de Venda Baixados.
Public Const ERRO_LEITURA_ITENSSOLICITACAODECOMPRA = 7555
'Erro na leitura da Tabela ItensSolicitacaoDeCompra.
Public Const ERRO_PRODUTO_ITENSSOLCOMPRA = 7556
'O Produto em questão não pode ser excluído pois ele participa em um Item de Solicitação de Compra.
Public Const ERRO_LEITURA_INVENTARIO = 7557
'Erro leitura da tabela Inventario.
Public Const ERRO_PRODUTO_INVENTARIO = 7558
'O Produto em questão não pode ser excluído pois ele participa em um Inventário.
Public Const ERRO_LEITURA_INVENTARIOPENDENTE = 7559
'Erro na leitura da Tabela InventarioPendente.
Public Const ERRO_LEITURA_RESERVA1 = 7560
'Erro na leitura da tabela Reserva.
Public Const ERRO_PRODUTO_RESERVA = 7561
'O Produto em questão não pode ser excluído pois ele participa de uma Reserva.
Public Const ERRO_LEITURA_PRODUTOKIT1 = 7562 'sProduto
'Erro na tentativa de leitura dos registros com Produto raiz = %s na Tabela ProdutoKit.
Public Const ERRO_LOCK_PRODUTOKIT = 7563
'Erro na tentativa de fazer 'lock' na tabela ProdutoKit.
Public Const ERRO_LEITURA_KIT = 7564 'sProduto
'Erro na tentativa de ler registro na tabela Kit com Produto Raiz = %s.
Public Const ERRO_LOCK_KIT = 7565
'Erro na tentativa de fazer 'lock' na tabela Kit.
Public Const ERRO_EXCLUSAO_KIT = 7566 'sProdutoRaiz
'Erro na tentativa de excluir registro da tabela Kit com o Produto Raiz = ?.
Public Const ERRO_EXCLUSAO_PRODUTOS = 7567 'sProduto
'Erro na tentativa de excluir o Produto Código %s da Tabela de Produtos.
Public Const ERRO_EXCLUSAO_SLDMESEST = 7568
'Erro na exclusão de registro na tabela SldMesEst.
Public Const ERRO_EXCLUSAO_SLDDIAEST = 7569
'Erro na exclusão de registro na tabela SldDiaEst.
Public Const ERRO_LOCK_PRODUTOSFILIAL = 7570
'Erro na tentativa de fazer 'lock' na tabela ProdutosFilial.
Public Const ERRO_EXCLUSAO_PRODUTOSFILIAL = 7571 'Produto
'Erro na exclusão de registro na tabela ProdutosFilial com Produto = %s.
Public Const ERRO_LEITURA_TABELASDEPRECOITENS = 7572
'Erro na leitura da tabela TabelasDePrecoItens.
Public Const ERRO_LEITURA_FORNECEDORPRODUTO = 7575
'Erro na leitura da tabela FornecedorProduto.
Public Const ERRO_LOCK_FORNECEDORPRODUTO = 7576
'Erro na tentativa de fazer 'lock' na tabela FornecedorProduto.
Public Const ERRO_EXCLUSAO_FORNECEDORPRODUTO = 7577
'Erro na exclusão de registro na tabela FornecedorProduto.
Public Const ERRO_LOCK_PRODUTOCATEGORIA = 7578
'Erro na tentativa de fazer 'lock' na tabela ProdutoCategoria.
Public Const ERRO_EXCLUSAO_PRODUTOCATEGORIA = 7579
'Erro na exclusão de registro na tabela ProdutoCategoria.
Public Const ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO = 7580
'O Código  do Produto não está preenchido.
Public Const ERRO_PRODUTO_MESMO_NOME_REDUZIDO = 7581 'sNomeReduzido
'Já existe um Produto cadastrado com o Nome Reduzido = %s
Public Const ERRO_PRODUTO_KIT_ALTERACAO = 7582
'O Produto em questão não pode ser alterado pois ele faz parte de um Kit.
Public Const ERRO_PRODUTO_UMESTOQUE_ALTERACAO = 7583
'O par (ClasseUM, UM de estoque) não pode ser alterado pois o produto está em EstoqueProduto.
Public Const ERRO_ATUALIZACAO_PRODUTOCATEGORIA = 7584 'Parâmetro: sProduto
'Erro na atualização de registro na Tabela ProdutoCategoria com o Produto %s.
Public Const ERRO_INSERCAO_PRODUTOS = 7587 'Parâmetro: sProduto
'Erro na tentativa de inserir o Produto %s na tabela de Produtos.
Public Const ERRO_ATUALIZACAO_PRODUTOS = 7588 'Parâmetro: sProduto
'Erro na tentativa de atualização do Produto %s na tabela de Produtos.
Public Const ERRO_DESCRICAO_PRODUTO_NAO_INFORMADA = 7589
'A descrição do Produto não foi informada.
Public Const ERRO_NOMEREDUZIDO_PRODUTO_NAO_INFORMADO = 7590
'O Nome Reduzido do Produto não foi informado.
Public Const ERRO_CLASSEUM_NAO_INFORMADA = 7591
'A Classe de Unidade de Medidas não foi informada.
Public Const ERRO_UM_COMPRA_NAO_INFORMADA = 7592
'A Unidade de Medida de Compra não foi informada.
Public Const ERRO_UM_VENDA_NAO_INFORMADA = 7593
'A Unidade de Medida de Venda não foi informada.
Public Const ERRO_UM_ESTOQUE_NAO_INFORMADA = 7594
'A Unidade de Medida de Estoque não foi informada.
Public Const ERRO_PRODUTO_SUBSTITUTO_INEXISTENTE = 7595 'sProduto
'O Produto com Código %s cadastrado como Produto Substituto não foi encontrado.
Public Const ERRO_QUANTIDADE_NULA = 7596
'para Quantidade nula Valor Total é nulo
Public Const ERRO_LEITURA_TIPOSDOCINFO = 7597
'Erro na leitura da Tabela TiposDocInfo.
Public Const ERRO_TIPO_NFISCAL_NAO_CADASTRADO = 7598 'Parâmtero: sTipo
'Tipo de Nota Fiscal %s não está cadastrado no Banco de Dados.
Public Const ERRO_NOTA_FISCAL_NAO_CADASTRADA1 = 7599 'lNumNotaFiscal
'A Nota Fiscal Número %l não está cadastrada no Banco de Dados
Public Const ERRO_TIPODOC_DIFERENTE_NF_ENTRADA = 7602 'iTipoNFiscal
'Tipo de Documento %i não é Nota Fiscal de Entrada.
Public Const ERRO_LEITURA_NFISCALBAIXADA1 = 7604 'Parâmetros: iTipoNFiscal, lFornecedor, iFilialForn, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscalBaixadas na Nota Fiscal com Tipo = %i, Fornecedor = %l, Filial = %i, Serie = %s e Número = %l.
Public Const ERRO_TIPO_NFISCAL_NAO_PREENCHIDO = 7605
'O Tipo de Nota Fiscal não está preenchido
Public Const ERRO_TIPO_NFISCAL_NAO_NORMAL = 7606
'O Tipo de Nota Fiscal não é Normal.
Public Const ERRO_RECEBIMENTO_MATERIAL_NAO_CADASTRADO = 7607 'Parâmetros: iTipoNFiscal, lFornecedor, iFilialForn, sSerie, lNumNotaFiscal
'O Recebimento de Material com os dados, Tipo = %i, Código de Fornecedor = %l, Código de Filial Fornecedor = %i, Serie = %s, Núm. Nota Fiscal = %l não está cadastrado no Banco de Dados.
Public Const ERRO_TAMANHO_SERIE = 7608
'Série tem limite de 3 dígitos.
Public Const ERRO_NATUREZAOP_NAO_PREENCHIDA = 7609
'Natureza da Operação não foi preenchida.
Public Const ERRO_DATAENTRADA_NAO_PREENCHIDA = 7610
'A Data de Entrada não foi preenchida.
Public Const ERRO_DATAEMISSAO_NAO_PREENCHIDA = 7611
'A Data de Emissão não foi preenchida.
Public Const ERRO_DATAENTRADA_ANTERIOR_DATAEMISSAO = 7612 'Parâmetros: dtDataEntrada, dtDataSaida
'A Data de Entrada %dt é anterior a Data de Emissão %dt.
Public Const ERRO_AUSENCIA_ITENS_NF = 7613
'Não há Ítens de Nota Fiscal no Grid.
Public Const ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA = 7614 'Parâmetro: iItem
'Quantidade do Ítem %i do Grid Ítens não foi preenchida.
Public Const ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO = 7615 'Parâmetro: iItem
'Almoxarifado do Ítem %i do Grid Ítens não foi preenchido.
Public Const ERRO_VALORTOTAL_NF_NEGATIVO = 7616
'Valor Total da Nota Fiscal é negativo
Public Const ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO = 7617
'Valor Unitário do Ítem %i do Grid Ítens não foi preenchido.
Public Const ERRO_ALTERACAO_NFISCAL = 7618 'Parâmetros: lFornecedor, iFilialForn, sSerie, lNumNotaFiscal, dtDataEmissao
'A Nota Fiscal com os dados Código do Fornecedor = %l, Código da Filial = %i, Série = %s, Número NF = %l, Data Emissao = %dt já está cadastrada e não é possível alterá-la.
Public Const ERRO_INSERCAO_NFISCAL = 7619 'Parâmetro: lFornecedor, iFilialForn, iTipoNFiscal, sSerie, lNumNotaFiscal
'Erro na inserção da Nota Fiscal com os dados Código do Fornecedor =%l, Código da Filial =%i, Tipo =%i, Serie =%s e Número NF =%l na tabela de Notas Fiscais.
Public Const ERRO_INSERCAO_ITENSNFISCAL = 7620 'Parâmetro: lNumNotaFiscal
'Erro na inserção dos Itens da NotaFiscal de Número = %l na tabela ItensNFiscal.
Public Const ERRO_LOCK_NFISCAL = 7621
'Erro na tentativa de fazer "lock" na tabela de NotasFiscais.
Public Const ERRO_INSERCAO_NFISCALBAIXADAS = 7622 'Parâmetro: lNumIntDoc
'Erro na tentativa de inserção da Nota Fiscal com Número Interno =%l na tabela de Notas Fiscais Baixadas.
Public Const ERRO_LOCK_ITENSNFISCAL = 7623
'Erro na tentativa de fazer "lock" na tabela ItensNFiscal
Public Const ERRO_INSERCAO_ITENSNFISCALBAIXADAS = 7624 'Parâmetro: lNumIntNF
'Erro na inserção de registros na tabela de Itens de Notas Fiscais para a Nota Fiscal com o Número Interno = %l.
Public Const ERRO_EXCLUSAO_ITENSNFISCAL = 7625 'Parâmetro: lNumIntNF
'Erro na exclusão dos Itens de Nota Fiscal com o Número Interno NF = %l.
Public Const ERRO_EXCLUSAO_NFISCAL = 7626 'lNumIntDoc
'Erro na exclusão da Nota Fiscal com o Número Interno =%l da tabela NFiscal.
Public Const ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA = 7627
'Série de Nota Fiscal Original não foi preenchida.
Public Const ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA1 = 7628
'A Nota Fiscal Original com Série = %s e Número =%l não está cadastrada no Banco de Dados.
Public Const ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO = 7629
'Número de Nota Fiscal Original não foi preenchido.
Public Const ERRO_VINCULO_NFENTRADA_NFPAGAR = 7630 'Parâmetros: lNumIntDoc, lNumIntDocCPR
'Nota Fiscal de Entrada com Código interno =%l está apontando para a Nota Fiscal a Pagar com Código interno =%l e esta não está cadastrada no Banco de Dados ou foi excluída.
Public Const ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA = 7631 'lNumIntDoc
'A Nota Fiscal Original com Número Interno =%l não está cadastrada no Banco de Dados.
Public Const ERRO_PESO_LIQUIDO_MAIOR_BRUTO = 7632 'dPesoLiq, dPesoBruto
'O Peso Líquido %d é maior que o Peso Bruto %d.
Public Const ERRO_PRODUTO_NAO_COMPRAVEL = 7633 'Parâmetro: sProduto
'O Produto %s não pode participar de compras.
Public Const ERRO_TIPONFISCALORIGINAL_NAO_ENCONTRADO = 7634
'O Tipo da Nota Fiscal Original não foi encontrado no Banco de Dados.
Public Const ERRO_VALOR_DESCONTO_100 = 7635
'Desconto não pode ser maior ou igual a 100%
Public Const ERRO_LOCK_ALMOXARIFADO1 = 7637 'Parametro: iCodAlmoxarifado
'Erro na tentativa de fazer "lock" na tabela Almoxarifado no Almoxarifado com o Código %i.
Public Const ERRO_UF_NAO_CADASTRADA = 7638 'sUF
'UF %s não está cadastrada no Banco de Dados.
Public Const ERRO_DESCRICAO_APROPRIACAO_NAO_PREENCHIDA = 7642
'Descrição de Apropriação deve ser preenchida.
Public Const ERRO_PRODUTO_ITENSNFISCAL = 7644 'Sem parametros
'O produto em questão não pode ser excluido pois está sendo utilizado em Notas Fiscais.
Public Const ERRO_PRODUTO_LANPENDENTE = 7645 'Sem parametros
'O produto em questão não pode ser excluido pois está relacionado com Lançamento Pendente
Public Const ERRO_CMP_NAO_APURADO = 7646 'sem parametros
'Custo Médio de Produção nunca foi apurado.
Public Const ERRO_LEITURA_ALMOXARIFADOS = 7648 'Parametro: iFilialEmpresa
'Ocorreu um erro na leitura dos Almoxarifados da Filial  %i.
Public Const ERRO_PAIS_NAO_CADASTRADO2 = 7649 'parametro: nome do pais
'País %s não está cadastrado.
Public Const ERRO_ALMOXARIFADO_RELACIONADO_ESTOQUEPRODUTO = 7650 'Parametro: iCodigo
'Almoxarifado com Código %i está relacionado com EstoqueProduto
Public Const ERRO_ALMOXARIFADO_RELACIONADO_PREVVENDA = 7651 'Parametro: iCodigo
'Almoxarifado com Código %i está relacionado com PrevVenda
Public Const ERRO_ALMOXARIFADO_RELACIONADO_EMPENHO = 7652 'Parametro: iCodigo
'Almoxarifado com Código %i está relacionado com Enpenho
Public Const ERRO_ALMOXARIFADO_RELACIONADO_INVENTARIO = 7653 'Parametro: iCodigo
'Almoxarifado com Código %i está relacionado com Inventario
Public Const ERRO_ALMOXARIFADO_RELACIONADO_INVENTARIOPENDENTE = 7654 'Parametro: iCodigo
'Almoxarifado com Código %i está relacionado com Inventario
Public Const ERRO_NAO_PODE_GRAVAR_FILIAL_DIFERENTE_DA_SUA = 7655 'Sem parametro
'Não é possível gravar um Almoxarifado de outra FilialEmpresa.
Public Const ERRO_NAO_PODE_EXCLUIR_FILIAL_DIFERENTE_DA_SUA = 7656 'Sem parametro
'Não é possível excluir um Almoxarifado de outra FilialEmpresa.
Public Const ERRO_LEITURA_FILIAIS = 7657 'Sem parametro
'Erro na leitura da tabela Filiais.
Public Const ERRO_INSERCAO_ALMOXARIFADO = 7658 'Sem parametro
'Erro na inserção do Almoxarifado.
Public Const ERRO_MODIFICACAO_ALMOXARIFADO = 7659 'Sem parametro
'Erro na modificação do Almoxarifado.
Public Const ERRO_EXCLUSAO_ALMOXARIFADO = 7660 'Parametro codigo do Almoxarifado
'Erro na tentativa de excluir Almoxarifado.
Public Const ERRO_CATEGORIAPRODUTO_SIGLA_NAO_INFORMADA = 7662 'Sem parâmetros
'A Sigla da Categoria deve ser informada.
Public Const ERRO_LEITURA_CLASSIFICACAOABC1 = 7663 'Parâmetro: iFilialEmpresa
'Erro na leitura da tabela ClassificacaoABC. FilialEmpresa %i.
Public Const ERRO_LEITURA_CLASSIFICACAOABC2 = 7664 'Parâmetro: lNumInt
'Erro na leitura da tabela ClassificacaoABC com número interno da Classificação ABC %l.
Public Const ERRO_CLASSIFICACAOABC_INEXISTENTE1 = 7665 'Parâmetros: sCodigo, iFilialEmpresa
'A ClassificacaoABC com Código %s da Filial %i não está cadastrada no Banco de Dados.
Public Const ERRO_CLASSIFICACAOABC_INEXISTENTE2 = 7666 'Parâmetro: lNumInt
'A ClassificacaoABC com número interno %l não está cadastrada no Banco de Dados.
Public Const ERRO_MES_INVALIDO = 7667 'Sem parâmetros
'O Mês deve estar entre 1 e 12.
Public Const ERRO_ANO_INVALIDO = 7668 'Sem parâmetros
'O Ano deve ser maior que 1900.
Public Const ERRO_FAIXA_INVALIDA = 7669 'Sem parâmetros
'As Faixas de Classificação devem estar entre 1 e 99.
Public Const ERRO_FAIXA_MAXIMA = 7670 'Sem parâmetros
'A soma dos valores das Faixas não pode ultrapassar o valor de 99.
Public Const ERRO_FALTA_TIPO_PRODUTO = 7671 'sem parametros
'Falta preencher Tipo de Produto.
Public Const ERRO_FALTA_MES_INICIAL = 7672 'sem parametros
'Falta preencher Mês Inicial
Public Const ERRO_FALTA_MES_FINAL = 7673 'sem parametros
'Falta preencher Mês Final
Public Const ERRO_FALTA_ANO_INICIAL = 7674 'sem parametros
'Falta preencher Ano Inicial
Public Const ERRO_FALTA_ANO_FINAL = 7675 'sem parametros
'Falta preencher Ano Final
Public Const ERRO_FALTA_FAIXA_A = 7676 'sem parametros
'Falta preencher Faixa A
Public Const ERRO_FALTA_FAIXA_B = 7677 'sem parametros
'Falta preencher Faixa B
Public Const ERRO_LEITURA_SLDMESEST2 = 7678  'Parametros:iFilialEmpresa
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst). FilialEmpresa=%i.
Public Const ERRO_SLDMESEST_INEXISTENTE2 = 7679 'Parametros  iFilialEmpresa, iAnoInicial, iAnoFinal
'Não existe registro de saldos mensais de estoque (SldMesEst) com os dados a seguir. FilialEmpresa=%i, Ano>=%i, Ano<=%i .
Public Const ERRO_LEITURA_ITENSCLASSIFABC = 7680 'Parametro: lNumIntClassABC
'Erro na leitura da tabela de Ítens de Classificação ABC (ItensClassifABC). ClassifABC = %l.
Public Const ERRO_FALTA_CODIGO_CLASSIFABC = 7681 'sem parametros
'Falta preencher código da Classificação ABC.
Public Const ERRO_FALTA_DATA_CLASSIFABC = 7682 'sem parametros
'Falta preencher data da Classificação ABC.
Public Const ERRO_LEITURA_CLASSIFICACAOABC = 7683 'Parametros: iFilialEmpresa, sCodClassifABC
'Erro na leitura da tabela ClassificacaoABC. Dados: FilialEmpresa=%i, Codigo=%s.
Public Const ERRO_CLASSIFICACAOABC_EXISTE_BD = 7684 'Parametros: iFilialEmpresa, sCodClassifABC
'Não é possível alterar a Classificação ABC com os dados a seguir pois está cadastrada no Banco de Dados. Dados: FilialEmpresa=%i, Codigo=%s.
Public Const ERRO_INSERCAO_ITEMCLASSIFABC = 7685 'Parametros: lNumIntClassifABC, sCodProduto
'Erro na inserção de registro na tabela ItensClassifABC. Dados do registro: ClassifABC=%l, Produto=%s.
Public Const ERRO_INSERCAO_CLASSIFABC = 7686 'Parametro: sCodClassifABC
'Erro na inserção de registro na tabela ClassificacaoABC. Dados: Codigo=%s.
Public Const ERRO_LEITURA_ITEMCLASSIFABC = 7687 'Parametro: lNumIntClassABC
'Erro na leitura da tabela ItensClassifABC. Dado: ClassifABC=%l.
Public Const ERRO_ITEMCLASSABC_INEXISTENTE = 7688 'Parametro: lNumIntClassABC
'Não existem Ítens Classificação ABC (ItensClassifABC) correspondentes à Classificação de número interno %l.
Public Const ERRO_SLDMESEST_INEXISTENTE3 = 7689 'Parametros: iTipoProduto, iFilialEmpresa, iAnoInicial, iAnoFinal
'Não existe registro de saldos mensais de estoque (SldMesEst) com os dados a seguir. TipoProduto=%i, FilialEmpresa=%i, Ano>=%i, Ano<=%i .
Public Const ERRO_MODIFICACAO_ITENSCLASSIFABC = 7690 'Parametros: lNumIntClassifABC, sCodProduto
'Erro na atualização de registro na tabela ItensClassifABC. Dados do registro: ClassifABC=%l, Produto=%s.
Public Const ERRO_PRODUTOSFILIAL_INEXISTENTE1 = 7691 'Parametro: iFilialEmpresa
'Não existem registros na tabela ProdutosFilial correspondentes a FilialEmpresa %i.
Public Const ERRO_MODIFICACAO_PRODUTOSFILIAL = 7692 'Parametros: iFilialEmpresa, sCodProduto
'Erro de atualização na tabela ProdutosFilial no registro com chave: FilialEmpresa = %i, Produto = %s.
Public Const ERRO_CLASSIFICACAOABC_INEXISTENTE = 7693 'Parametro: lNumInt
'ClassificacaoABC com número interno %l não está cadastrada no Banco de Dados.
Public Const ERRO_MODIFICACAO_CLASSIFABC = 7694 'parametro: sCodigo
'Erro na atualização de Classificação ABC (tabela ClassificacaoABC) com código %s.
Public Const ERRO_AUSENCIA_CLASSEB = 7695 'parametro: lClassifABC
'Ausência de Produtos com classe B na Classificação ABC com nº interno %l.
Public Const ERRO_AUSENCIA_CLASSEC = 7696 'parametro: lClassifABC
'Ausência de Produtos com classe C na Classificação ABC com nº interno %l.
Public Const ERRO_ANOINIC_MAIOR_ANOFINAL = 7697 'sem parametro
'Ano inicial é maior do que ano final.
Public Const ERRO_MESINIC_MAIOR_MESFINAL = 7698 'sem parametro
'Mês inicial é maior do que mês final para o mesmo ano.
Public Const ERRO_LOCK_CLASSABC = 7699 'parametros: sCodigo, iFilialEmpresa
'Erro na tentativa de lock na tabela ClassificacaoABC. Dados do registro: Codigo=%s, FilialEmpresa=%i.
Public Const ERRO_CLASSIFICACAOABC_MAIS_RECENTE = 7700 'parametros: iFilialEmpresa, sCodigo
'A classificação ABC com chave FilialEmpresa=%i e Codigo=%s é a mais recente que atualiza Produtos (do tipo especificado).
Public Const ERRO_EXCLUSAO_CLASSIFICACAOABC = 7701 'parametros: sCodigo, iFilialEmpresa
'Erro na tentativa de excluir na tabela ClassificacaoABC o registro com código %s e FilialEmpresa %i.
Public Const ERRO_EXCLUSAO_ITEMCLASSIFABC = 7702 'parametros: lNumInt
'Erro na tentativa de exclusão na tabela ItensClassifABC. Ítem da classificação ABC com número interno %l.
Public Const ERRO_DEMANDA_TOTAL_NULA = 7703 'Parâmetros: sCodigo, iFilialEmpresa
'Produtos da Classificação ABC com código %s da Filial %i têm demanda total nula.
Public Const ERRO_CUSTO_PRODUCAO_APURADO = 7704 'Sem parametro
'O Custo Médio de Produção foi calculado em todos os meses fechados.
Public Const ERRO_AUSENCIA_PRODUTO_SLDMESEST = 7705 'Parametros: sProduto
'Falta do Produto %s na tabela SldMesEst.
Public Const ERRO_CP_MES_ANTERIOR_NAO_APURADO = 7706 'Parametros:sMes,iAno
'Custo de Producao do mes de %s de %i não foi apurado
Public Const ERRO_LEITURA_SLDMESEST1 = 7707  'Parametros:iAno, iFilialEmpresa
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst). Ano=%i, FilialEmpresa=%i.
Public Const ERRO_ATUALIZACAO_CUSTO_PRODUCAO = 7708 'Parametros : iAno , iFilialEmpresa
'Ocorreu um erro na atualização da tabela SldMesEst, campo Custo Producao. Ano:%i, FilialEmpresa:%i.
Public Const ERRO_ANO_NAO_PREENCHIDO = 7709
'O ano deve ser preenchido.
Public Const ERRO_MES_NAO_PREENCHIDO = 7710
'O mês deve ser preenchido.
Public Const ERRO_CUSTOPRODUCAO_NAO_PREENCHIDO = 7711
'O valor do Custo Real de Produção deve ser preenchido.
Public Const ERRO_CUSTOSTANDARD_NAO_PREENCHIDO = 7712
'O valor do Custo Standard deve ser preenchido.
Public Const ERRO_VALOR_NAO_NEGATIVO = 7713  'Parametro String com um valor monetário
'O valor digitado não pode ser negativo. Valor = %s.
Public Const ERRO_ESTOQUEMES_ANOS_INEXISTENTES = 7714 'Parametro iFilialEmpresa
'Não existe nenhum ano disponível na tabela de EstoqueMes para a FilialEmpresa = %i
Public Const ERRO_ESTOQUEMES_MESES_INEXISTENTES = 7715 'Parametros iFilialEmpresa , iAno
'Não existe nenhum mês disponível na tabela de EstoqueMes para a FilialEmpresa = %i no Ano = %i
Public Const ERRO_CUSTOS_INEXISTENTES = 7716 'Parametros iFilialEmpresa , iAno , sProduto
'Não foi encontrado o registro com FilialEmpresa = %i , Ano = %i e Produto = %s na Tabela SldMesEst.
Public Const ERRO_ESTOQUEMES_CUSTO_APURADO = 7717 'Parametro : iFilialEmpresa ,iAno,iMes
'Não é possível alterar o custo, pois o custo médio de produção já foi apurado para o mes em questão. FilialEmpresa = %i , Ano = %i , Mês = %i.
Public Const ERRO_CUSTO_PROD_MES_ABERTO = 7718 'Parâmetros: sProduto, iMes, iAno
'Não é possível alterar Custo de Produção do produto %s no mês %i do ano %i.
Public Const ERRO_CONTROLE_ESTOQUE_NAO_PREENCHIDO = 7719 'Sem parâmetros
'O preenchimento do Controle de Reserva/Estoque é obrigatório.
Public Const ERRO_CONTROLE_ESTOQUE_NAO_CADASTRADO = 7720 'Parâmetro: ControleEstoque.Text
'O Controle de Estoque %s não está cadastrado no Banco de Dados.
Public Const ERRO_CLASSE_PRODUTO_INEXISTENTE = 7721 'Parâmetro: ClasseABC.Text
'O campo Classe ABC só pode ser preenchido com "A", "B" ou "C".
Public Const ERRO_PRODUTO_NAO_CADASTRADO_ALMOXARIFADO_ESTOQUEPRODUTO = 7722 'Parametros: sProduto e iAlmoxarifado
'O produto %s não se encontra cadastrado no almoxarifado %i na tabela EstoqueProduto.
Public Const ERRO_EXCLUSAO_MOVIMENTO_ESTOQUEPRODUTO = 7723 'Parâmetros: iFilialEmpresa, sProduto, iAlmoxarifado
'A associação entre o Produto %s e o Almoxarifado %s da Filial %s não pode ser excluída pois existe movimentação de estoque.
Public Const ERRO_INSERCAO_PRODUTOSFILIAL = 7724 'Parâmetros: iFilialEmpresa, sProduto
'Erro na tentativa de inserir registro na tabela ProdutosFilial da Filial %i e Produto %s.
Public Const ERRO_INSERCAO_SLDMESEST = 7725 'Parâmetros: iAno, iFilialEmpresa, sProduto
'Erro na tentativa de inserir um registro na tabela SldMesEst com Ano %i, Filial %i e Produto %s.
Public Const ERRO_LEITURA_ESTOQUEPRODUTO2 = 7727 'Parâmetros: sProduto
'Erro na leitura da tabela de EstoqueProduto com Produto %s.
Public Const ERRO_INSERCAO_ESTOQUEPRODUTO = 7728 'Parametros: sProduto e iAlmoxarifado
'Erro na tentativa de inserir registro na tabela EstoqueProduto com Produto %s e Almoxarifado %i.
Public Const ERRO_PRODUTO_FILIAL_INEXISTENTE = 7729 'Parâmetros: giFilialEmpresa, sProduto
'O Produto %s da Filial %i não existe na tabela de ProdutosFilial.
Public Const ERRO_NAO_EXISTE_RESERVAS = 7730 ' Parametros lCodPedido, sCodProduto
'Não existe Reservas para o Pedido %l do Produto %s.
Public Const ERRO_QUANTIDADE_FATURADA_MAIORZERO = 7731 'Parametro dQuantFaturada
'Produto com quantidade faturada %d não pode ser substítuido
Public Const ERRO_TRATAMENTO_NAO_INFORMADO = 7732 'Sem parametros
'Uma opção de tratamento deve ser escolhida.
Public Const ERRO_NAOEXISTE_MES_ABERTO = 7733 'Sem parametro
'Não foi encontrado mês e ano abertos.
Public Const ERRO_ABERTURA_NOVOANO_SDLMESEST = 7734 'Parametros: iAno, iFilialEmpresa, sProduto
'Não foi possivel abrir %i para o produto %s da Filial %i
Public Const ERRO_ABERTURA_NOVOMES_SLDMESEST = 7735 'Parametros: iAno, iFilialEmpresa, iMes
'Não foi possivel abrir um novo mês com os dados a seguir. Ano: i%, FilialEmpresa: %i, Mes: %i
Public Const ERRO_ALTERACAO_STATUS12 = 7736 'Sem Parametros
'Não foi possivel a alteração do Status do mês 12
Public Const ERRO_FECHAMENTO_MES_ANTERIOR = 7737 'Sem Parametros
'Não foi possivel fechar o mes na tabela EstoqueMes
Public Const ERRO_INSERCAO_NOVO_MES = 7738 'Sem Parametros
'Não foi possivel inserir um novo mes na tabela EstoqueMes
Public Const ERRO_LEITURA_FORNECEDORPRODUTO1 = 7739 'Parâmetros: sCodigo
'Erro na leitura da tabela de FornecedorProduto com Produto %s.
Public Const ERRO_FORNECEDORPRODUTO_NAO_ENCONTRADO = 7740 'Parâmetros: lFornecedor, sProduto
'O Fornecedor %l do Produto %s não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_FORNECEDORPRODUTO_FORNECEDOR = 7741 'Parâmetro: sCodigo
'Erro na leitura de Fornecedores para associados ao Produto %s.
Public Const ERRO_LEITURA_FORNECEDORES1 = 7742 'Parâmetro: lFornecedor
'Erro na leitura da tabela de Fornecedores com Fornecedor %l.
Public Const ERRO_ATUALIZACAO_FORNECEDORPRODUTO = 7743 'Parâmetros: lFornecedor, sProduto
'Erro na tentativa de atualizar registro na tabela FornecedorProduto com Fornecedor %l e Produto %s.
Public Const ERRO_INSERCAO_FORNECEDORPRODUTO = 7744 'Parâmetros: lFornecedor, sProduto
'Erro na tentativa de inserir registro na tabela FornecedorProduto com Fornecedor %l e Produto %s.
Public Const ERRO_ATUALIZACAO_PRODUTOSFILIAL = 7745 'Parâmetros: giFilialEmpresa, sProduto
'Erro na tentativa de atualizar registro na tabela ProdutosFilial com Filial %i e Produto %s.
Public Const ERRO_PRODUTO_FILIAL_PADRAO = 7746 'Parâmetros: giFilialEmpresa, sProduto
'O Produto %s da Filial %i já possui um fornecedor padrão.
Public Const ERRO_PRODUTO_JA_SELECIONADO = 7747 'Parametro sProduto
'O produto já foi selecionado com outro item.
Public Const ERRO_ORDENACAO_NAO_ENCONTRADA = 7748 'Parâmetro sOrdenacao
'O tipo de ordenaçao %s não foi encontrada.
Public Const ERRO_CLIENTEDE_MAIOR_CLIENTEATE = 7749 'Sem parâmetros
'Cliente De não pode ser maior do que o Cliente Até.
Public Const ERRO_PRODUTODE_MAIOR_PRODUTOATE = 7750
'Produto De não pode ser maior do que o Produto Até.
Public Const ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL = 7751 'Sem parâmetros
'Pedido Inicial não pode ser maior do que o Pedido Final.
Public Const ERRO_DATAPREVISAOINICIO_NAO_INFORMADO = 7752 'Parametro : iIndice
'A Data de Previsão de Início não foi informada.
Public Const ERRO_DATAPREVISAOFIM_NAO_INFORMADO = 7753 'Parametro : iIndice
'A Data de Previsão de Fim não foi informada.
Public Const ERRO_QUANTOP_MAIOR_QUANTFALTA = 7754
'A quantidade ordenada não pode ultrapassar a quantidade do pedido que falta ser atendida.
Public Const ERRO_LEITURA_ITENSPV_GERACAO_OP = 7755
'Erro na leitura na tabela de Itens de Pedido de Venda.
Public Const ERRO_SEM_ITENSPV_GERACAO_OP = 7756
'Não existe nenhum pedido de venda para o qual possa ser gerado ordem de produção.
Public Const ERRO_ETIQUETA_NAO_PREENCHIDO = 7757 'Parametro iLinhaGrid
'Etiqueta do ítem %i do Grid Inventarios não foi preenchido.
Public Const ERRO_CUSTOUNITARIO_NAO_PREENCHIDO = 7758 'Parametro iLinhaGrid
'Custo Unitário do ítem %i do Grid de Inventário não foi preenchido.
Public Const ERRO_CUSTOUNITARIO_PREENCHIDO = 7759 'Parametro iLinhaGrid
'Relação Tipo de Quantidade e Custo Unitário inválida no ítem %i do Grid de Inventário.
Public Const ERRO_LEITURA_INVENTARIO_PENDENTE = 7760 'Parametro: codigo do inventario
'Erro de Leitura na Tabela Inventarios Pendentes para o código %s.
Public Const ERRO_INSERCAO_INVENTARIO = 7761 'Parametro sCodigo
'Erro na inserção do Inventario de código = %s na tabela Inventario.
Public Const ERRO_LOCK_INVENTARIO = 7762 'Sem Paramentro
'Erro na tentativa de fazer "lock" na tabela de Inventarios.
Public Const ERRO_ATUALIZACAO_INVENTARIO = 7763 'Parametro dtData
'Erro na atualização da Data de Inventario %dt na Tabela EstoqueProduto.
Public Const ERRO_LOTE_INCOMPATIVEL = 7764 'Parametros iLote , iLote
'Lote incompatível.
Public Const ERRO_INV_PENDENTE_EM_ATUALIZACAO = 7765 'Parametro iLote
'O Lote %i está em atualização na tabela InvLotePendente.
Public Const ERRO_INV_PENDENTE_NAO_CADASTRADO = 7766 'Parametro sCodigo
'O Inventário Pendente com código %s não está cadastrado no Banco de Dados.
Public Const ERRO_INV_LOTE_PEND_CADASTRADO = 7767 'Parametro iLote
'O Lote com código %i já foi atualizado/lançado e portanto não é mais um lote pendente.
Public Const ERRO_INVENTARIO_LANCADO = 7768 'Parametro sCodigo
'O Inventário com código %s está cadastrado como inventário lançado no Banco de Dados.
Public Const ERRO_INVENTARIO_NAO_INFORMADO = 7770 'Sem Parametros
'Não há Ítens de Inventario no Grid.
Public Const ERRO_INVENTARIO_CADASTRADO = 7771 'Parametro sCodigo
'O Inventário com código %s está cadastrado no Banco de Dados. Não é possível alterar.
Public Const ERRO_INVENTARIO_PENDENTE_CADASTRADO = 7772 'Parametro sCodigo
'O Inventário com código %s está cadastrado como Inventário Pendente no Banco de Dados.
Public Const ERRO_LOCK_INVENTARIOPENDENTE = 7773 'Parametros sCodigo , iFilialEmpresa
'Erro na tentativa de fazer "lock" na tabela de Inventarios para o Código %s com Filial %i.
Public Const ERRO_EXCLUSAO_INVENTARIOPENDENTE = 7774 'Parametros sCodigo , iFilialEmpresa
'Erro na tentativa de excluir o Inventario Pendente com código %s e Filial %i.
Public Const ERRO_INSERCAO_INVENTARIOPENDENTE = 7775 'Parametros sCodigo , iFilialEmpresa
'Erro na tentativa de inserção do Inventario Pendente de código %s e Filial %i.
Public Const ERRO_KIT_LIMPAR_ANTES = 7776 'sem parametros
'Antes de começar a definir outro kit deve-se apertar o botão para limpar a tela
Public Const ERRO_PRODUTO_NAO_CONFERE_SEL = 7777 'sem parametros
'O Produto não confere com o do elemento que está sendo alterado.
Public Const ERRO_QUANTIDADE_NAO_INFORMADA = 7778 'sem parametro
'A Quantidade do componente não foi informada.
Public Const ERRO_COMPOSICAO_NAO_INFORMADA = 7779 'sem parametros
'Tipo de composição do componente não informada.
Public Const ERRO_SIGLAUM_NAO_INFORMADA = 7780 'sem parametros
'A sigla da unidade de medida do componente não foi informada.
Public Const ERRO_PRODUTO_RAIZ = 7781 'parametro sProduto
'O Produto %s é o produto raiz, portanto não pode ser usado como componente.
Public Const ERRO_NAO_KIT_INTERMEDIARIO = 7782 'Parametro sCodigo
'O produto %s não pode ser um produto intermediário de um Kit.
Public Const ERRO_NIVEL_MAXIMO_KIT = 7783 'parametro sCodigo , iNivel
'O produto está no último nível permitido. Produto=%s, Nível = %i.
Public Const ERRO_LEITURA_INVLOTE = 7784 'Parametros iFilialEmpresa, ilote
'Erro na leitura do Inventário Lote. Filial = %i, Lote = %i
Public Const ERRO_INVLOTE_ATUALIZADO = 7785 'Parametros iFilial, iLote
'Este lote já foi contabilizado, portanto não pode ser editado. Filial = %i, Lote = %i.
Public Const ERRO_LOCK_INVLOTEPENDENTE = 7786 'Parametros iLote, iFilialEmpresa
'Erro na tentativa de fazer "lock" na tabela InvLotePendente para o Lote %i com Filial %i.
Public Const ERRO_ASSOCIACAO_INVENTARIO = 7787 ' Parametros iFilial , iLote
'Impossível excluir . O Lote %i da Filial %i possui associações na Tabela InventarioPendente.
Public Const ERRO_EXCLUSAO_INVLOTEPENDENTE = 7788 ' Parametros iFilialEmpresa , iLote
'Erro na tentativa de excluir o Lote %i da Filial %i na tabela InvLotePendente.
Public Const ERRO_ATUALIZACAO_INVLOTEPENDENTE = 7789 ' Parametros iFilialEmpresa , iLote
'Erro na tentativa de atualizar o Lote %i da Filial %i na tabela InvLotePendente.
Public Const ERRO_INVLOTE_CADASTRADO = 7790 'Parametro iLote
'O Lote número %i já está cadastrado na tabela InvLote do Banco de Dados .
Public Const ERRO_INSERCAO_INVLOTEPENDENTE = 7791 'Parametros iLote , iFilialEmpresa
'Erro na tentativa de inserção do Lote de Inventário %i da Filial %i.
Public Const ERRO_INV_LOTE_PEND_NAO_CADASTRADO = 7792 'Parametro iLote
'Lote Pendente de Inventário com código %i não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_INVLOTEPENDENTE = 7793 'Sem parametros
'Erro na leitura da tabela de lotes de inventário pendentes.
Public Const ERRO_CODIGO_INCOMPATIVEL_MOVESTOQUE = 7794 'Parametro lCodigo
'O Movimento com o código %l já existe mas não é Tipo de Movimento de Estoque Interno.
Public Const ERRO_TIPODOC_DIFERENTE_NF_ENTRADA_DEVOLUCAO = 7795 'iTipoNFiscal
'Tipo de Documento %i não é Nota Fiscal de Entrada de Devolução.
Public Const ERRO_SERIE_NUMERO_ORIGINAL_FALTANDO2 = 7799
'Para trazer dados de Nota Fiscal Original é necessário preencher Série e Número.
Public Const ERRO_RECEBIMENTO_MATERIAL_NAO_CADASTRADO2 = 7800 'Parâmetros: iTipoNFiscal, lFornecedor, iFilialForn, lCliente, iFilialCli, sSerie, lNumNotaFiscal
'O Recebimento de Material com dados, Tipo=%l, Código de Fornecedor=%l, Código de Filial Fornecedor=%l, Código de Cliente=%l, Código de Filial Cliente=%l, Série=%s, Núm. Nota Fiscal =%l não está cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_ITENSNFISCAL2 = 7801
'Erro na leitura da tabela ItensNFiscal.
Public Const ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO = 7802 'Parâmetros: iItem, sSerie, lNumNotaFiscal
'O Ítem %i da Nota Fiscal Original série %s, número %l não está cadastrado no Banco de Dados.
Public Const ERRO_PRODUTO_NAO_CASADO = 7803 'Parâmetros: iLinha, sProdutoItemNF, sProdutoItemNFOrig
'Na linha %i do Grid Ítens o produto não corresponde ao produto da NotaFiscalOriginal. Código do Produto: %l, código do Produto da NF Original: %s.
Public Const ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO2 = 7804 'Parâmetro: lNumIntDoc
'Ítem de Nota Fiscal Original com Número Interno %l não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_ITEM_NFORIGINAL = 7805 'lNumIntDoc
'Erro na tentativa de fazer "lock" no Ítem de Nota Fiscal Original com o Número Interno %l.
Public Const ERRO_QUANT_DEVOLVIDA_A_MAIOR = 7806 'Parâmetros: iItem, dQuantDevolvida, sUnidadeMed, dQuantidade, sUnidadeMed
'A Quantidade do ítem %i juntamente com a quantidade já devolvida totaliza %d %s que ultrapassa a quantidade original de %d %s
Public Const ERRO_ALTERACAO_NFISCAL_DEV = 7807 'Parâmetros: lFornecedor, iFilialForn, lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal de Devolucao com os dados Código do Fornecedor =%l, Código da Filial Fornecedor =%i, Código do Cliente =%l AND Código da Filial Cliente =%i, Tipo =%i, Série NF =%s, Número NF =%l, Data Emissão =%dt está cadastrada no Banco de Dados. Não é possível alterar."
Public Const ERRO_INSERCAO_NFISCAL1 = 7808 'Parâmetro: lFornecedor, iFilialForn, lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal
'Erro na inserção da Nota Fiscal com os dados Código do Fornecedor =%l, Código da Filial =%i, Código do Cliente =%l, Código da Filial CLiente =%i, Tipo =%i, Serie =%s e Número NF =%l na tabela de Notas Fiscais.
Public Const ERRO_ALTERACAO_NFISCAL_EXTERNA2 = 7809 'Parâmetros: lFornecedor, iFilialForn, lCliente, iFilialCli, sSerie,lNumNotaFiscal,dtDataEmissao
'Nota Fiscal Externa com os dados Código do Fornecedor = %l, Código da Filial Fornecedor = %i, Cliente = %l, Código da Filial Cliente = %i, Série = %s, Número = %l, Data Emissão = %dt. Não é possível alterar.
Public Const ERRO_TIPODOC_DIFERENTE_NF_ENTRADA_REMESSA = 7810 'iTipoNFiscal
'Tipo de Documento %i não é Nota Fiscal de Entrada Remessa.
Public Const ERRO_ALTERACAO_NFISCAL_REM = 7811 'Parâmetros: lFornecedor, iFilialForn, lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal de Remessa com os dados Código do Fornecedor =%l, Código da Filial Fornecedor =%i, Código do Cliente =%l AND Código da Filial Cliente =%i, Tipo =%i, Série NF =%s, Número NF =%l, Data Emissão =%dt está cadastrada no Banco de Dados. Não é possível alterar."
Public Const ERRO_TIPODOC_DIFERENTE_NF_FATURA_ENTRADA = 7812 'Parâmetro: iTipoDocInfo
'Tipo de Documento %i não é Nota Fiscal Fatura de Entrada.
Public Const ERRO_ALTERACAO_NFISCAL_EXTERNA = 7813 'Parâmetros:lFornecedor, iFilialForn, sSerie, lNumero, dtDataEmissao
'Nota Fiscal Externa com os dados Código do Fornecedor = %l, Código da Filial = %i, Série = %s, Número = %l, Data Emissão = %dt está cadastrada no Banco de Dados. Não é possível alterar.
Public Const ERRO_ESTOQUEPRODUTO_QTDE_EMPENHADA_NEG = 7814 'parametros: produto mascarado e numero do almoxarifado
'A quantidade empenhada do produto %s no almoxarifado %d se tornaria negativa
Public Const ERRO_LEITURA_EMPENHO_ITEMOP = 7815 'sem parametros
'Erro na leitura de empenho associado a item de ordem de produção
Public Const ERRO_EMPENHO_COM_REQUISICAO = 7816 'parametro = codigo do empenho
'O empenho %l nao pode ser excluido pois alguma quantidade foi requisitada utilizando-o
Public Const ERRO_FILIALALMOXARIFADO_DIFERENTE_FILIALCORRENTE = 7817 'Parametro: iCodigoAlmoxarifado, iCodigoFilial
'O Almoxarifado com código %i não pertence a filial corrente(%i) da empresa.
Public Const ERRO_ALMOXARIFADO_NAO_CADASTRADO1 = 7818 'Parametro: sNomeReduzido
'O Almoxarifado com nome %s não está cadastrado no Banco de Dados.
Public Const ERRO_INCLUSAO_ITENSORDENSDEPRODUCAO = 7819 'Sem parametros
'Erro na inclusão de registro na tabela de itens de ordens de produção.
Public Const ERRO_LEITURA_ORDENSDEPRODUCAOBAIXADAS = 7820 'Sem parametros
'Erro na leitura da tabela de ordens de produção baixadas.
Public Const ERRO_LEITURA_ITENSORDENSPRODUCAOBAIXADAS = 7821 'Sem parametros
'Erro na leitura da tabela de itens de ordens de produção baixadas.
Public Const ERRO_ORDEMDEPRODUCAO_BAIXADA = 7822 'Parametro: sCodigo
'A ordem de produção de codigo %s já foi baixada.
Public Const ERRO_ITEMOP_SITUACAO_NAO_EXCLUIVEL = 7823 'iItem, sCodigoOP
'Não é possível realizar exclusão do item %i da ordem de produção %s, pois sua situação não é "normal" ou "desabilitada".
Public Const ERRO_ITEMOP_NAO_EXCLUIVEL = 7824 'iItem, sCodigoOP
'Não é possível realizar exclusão do item %i da ordem de produção %s.
Public Const ERRO_ITEMOP_ASSOCIADO_MOVESTOQUE = 7825 'iItem, sCodigoOP
'Não é possível realizar exclusão do item %i da ordem de produção %s, pois este item está associado a um movimento de estoque.
Public Const ERRO_ITEMPEDIDO_INEXISTENTE = 7826 'Parametros lPedidoDeVenda,sProduto
'O item do Pedido de Venda=%l com Produto=%s não existe .
Public Const ERRO_ORDEMDEPRODUCAO_INEXISTENTE = 7827 'Parametro: sCodigoOP
'A ordem de producao de código %s não existe.
Public Const ERRO_LOCK_PRODUTO = 7828 'Sem parametro
'Erro na tentativa de fazer "lock" na tabela de Produtos.
Public Const ERRO_ITEMOP_QTDE_MENOR_PROD = 7829 'Sem parametro
'A quantidade não pode ser alterada para um valor menor que o já produzido
Public Const ERRO_PRODUTO_DUPLICADO = 7830 'Parametro sProduto , sCodigoOP
'O Produto %s adicionado a Ordem de Produção %s já existe no Grid de Ítens.
Public Const ERRO_BAIXAR_ITEMNOVO = 7831 'Sem parametro
'Impossível baixar um ítem novo do Grid de Ítens !
Public Const ERRO_ALTERACAO_SITUACAO = 7832 ' sSituacao , sSituacao
'Impossível fazer a alteração na Situação do ítem de %s para %s.
Public Const ERRO_ITEMPV_NAO_PREENCHIDO = 7833 'Parametro iLinhaGrid
'ItemPV da linha %i do Grid de Ítens não foi preenchido.
Public Const ERRO_FILIALPEDIDO_NAO_PREENCHIDA = 7834 'Parametro iLinhaGrid
'Filial do Pedido da linha %i do Grid de Ítens não foi preenchida.
Public Const ERRO_PEDIDOVENDAID_NAO_PREENCHIDO = 7835 'Parametro iLinhaGrid
'Identificador do Pedido de Venda da linha %i do Grid de Itens não foi preenchido.
Public Const ERRO_QUANTIDADE_ESTOQUEMAXIMO = 7836 'parametro sProduto
'A soma da quantidade ordenada mais a quantidade disponível é maior que a quantidade de estoque máximo do produto %s.
Public Const ERRO_LOCK_ITENSPEDIDODEVENDA = 7837 'Sem parametro
'Erro na tentativa de fazer "lock" na tabela ItensPedidoDeVenda.
Public Const ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA = 7838
'Erro na atualização da tabela ItensPedidoDeVenda.
Public Const ERRO_DATA_FIM_MENOR = 7839
'A Data de Previsão de Fim de Produção é menor que a Data de Previsão de Início.
Public Const ERRO_DESTINACAO_DEPENDENTE = 7840
'A Destinação não é Pedido de Venda .
Public Const ERRO_DETERMINACAO_QUANTMAX = 7841 'Parametro : iFilialEmpresa
'Não foi possível determinar a quantidade máxima de estoque da Filial = ?
Public Const ERRO_CODIGO_INCOMPATIVEL_PENTRADA = 7842 'Parametro lCodigo
'O Movimento com o código %l já existe mas não é Tipo de Produção Entrada.
Public Const ERRO_PRODUTO_NAO_PRODUZIVEL = 7843 'Parametro sCodProduto
'O Produto %s não pode ser produzido.
Public Const ERRO_MOV_EST_REQ_NAO_PRODUCAO = 7844 'Parametro lCodigo
'Movimento %l não é Requisição de Material para Produção.
Public Const ERRO_OP_NAO_PREECHIDA = 7845 'Sem Parametro
'Ordem de Produção não foi preenchida.
Public Const ERRO_QUANTIDADE_SEM_PRODUTO = 7846 'Sem Parametro
'Produto não foi preenchido. A Quantidade deve estar acompanhada do Produto
Public Const ERRO_PRODUTO_NAO_E_KIT = 7847 'Parametro sCodigo
'O Produto %s não é um Kit.
Public Const ERRO_PRODUTO_NAO_PARTICIPA_KIT = 7848 'Parametros sProduto , sProdutoOP
'O Produto %s não faz parte do Kit do Produto %s.
Public Const ERRO_OPCODIGO_NAO_PREENCHIDO = 7849 'Parametro iLinhaGrid
'O campo Ordem de Produção do ítem %s do Grid de Requisição de Material não foi preenchido.
Public Const ERRO_ITEM_OP_PRODUZIDO = 7850 'Parametros sProduto , sCodigoOP
'O Produto %s da Ordem de Produção %s já foi totalmente produzido.
Public Const ERRO_ITEM_OP_NAO_E_KIT = 7851 'Parametros sProduto , sCodigoOP
'O Produto %s faz parte da Ordem de Produção %s e não é um Kit.
Public Const ERRO_KIT_SEM_PRODUTO_RAIZ = 7852 'Parametros sProduto
'O produto %s não é raiz de um Kit.
Public Const ERRO_KIT_SEM_PRIMEIRO_NIVEL = 7853 'Parametro sProduto
'O produto %s é um Kit e não possui primeiro nível.
Public Const ERRO_CODIGO_INCOMPATIVEL_PSAIDA = 7854 'Parametro lCodigo
'O Movimento com o código %l já existe mas não é Tipo de Produção Saída .
Public Const ERRO_LOCK_EMPENHO1 = 7855 'Parametros sProduto , iAlmoxarifado , lNumIntDoc
'Erro na tentativa de fazer "lock" na tabela de Empenho com Produto %s, Almoxarifado %i e Número interno do item da OP %l.
Public Const ERRO_ATUALIZACAO_EMPENHO1 = 7856 'Parametros sProduto , iAlmoxarifado , lNumIntDoc
'Erro na tentativa de atualizar um registro na tabela de Empenho com Produto %s, Almoxarifado %i e Número interno do item da OP %l.
Public Const ERRO_EMPENHO_INEXISTENTE = 7857 'Parametros sProduto , iAlmoxarifado , lNumIntDoc
'O Empenho com com Produto %s, Almoxarifado %i e Número interno do item da OP %l, não está cadastrado no Banco de Dados.
Public Const ERRO_EXCLUSAO_RECEB_MAT_CLI_NF = 7858 'Parametros lCliente, iFilialCli, dtDataEntrada, sSerie, lNumNotaFiscal
'Não é possível excluir o Recebimento de Material com os dados: Cliente %l, Filial Cliente %i, Data Entrada %dt, Série %s, Nota Fiscal %l. A Nota Fiscal correspondente está registrada.
Public Const ERRO_RECEB_MAT_CLI_NAO_CADASTRADO = 7859 'Parametros lRecebimento
'O Recebimento de Material %l não está cadastrado no Banco de Dados.
Public Const ERRO_ATUALIZACAO_RECEB_MAT_CLI_NF = 7860 ' Parametros lCliente, iFilialCli, dtDataEntrada, sSerie, lNumNotaFiscal
'Não é possível alterar Recebimento de Material com dados: Cliente %l, Filial Cliente %i, Data Entrada %dt, Série %s, Nota Fiscal %l. A Nota Fiscal correspondente está registrada.
Public Const ERRO_INSERCAO_RECEB_MAT_CLI_NF = 7861 ' Parametros lCliente, iFilialCli, dtDataEntrada, sSerie, lNumNotaFiscal
'Não é possível inserir Recebimento de Material com dados: Cliente %l, Filial Cliente %i, Data Entrada %dt, Série %s, Nota Fiscal %l. A Nota Fiscal correspondente está registrada.
Public Const ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA = 7862 ' Parametros sSerie, lNumNotaFiscal
'Nota Fiscal Interna de Saída com série %s e número %l não está cadastrada no Banco de Dados.
Public Const ERRO_SERIE_NUMERO_ORIGINAL_FALTANDO = 7863 ' Sem parametros
'Para estabelecer vínculo com Nota Fiscal Original é necessário preencher Série e Número.
Public Const ERRO_NF_NAO_CADASTRADA = 7864 'Parametros lNumIntDoc
'O Nota Fiscal com NumIntDoc %l não está cadastrado no Banco de Dados.
Public Const ERRO_TIPODOC_NAO_RECEBCLI = 7865 'Parametro iTipoNFiscal
'Tipo de Documento %i não é Recebimento de Cliente.
Public Const ERRO_NF_INT_SAIDA_NAO_DEV_NAO_CADASTRADA = 7866 'sSerie, lNumero
'Nota Fiscal Interna de Saída que não é devolução, com série %s e número %l não esta cadastrada no Banco de Dados.
Public Const ERRO_TIPO_NFISCAL_NAO_INFORMADO = 7867 'Sem parametros
'Tipo de Nota Fiscal não foi selecionado.
Public Const ERRO_MOVESTOQUE_NAO_CADASTRADO = 7868 'Parametro lNumIntDocOrigem
'Movimento de Estoque com NumIntDocOrigem %l não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_MOVIMENTOESTOQUE3 = 7869 ' Sem Parametros
'Erro de Leitura na Tabela MovimentoEstoque.
Public Const ERRO_RECEB_MAT_FORN_NAO_CADASTRADO = 7870 'Parametros lFornecedor, iFilialForn, lNumNotaFiscal, sSerie, dtDataEntrada
'Recebimento de Material com dados: Fornecedor %l, Filial Fornecedor %i, Número Nota Fiscal %l, Série Nota Fiscal %s, Data Entrada %dt, não está cadastrado no Banco de Dados.
Public Const ERRO_PESOBRUTO_MENOR_PESOLIQ = 7871 'Parametros dPesoBruto, dPesoLiq
'O Peso Bruto %d não pode ser menor que o Peso Líquido %d.
Public Const ERRO_ATUALIZACAO_RECEB_MAT_FORN_NF = 7872 ' Parametros lFornecedor, iFilialForn, dtDataEntrada, sSerie, lNumNotaFiscal
'Não é possível alterar Recebimento de Material com dados: Fornecedor %l, Filial Fornecedor %i, Data Entrada %dt, Série %s, Nota Fiscal %l. A Nota Fiscal correspondente está registrada.
Public Const ERRO_EXCLUSAO_RECEB_MAT_FORN_NF = 7873 'Parametros lFornecedor, iFilialForn, dtDataEntrada, sSerie, lNumNotaFiscal
'Não é possível excluir o Recebimento de Material com os dados: Fornecedor %l, Filial Fornecedor %i, Data Entrada %dt, Série %s, Nota Fiscal %l. A Nota Fiscal correspondente está registrada.
Public Const ERRO_VALORTOTAL_RECEB_NEGATIVO = 7874 'Sem parametros
'Valor Total do Recebimento é negativo.
Public Const ERRO_ALMOXARIFADOITEM_NAO_PREENCHIDO = 7875 'Parametro iItem
'Almoxarifado do ítem %i do Grid Ítens não foi preenchido.
Public Const ERRO_QUANTIDADEITEM_NAO_PREENCHIDA = 7876 'Parametro iItem
'Quantidade do ítem %i do Grid Ítens não foi preenchida.
Public Const ERRO_VALORUNITARIOITEM_NAO_PREENCHIDO = 7877 'Parametro iItem
'Valor Unitário do ítem %i do Grid Ítens não foi preenchido.
Public Const ERRO_ITENSRECEB_NAO_INFORMADOS = 7878 'Sem parametros
'Não há Ítens de Recebimento de Material no Grid
Public Const ERRO_TIPODOC_NAO_RECEBFORN = 7879 'Parametro iTipoNFiscal
'Tipo de Documento %i não é Recebimento de Fornecedor.
Public Const ERRO_RECEB_NAO_CADASTRADO = 7880 'Parametro lNumNotaFiscal
'Recebimento %l não está cadastrado no Banco de Dados.
Public Const ERRO_SERIE_MAIOR = 7881 'Sem Parametros
'Série tem limite de 3 dígitos.
Public Const ERRO_CODIGO_INCOMPATIVEL_CONSUMO = 7882 'Parametro lCodigo
'O Movimento com o código %l já existe mas não é Tipo de Consumo .
Public Const ERRO_EXCLUSAO_TIPOSMOVEST = 7883 'Parametro: iCodigo
'Erro na exclusão do tipo de código %i na tabela de Tipos de Movimentos de Estoque.
Public Const ERRO_INCLUSAO_TIPOSMOVEST = 7884 'Parametro: iCodigo
'Erro na inclusão do tipo de código %i na tabela de Tipos de Movimentos de Estoque.
Public Const ERRO_ATUALIZACAO_TIPOSMOVEST = 7885 'Parametro: iCodigo
'Erro na atualização do tipo de código %i na tabela de Tipos de Movimentos de Estoque.
Public Const ERRO_TIPOSMOVEST_INEXISTENTE = 7886 'Parametro: iCodigo
'O tipo de código %i não existe na tabela de Tipos de Movimentos de Estoque.
Public Const ERRO_TIPOSMOVEST_NAOEDITAVEL = 7887 'Parametro: iCodigo
'O tipo de código %i não é editável. Não pode ser alterado ou excluído.
Public Const ERRO_EXCLUSAO_TIPOSMOVEST1 = 7888 'Parametro: iCodigo
'Não é possível excluir o tipo de código %i na tabela de Tipos de Movimentos de Estoque, pois já foi utilizado na tabela de Movimentos de Estoque.
Public Const ERRO_ENTRADAOUSAIDA_NAO_PREENCHIDA = 7889 'Parametro: iCodigo
'O preenchimento do campo tipo é obrigatório.
Public Const ERRO_TIPOMOVEST_NAO_PREENCHIDO = 7890 'Sem parametros
'O preenchimento do tipo de movimento é obrigatório.
Public Const ERRO_SIGLA_TIPOPRODUTO_NAO_PREENCHIDA = 7891 'Sem parâmetros
'O preenchimento do campo Sigla é obrigatório.
Public Const ERRO_CODIGO_INCOMPATIVEL_TRANSFERENCIA = 7892 'Parametro lCodigo
'O Movimento com o código %l já existe mas não é Tipo de Transferência .
Public Const ERRO_TRANSF_OD = 7893 'Parametros sTipoOrigem , sTipoDestino
'Não é possível efetuar a transferência do Tipo Origem %s para o Tipo Destino %s.
Public Const ERRO_LEITURA_ESTOQUEPRODUTO3 = 7894 'Sem parâmetros
'Erro na leitura da tabela de EstoqueProduto.
Public Const ERRO_PRODUTO_REPETIDO = 7895
'Produto repetido.
Public Const ERRO_SOMA_QUANTIDADES_MAIOR_MAXIMO = 7896
'Quantidade maior que o estoque máximo.
Public Const ERRO_PRODUTO_JA_EXISTENTE = 7897 'sProduto, iItem
'O produto %s já participa deste Pedido de Venda no Item %i.
Public Const ERRO_VALOR_COMISSAO_MAIOR_VALORBASE = 7898 'Parametros: dValorComissoa, dValorBase
'Valor de Comissao %d não pode ser superior ao Valor Base %d.
Public Const ERRO_VALOR_COMISSAO_EMISSAO_MAIOR = 7899
'Valor emissão maior do que o valor da comissão.
Public Const ERRO_TOTAL_PERCENTUAIS_MAIOR_100 = 7900
'Total dos percentuais de Comissao deve ser menor do que 100%.
Public Const ERRO_FALTA_PARCELA_COBRANCA = 7901
'Não existe informação de Parcelas em Cobrança.
Public Const ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA = 7902 'Parâmetro: iParcela
'Em Cobrança, Parcela número %i não teve Data de Vencimento preenchida.
Public Const ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR = 7903 'Parâmetro: iParcela, dtDataVencimento, dtDataEmissao
'Em Cobrança, Parcela número %i tem Data de Vencimento %dt anterior à Data Emissao do Pedido %dt.
Public Const ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA = 7904
'Data de Vencimento das Parcelas em Cobrança devem estar em ordem crescente.
Public Const ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO = 7905 'Parâmero: iParcela
'Em Cobrança, Parcela número %i não teve valor preenchido.
Public Const ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA = 7906
'Valor Total menos Imposto de Renda na Fonte não é igual à soma dos valores das Parcelas em Cobrança.
Public Const ERRO_LEITURA_RESERVA = 7909 'Parametros iFilialEmpresa, lDocOrigem, sProduto, iAlmoxarifado, iNumItemDocOrigem
'Ocorreu um erro na leitura da tabela de Reservas. FilialEmpresa = %i, DocOrigem = %l, Produto = %s, Almoxarifado = %i, ItemDocOrigem = %i.
Public Const ERRO_RESERVA_NAO_CADASTRADA = 7910 'Parametros iFilialEmpresa, lDocOrigem, sProduto, iAlmoxarifado, iNumItemDocOrigem
'A reserva em questão não está cadastrada. FilialEmpresa = %i, DocOrigem = %l, Produto = %s, Almoxarifado = %i, ItemDocOrigem = %i.
Public Const ERRO_QUANT_FATURADA_MAIOR_RESERVADA = 7911 'Parametros dQuantFaturada, dQuantReservada, iFilialEmpresa, lDocOrigem, sProduto, iAlmoxarifado, iNumItemDocOrigem
'A Quantidade Faturada = %d ultrapassa a Quantidade Reservada = %d. Reserva -> FilialEmpresa = %i, DocOrigem = %l, Produto = %s, Almoxarifado = %i, ItemDocOrigem = %i.
Public Const ERRO_EXCLUSAO_RESERVA = 7912 'Parametros iFilialEmpresa, lDocOrigem, sProduto, iAlmoxarifado, iNumItemDocOrigem
'Ocorreu um erro na exclusão da reserva identificada por: FilialEmpresa = %i, DocOrigem = %l, Produto = %s, Almoxarifado = %i, ItemDocOrigem = %i.
Public Const ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO = 7913  'Parametros sProduto, iAlmoxarifado
'O Produto %s não está associado ao Almoxarifado %i na tabela EstoqueProduto.
Public Const ERRO_ITEMPV_QAF_ZERO = 7915 'parametros: numero do pedido e descricao do produto
'No pedido %l, o item %s já foi totalmente faturado mas não está marcado como "atendido"
Public Const ERRO_FATURAR_PEDIDO_OUTRA_FILIAL = 7916 'parametros: numero do pedido
'O pedido %l deve ser faturado por outra filial que não a corrente
Public Const ERRO_DATAEMISSAODE_MAIOR_DATAEMISSAOATE = 7917 'Parâmetros: DataEmissaoDe.Text, DataEmissaoAte.Text
'A Data Emissão De %s deve ser menor que Data Emissão Até %s.
Public Const ERRO_DATAENTREGADE_MAIOR_DATAENTREGAATE = 7918 'Parâmetros: DataEntregaDe.Text, DataEntregaAte.Text
'A Data Entrega De %s deve ser menor que Data Entrega Até %s.
Public Const ERRO_NFISCAL_EDICAO_NAO_MARCADA = 7919 'Sem parâmetro
'Deve haver uma Nota Fiscal marcada para Edição.
Public Const ERRO_TIPO_MOVIMENTOESTOQUE_INVALIDO = 7920 'Parametro iTipoMovEstoque
'O Tipo de Movimento de Estoque %i não é válido para esta transação.
Public Const ERRO_LEITURA_SLDDIAESTALM = 7921 'Parametros iAlmoxarifado, sProduto, sData
'Ocorreu um erro na leitura da tabela de saldos diários de estoque por almoxarifado. Almoxarifado=%i, Produto=%s, Data=%s.
Public Const ERRO_INSERCAO_SLDDIAESTALM = 7922 'Parametros iAlmoxarifado, sProduto, sData
'Ocorreu um erro na inclusão de registro na tabela de saldos diários de estoque por almoxarifado. Almoxarifaddo=%i, Produto=%s, Data=%s.
Public Const ERRO_LOCK_SLDDIAESTALM = 7923  'Parametros iAlmoxarifado, sProduto, sData
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos diários de estoque por almoxarifado. Almoxarifado=%i, Produto=%s, Data=%s.
Public Const ERRO_ATUALIZACAO_SLDDIAESTALM = 7924  'Parametros iAlmoxarifado, sProduto, sData
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos diários de estoque por almoxarifado. Almoxarifado=%i, Produto=%s, Data=%s.
Public Const ERRO_ATUALIZACAO_SLDMESESTALM = 7925  'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque por almoxarifado. Ano=%i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_LEITURA_SLDMESESTALM = 7926 'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque por almoxarifado. Ano = %i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_LOCK_SLDMESESTALM = 7927 'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque por almoxarifado. Ano = %i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_PRODUTO_SEM_ALMOX_PADRAO = 7928 'Parametro: sProduto
'Produto %s não tem almoxarifado padrão.
Public Const ERRO_LEITURA_TABELASDEPRECOITENS2 = 7929 'Parametros: iFilialEmpresa, iCodTabela, sCodProduto
'Erro na leitura da tabela TabelasDePrecoItens com Filial %i, com código da Tabela %i e código do Produto %s.
Public Const ERRO_TABELAPRECOITEM_INEXISTENTE = 7930 'Parametros: iCodTabela, sCodProduto
'O Item de Tabela de Preço com código da Tabela %i e código do Produto %s não está cadastrada no Banco de Dados.
Public Const ERRO_EXCLUSAO_TABELASDEPRECOITENS1 = 7931 'Parametro: iCodTabela, sCodProduto
'Erro na tentativa de excluir o registro da tabela de TabelasDePrecoItens com código da Tabela %i e código do Produto %s.
Public Const ERRO_CUSTO_STANDARD_MOVIMENTO = 7932 'Parametros: sMes, iAno
'Não é possivel mudar o Custo Standard pois já houve movimento de Estoque no Mes %s, Ano %i.
Public Const ERRO_CUSTO_STANDARD_PREENCHIDO = 7933
'Não é permitido preencher o Custo Standard, pois a apropriação do custo do produto não é de Custo Standard.
Public Const ERRO_PRODUTO_ALTERACAO_COMPRA = 7934 'Parâmetro: sCodigo
'Não é permitida a alteração do Produto %s de Comprado para Produzido ou vice-versa.
Public Const ERRO_PRODUTO_ALTERACAO_APROPRIACAO = 7935 'Parâmetro: sCodigo
'Não é permitida a alteração da apropriação do Produto %s.
Public Const ERRO_PRODUTO_NAO_CORRESPONDE = 7936 'Parametros sCodProduto,iNumItem
'O Produto %s informado não corresponde ao produto do Item %i encontrado no Banco de Dados.
Public Const ERRO_BLOQUEIO_INVALIDO = 7938 'Parametros iSequencial,lPedidoDeVenda,iFilialEmpresa
'Bloqueio %i do Pedido de Venda %l da FilialEmpresa %i é Bloqueio de Estoque e não é Bloqueio Parcial . Não Pode haver Reserva junto com esse tipo de Bloqueio.
Public Const ERRO_ATUALIZACAO_BLOQUEIOSPV = 7939 'Parametro lPedidoDeVenda
'Erro na Atualização do Bloqueio do Pedido de Venda %l.
Public Const ERRO_INCLUSAO_BLOQUEIOSPV = 7940 'Parametro lPedidoDeVenda
'Erro ao tentar inserir Bloqueio do Pedido de Venda %l.
Public Const ERRO_INCLUSAO_RESERVA = 7941 'Sem parametros
'Erro ao tentar inserir Reserva.
Public Const ERRO_ATUALIZACAO_RESERVA = 7942 ' Parametro lCodReserva
'Erro na atualização da Reserva %l.
Public Const ERRO_QUANTIDADE_INSUFICIENTE = 7943 'Parametros sCodProduto, iCodAlmoxarifado, dQuantDisponivel, dQuantReservar
'A Quantidade disponível do Produto %s no Almoxarifado %i é %d que é menor do que a quantidade reservada %d.
Public Const ERRO_ITEM_NAO_COINCIDE = 7944 'Sem Parametro
'Esta reserva já está cadastrada para um produto almoxarifado diferente do informado nesta tela para o pedido em questão.
Public Const ERRO_DOCORIGEM_NAO_COINCIDE = 7945 'Parametros lPedidoBanco de Dados,lPedidoTela
'Esta reserva está cadastrada no Banco de Dados.O Pedido de Venda %l lido no Banco de Dados não coincide com o Pedido %l da Tela.
Public Const ERRO_PRODUTO_NAO_COINCIDE = 7946  'Parametros sProdutoBanco de Dados,sProdutoTela
'Esta reserva está cadastrada no Banco de Dados. O produto %s lido no Banco de Dados não coincide com o produto %s da Tela.
Public Const ERRO_ALMOXARIFADO_NAO_COINCIDE = 7947 'Parametros iAlmoxarifadoBanco de Dados,iAlmoxarifadoTela
'Esta reserva está cadastrada no Banco de Dados.O Almoxarifado %i lido no Banco de Dados não coincide com Almoxarifado %i da Tela.
Public Const ERRO_TIPODOC_NAO_COINCIDE = 7948 'Parametrso iTipoDocBanco de Dados,iTipoDocTela
'Não pode ser alterado a origem da reserva, se esta reserva foi feita por um Pedido de Venda.
Public Const ERRO_ITEMPEDIDO_NAO_CADASTRADO = 7949 'Parametro lCodReserva
'Item de Pedido de Venda associado à Reserva %l não está cadastrado no Banco de Dados
Public Const ERRO_PEDIDOVENDA_INEXISTENTE = 7950 'Parametro lCodReserva
'Pedido de Venda associado à Reserva %l não está cadastrado no Banco de Dados
Public Const ERRO_OBJPRODUTO_NAO_CADASTRADO = 7951 ' Parametro sCodProduto
'Produto %s que esta relacionado à Reserva não está cadastrado no Banco de Dados
Public Const ERRO_OBJALMOXARIFADO_NAO_CADASTRADO = 7952 'Parametro iCodAlmoxarifado
'Almoxarifado %i que está relacionado à Reserva não está cadastrado no Banco de Dados
Public Const ERRO_OBJESTOQUEPRODUTO_NAO_CADASTRADO = 7953 ' Parametros sCodProduto e iCodAlmoxarifado
'Estoque do Produto %s no Almoxarifado %i que está relacionado à Reserva não está cadastrado no Banco de Dados
Public Const ERRO_OBJPEDIDODEVENDA_NAO_CADASTRADO = 7954 'Parametro lCodPedidoDeVenda
'Pedido de Venda %l que está relacionado à Reserva não está cadastrado no Banco de Dados
Public Const ERRO_LOCK_RESERVA = 7955 'Sem parametros
'Não conseguiu fazer o Lock na Tabela Reserva.
Public Const ERRO_RESERVAITEM_INEXISTENTE = 7956 'Parametros iItem e iCodAlmoxarifado
'Não existe reserva do Item %i no Almoxarifado %i
Public Const ERRO_ESTOQUEPRODUTO_INEXISTENTE = 7957 'Parametro sProduto
'O Produto %s não consta do Almoxarifado escolhido
Public Const ERRO_FILIALEMPRESA_NAO_CADASTRADA = 7958 'Parametro iFilialEmpresa
'FilialEmpresa com código %i não está cadastrada
Public Const ERRO_PEDIDO_SEM_PRODUTO = 7959 'Parametro sCodProduto
'O Produto %s não faz parte do Pedido de Venda associado.
Public Const ERRO_DOCUMENTO_NAO_PREENCHIDO = 7960 'Sem parametros
'Número do Pedido de Venda não foi preenchido
Public Const ERRO_PRODUTO_NAO_INFORMADO = 7961 'Sem parametros
'Produto não foi informado
Public Const ERRO_QUANTIDADE_RESERVADA_MAIOR = 7962 'Sem parametros
'A quantidade que está sendo reservada não pode ultrapassar a quantidade disponível no almxarifado.
Public Const ERRO_PEDIDOVENDA_NAO_CADASTRADA = 7963 ' Parametro lCodPedidoVenda
'Pedido de Venda com código %l não está cadastrado no Banco de Dados
Public Const ERRO_DATA_VALIDADE_MENOR = 7964 'Sem Parametros
'Data de Validade não pode ser inferior à Data de Reserva
Public Const ERRO_PRODUTO_NAO_ADMITE_RESERVA = 7965 'Parametro sCodProduto
'Produto %s não admite reserva
Public Const ERRO_ITENSPV_NAO_UTILIZAM_PRODUTO = 7966 'Parâmetro: lPedidoVenda, sProduto
'O Pedido %l não possui itens que utilizem o produto %s.
Public Const ERRO_PRODUTO_CONTROLE_NAO_RESERVA = 7967 'Parâmetro: sProduto
'O Controle de Estoque do Produto %s não é do tipo reserva.
Public Const ERRO_INSERCAO_TRIBPV = 7968
'Erro na inserção de registro na tabela de Tributação de Pedido de Venda.
Public Const ERRO_INSERCAO_TRIBCOMPLPV = 7969
'Erro na inserção de registro na tabela de Complemento de Tributação de Pedido de Venda.
Public Const ERRO_INSERCAO_TRIBITEMPV = 7970
'Erro na inserção de registro na tabela de Tributação de Itens de Pedido de Venda.
Public Const ERRO_FILIAL_FATURAMENTO_NAO_PREENCHIDA = 7971
'A Filial de Faturamento deve ser informada.
Public Const ERRO_EXCLUSAO_TRIBPEDIDO = 7972
'Erro na exclusão de registro da tabela de Tributação de Pedidos de Venda.
Public Const ERRO_TRIBPEDIDO_NAO_ENCONTRADA = 7973 'lcodPedido
'Não foi encontrado nenhum registro de Tributação para o Pedido %l.
Public Const ERRO_LEITURA_TRIBPEDIDO = 7974
'Erro na leitura da tabela de Tributação de Pedido de Venda.
Public Const ERRO_EXCLUSAO_TRIBITEMPEDIDO = 7975
'Erro na exclusão na tabela de Tributação de Itens de Pedido de Venda.
Public Const ERRO_EXCLUSAO_TRIBCOMPLPEDIDO = 7976
'Erro na exclusão  de Itens de Pedido de Venda da tabela de Tributação.
Public Const ERRO_LEITURA_TRIBITEMPEDIDO = 7977
'Erro na leitura da tabela de Tributação de Itens de Pedido de Venda.
Public Const ERRO_LEITURA_TRIBCOMPLPEDIDO = 7978
'Erro na leitura da tabela de Tributação de Complemento de Pedido de Venda.
Public Const ERRO_DATADESCONTO_MAIOR_DATAVENCIMENTO = 7979 'Parâmetros: dtDataDesconto, dtDataVencimento
'Data de desconto da parcela %dt não pode ultrapassar data de vencimento da parcela %dt.
Public Const ERRO_DATAS_DESCONTO_IGUAIS = 7980 'iDesconto, iDescontoOutro, iParcela
'Não é possível dois descontos terem a mesma Data Limite. As datas dos desconto %i e %i da Parcela %i são iguas.
Public Const ERRO_NENHUM_ITEM_TRIB_SEL = 7981 'sem parametros
'Algum item do pedido tem que estar selecionado.
Public Const ERRO_PV_VINCULADO_OP = 7982 'Parâmetros: loCodPedido, lNumIntItemPV, lNumIntItemOP
'Pedido de Venda com código %l tem item com número interno %l vinculado ao ítem de Ordem de Producao com número interno %l
Public Const ERRO_ITEM_PV_FATURADO = 7983 'Parametros: lCodPedido, lNUmIntItem, dQuantFaturada.
'Não é possível excluir Pedido Venda com código %l pois tem ítem %l com quantidade %d faturada.
Public Const ERRO_INCLUSAO_BLOQUEIOSPVBAIXADOS = 7984 'Parametro: lCodPedido
'Ocorreu um erro na tentativa de inserir um registro na tabela de Bloqueios de Pedido de Venda Baixados. Pedido = %l.
Public Const ERRO_EXCLUSAO_PARCELASPV = 7985 'Parametro: lCodPedido
'Ocorreu um erro na tentativa de excluir um registro da tabela de Parcelas de Pedido de Venda. Pedido = %l.
Public Const ERRO_INSERCAO_PARCELASPV_BAIXADOS = 7986
'Ocorreu um erro na tentativa de inserir um registro na tabela de Parcelas de Pedido de Venda Baixados. Pedido = %l.
Public Const ERRO_INSERCAO_COMISSOESPV_BAIXADOS = 7987 'Parametro: lCodPedido
'Ocorreu um erro na tentativa de inserir um registro na tabela de Comissões de Pedido de Venda Baixados. Pedido = %l.
Public Const ERRO_ITEM_OP_VINCULADO_ITEM_PV = 7988 'lCodPedido, lNumIntItemPV, lNumIntItemOP
'O Pedido de Venda com Código %l tem item com número interno %l a ser excluído está
'vinculado a Item de Ordem de Produção com o número interno %l.
Public Const ERRO_ITEMPV_NAO_CADASTRADO = 7989 'lNumIntItem
'O Item de Pedido de Venda com o número interno = %l não está cadastrado.
Public Const ERRO_QUANT_FATURADA_ALTERADA = 7990
'A quantidade faturada do item %i do Pedido de Vendas %l foi alterada.
Public Const ERRO_ITEMPV_NAO_ENCONTRADO = 7991 'lNumIntItem
'O item do Pedido de Venda com número interno = %l não foi encontrado.
Public Const ERRO_EXCLUSAO_RESERVASPV = 7992 'lCodPedido
'Erro na exclusão das Reservas do Pedido de Venda %l.
Public Const ERRO_PEDVENDA_BAIXADO_ALTERACAO = 7993
'O Pedido de Venda com o código %l está baixado. Não pode ser alterado.
Public Const ERRO_PEDVENDA_BAIXADO_EXCLUSAO = 7994
'O Pedido de Venda com o código %l está baixado. Não pode ser excluído.
Public Const ERRO_INSERCAO_RESERVASPV = 7995 'Parâmetro: lCodPedido
'Erro na inserção das reservas do Pedido de Venda %l na Tabela de Reservas.
Public Const ERRO_INSERCAO_ITENSPV = 7996
'Erro na inserção na tabela de Itens de Pedido de Venda.
Public Const ERRO_INSERCAO_PARCELASPV = 7997
'Erro na inserção na tabela de Parcelas de Pedido de Venda.
Public Const ERRO_INSERCAO_COMISSOESPV = 7998
'Erro na inserção na tabela de Comissões de Pedidos de Venda.
Public Const ERRO_AUSENCIA_ITENS_PV = 7999
'Não existem itens para o Pedido de Venda
Public Const ERRO_PEDVENDA_SEM_QUANTIDADE = 8500
'Não é possível gravar Pedido de Venda que não tenha quantidade de Produto livre de cancelamento.
Public Const ERRO_LEITURA_TIPOSDEPEDIDO = 8501 'sSigla
'Erro na tentativa de leitura do Tipo de Pedido com a Sigla %s na Tabela TiposDePedido.
Public Const ERRO_PRODUTO_ITEM_VAZIO = 8502
'não é possível deixar um Item do Pedido sem Produto.
Public Const ERRO_VALOR_DESCONTO_MAIOR = 8503 'dDesconto, dValorProdutos
'Desconto = %d não pode ultrapassar ValorProdutos = %d.
Public Const ERRO_PRODUTO_PV_NAO_ALTERAVEL = 8504
'Não é possível alterar Produto de um item de Pedido de Venda.
Public Const ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA = 8505
'Quantidade Pedida não pode ser inferior à Quantidade Cancelada
Public Const ERRO_QUANT_FATURADA_SUPERIOR = 8506
'Quantidade Pedida - Quantidade Cancelada não pode ser inferior à Quantidade Faturada.
Public Const ERRO_RESERVA_NAO_DECIDIDA = 8507 'sProduto
'Não foi tomada uma decisão sobre a reserva do Produto %s.
Public Const ERRO_PEDIDOVENDA_JA_CADASTRADO = 8508 'lCodigo
'O Pedido de Venda %l já está cadastrado. Caso deseje alterá-lo, chame-o pelo browse ou pelo sistema de setas.
Public Const ERRO_DATAEMISSAO_MAIOR_DATAENTREGA = 8509 'dtDataEntrega, dtDataEmissao
'Data de Entrega = dt% não pode ser inferior a Data de Emissão dt%.
Public Const ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA = 8510 'dtDataReferencia, dtDataEmissao
'Data de Referência = dt% não pode ser inferior a Data de Emissão do Pedido = dt%.
Public Const ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA = 8511 'Parâmetros: sDataVencimento, sDataEmissao, iParcela
'A Data de Vencimento %s da Parcela %i é menor do que a Data de Referencia %s.
Public Const ERRO_VALOR_COMISSAO_EMISSAO_MAIOR1 = 8512 'Parametros: dValorComissaoEmissao, dValorComissao
'Valor de Comissão na Emissão %d não pode ser superior ao Valor da Comissao %d.
Public Const ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO = 8513 'Parametro: iCodigo
'O Tipo de Bloqueio %i não está cadastrado
Public Const ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1 = 8514 'Parametro: sTipoBloqueio
'O Tipo de Bloqueio %s não está cadastrado
Public Const ERRO_TIPOBLOQUEIO_BLOQUEIO_NAO_INFORMADO = 8515 'Parametro: iLinha
'Tipo de Bloqueio do bloqueio %i do Grid de bloqueios não está preenchido.
Public Const ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO = 8516
'Número do Pedido não foi preenchido.
Public Const ERRO_FILIALCLIENTE_NAO_INFORMADA = 8517
'Filial de Cliente não foi preenchida.
Public Const ERRO_VALORTOTAL_PV_NEGATIVO = 8518
'Valor Total do Pedido de Venda é negativo.
Public Const ERRO_FILIALCLIENTE_NAO_CADASTRADA1 = 8519 'Parametros: Cliente.Text, objFilialCliente.iCodFilial
'Filial de Cliente não está cadastrada. Dados: Nome reduzido Cliente = Cliente.Text, Código da Filial = %i.
Public Const ERRO_LEITURA_ALMOXPRODFILIAL = 8520 'Sem Parametro
'Erro na leitura das Tabelas: Almoxarifado e EstoqueProduto.
Public Const ERRO_ALMOXARIFADO_FILIAL_DIFERENTE = 8521 'Parâmetros: iCodigo, giFilialEmpresa
'O Almoxarifado %i não pertence à Filial %i.
Public Const ERRO_ALMOXARIFADO_FILIAL_DIFERENTE1 = 8522 'Parâmetros: sNomeReduzido, giFilialEmpresa
'O Almoxarifado %s não pertence a Filial da Empresa %i.
Public Const ERRO_ITEM_PV_VINCULADO_ITEM_OP = 8523 'Parametros: iItemAtual, lItemOPNumIntDoc
'Ítem %i do Pedido Venda não pode ser excluído pois está ligado a Ítem de Ordem de Produção com número interno %l.
Public Const ERRO_DATA_PRODUTO_EXISTENTE = 8524 'sProduto, dtdata
'Erro já existe um Kit com o Produto %s e a data %dt na Tabela de Kit, não pode haver duplicidade.
Public Const ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2 = 8525 'Sem parametros
'É obrigatório o preenchimento do produto para mostrar a lista de versões.
Public Const ERRO_ALMOXARIFADO_PADRAO = 8526
'Para desmarcar um Almoxarifado Padrão marque outro Almoxarifado como Padrão.
Public Const ERRO_EXCLUSAO_ALMOXARIFADO_PADRAO = 8527
'Para excluir um Almoxarifado Padrão marque outro Almoxarifado como Padrão.
Public Const ERRO_SALDO_QUANTIDADE_NAO_NULOS = 8528 'Parâmetro: sProduto
'Há movimentos de estoque para o Produto %s. Quantidade e Valor Iniciais de Estoque deverão ser nulos.
Public Const ERRO_INSERCAO_SLDMESESTALM = 8529
'Erro na inserção de registro na tabela SaldoMesEstAlm.
Public Const ERRO_EXCLUSAO_ESTOQUEPRODUTO_SALDO_ZERO = 8530
'Não é possível excluir essa combinação Produto-Almoxarifado pois o saldo inicial afeta a valoração de movimentos existentes desse produto.
Public Const ERRO_QUANTIDADE_VALOR_INICIAL = 8531
'O produto não possui nenhum registro em Estoque produto. A quantidade e o valor devem ser positivos.
Public Const ERRO_SLDMESESTALM_INEXISTENTE = 8532 'Parametros: iAno, iAlmoxarifado, sProduto
'Não existe registro de saldos mensais de estoque Almoxarifado (SldMesEstAlm) com os dados a seguir. Ano=%i, Almoxarifado=%i, Codigo do Produto=%s.
Public Const ERRO_LEITURA_SLDDIAESTALM1 = 8533 'Parametros: iAlmoxarifado, sProduto
'Ocorreu um erro na leitura da tabela de saldos diários de estoque por almoxarifado. Almoxarifado=%i, Produto=%s.
Public Const ERRO_LEITURA_SLDDIAEST1 = 8534 'Parametros: iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos diários de estoque. Filial =%i, Produto=%s.
Public Const ERRO_EXCLUSAO_SLDDIAESTALM = 8535 'Parâmetros: iAlmoxarifado, sProduto
'Ocorreu um erro na exclusão de registros da tabela de Saldos Diários de Estoque de Almoxarifado com os dados: Almoxarifado = %i, Produto = %s)
Public Const ERRO_EXCLUSAO_SLDMESESTALM = 8536 'Parâmetros: iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na exclusão de registros da tabela de Saldos Mensais de Estoque de Almoxarifado com os dados: iAno =%i, Almoxarifado = %i, Produto = %s)
Public Const ERRO_PRODUTO_COM_MOVIMENTOS = 8537 'Parâmetro: sProduto
'Há Movimentos de Estoque para o Produto %s. Quantidade, Valor e Data Iniciais não podem ser alterados.
Public Const ERRO_LEITURA_EMPENHO2 = 8538 'sem parametro
'Erro na leitura de registro da tabela de Empenhos
Public Const ERRO_KIT_INEXISTENTE1 = 8539 'Sem parâmetros: sProdutoRaiz, dtData
'O Produto %s não possui Kit com Data anterior à %dt.
Public Const ERRO_LOTEINVPEN_NAO_CADASTRADO = 8540
'O lote %i não está cadastrado, por isso não existe itens para ele.
Public Const ERRO_SEGMENTO_PRODUTO_INVALIDO = 8541 'Parametro sCodigo
'Esperava o segmento produto e está tentando gravar o segmento %s.
Public Const ERRO_PRODUTO_KIT_NAO_BASICO = 8542
'O Produto em questão participa de um ou mais kits em que é produto básico.
Public Const ERRO_PRODUTO_KIT_NAO_INTERMEDIARIO = 8543
'O Produto em questão participa de um ou mais kits em que é produto intermediário.
Public Const ERRO_SALDO_MAT_BENEF = 8544 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto em poder de terceiros para beneficiamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_SALDO_MAT_BENEF3 = 8545 'Parametros sProduto, iAlmoxarifado, Saldo
'O Saldo do Produto de terceiros em beneficiamento é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_LEITURA_ITENSNFISCALBAIXADAS2 = 8546 'Sem Parametros
'Ocorreu um erro na leitura da tabela de itens de nota fiscal baixada.
Public Const ERRO_PRODUTO_ITENSNFISCALBAIXADAS = 8547 'Sem parametros
'O produto em questão não pode ser excluido pois está relacionado a pelo menos uma nota fiscal baixada.
Public Const ERRO_PRODUTO_SUBSTITUTO_GERFINAL = 8548 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois é usado como Produto Substituto de outro Produto.
Public Const ERRO_PRODUTO_MOVIMENTOESTOQUE_GERFINAL = 8549 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois faz parte de pelo menos um Movimento de Estoque.
Public Const ERRO_LEITURA_ITENSORDEMPRODUCAO3 = 8550 'Sem Parâmetros.
'Ocorreu um erro na leitura da tabela de Itens de Ordem de Produção.
Public Const ERRO_PRODUTO_ITENSORDEMPRODUCAOBAIXADAS = 8551 'Parametro sCodProduto
'O Produto %s não pode ser excluído pois está sendo utilizado por pelo menos um Item de Ordem de Produção Baixado.
Public Const ERRO_PRODUTO_ITENSORDEMPRODUCAOGERFINAL = 8552 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois está sendo utilizado por pelo menos um Item de Ordem de Produção.
Public Const ERRO_PRODUTO_ITENSOPBAIXADASGERFINAL = 8553 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois está sendo utilizado por pelo menos um Item de Ordem de Produção Baixada.
Public Const ERRO_PRODUTO_EMPENHOGERFINAL = 8554 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois está sendo utilizado em pelo menos um Empenho.
Public Const ERRO_PRODUTO_KIT_GERFINAL = 8555 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois ele faz parte de pelo menos um Kit.
Public Const ERRO_PRODUTO_ITEMPVGERFINAL = 8556 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois ele participa de pelo menos um Item de Pedido de Venda.
Public Const ERRO_PRODUTO_ITEMPVBAIXADOGERFINAL = 8557 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois ele participa de pelo menos um Item de Pedido de Venda Baixado.
Public Const ERRO_PRODUTO_ITENSSOLCOMPRAGERFINAL = 8558 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois ele participa de pelo menos um Item de Solicitação de Compra.
Public Const ERRO_PRODUTO_INVENTARIOGERFINAL = 8559 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois ele participa de pelo menos um Inventário.
Public Const ERRO_PRODUTO_INVENTARIOPENDGERFINAL = 8560 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois ele participa de pelo menos um Inventário Pendente.
Public Const ERRO_PRODUTO_RESERVAGERFINAL = 8561 'Parametro sCodProduto
'O Produto %s não pode ser alterado de final para gerencial pois ele participa de pelo menos uma Reserva.
Public Const ERRO_PRODUTO_ITENSNFISCALGERFINAL = 8562 'Parametro sCodProduto
'O produto %s não pode ser alterado de final para gerencial pois está relacionado com pelo menos um item de nota fiscal.
Public Const ERRO_PRODUTO_ITENSNFISCALBAIXADASGERFINAL = 8563 'Parametro sCodProduto
'O produto %s não pode ser alterado de final para gerencial pois está relacionado a pelo menos um item de nota fiscal baixada.
Public Const ERRO_PRODUTO_LANPENDENTEGERFINAL = 8564 'Parametro sCodProduto
'O produto %s não pode ser alterado pois está relacionado com pelo menos um Lançamento Pendente.
Public Const ERRO_PRODUTO_CONTROLE_NAO_ATUALIZA = 8565 'Parâmetro: sCodProduto
'O Controle de Estoque do Produto %s não pode ser alterado para não inventariado, pois o produto possui pelo menos uma associação com um almoxarifado.
Public Const ERRO_ESTOQUEMES_ABERTO_INEXISTENTE = 8566 'Parametros iFilialEmpresa
'Não há mês em aberto no estoque para a filial %i.
Public Const ERRO_ESTOQUEMES_INEXISTENTE1 = 8567 'Parametros iFilialEmpresa
'Não há registro em EstoqueMes para a filial %i.
Public Const ERRO_APROPRIACAO_CUSTO_INEXISTENTE = 8571 'Parametro: sProduto
'Erro apropriação inexistente para o produto %s.
Public Const ERRO_LOTEINVPEN_NAO_CADASTRADO1 = 8572 'Parametro iLote
'O lote %i não está cadastrado.
Public Const ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1 = 8573 'Sem parametro
'O Item da Categoria não foi informado.
Public Const ERRO_NATUREZA_PRODUTO_NAO_PREENCHIDA = 8574 'Sem Parametros
'Erro é obrigatório o preenchimento da Natureza do Produto.
Public Const ERRO_CLASSIFICACAOFISCAL_NAOPREENCHIDA = 8575 'Sem Parametros
'Erro é obrigatório o preenchimento da Classificação Fiscal.
Public Const ERRO_EXCLUSAO_ESTOQUEPRODUTO_RESERVA = 8576 'Parametros CodProduto, iAlmoxarifado
'Não é possível excluir a associação entre o Produto %s e o Almoxarifado %i pois existe reserva deste produto neste almoxarifado.
Public Const ERRO_EXCLUSAO_ESTOQUEPRODUTO_EMPENHO = 8577 'Parametros CodProduto, iAlmoxarifado
'Não é possível excluir a associação entre o Produto %s e o Almoxarifado %i pois existe empenho deste produto neste almoxarifado.
Public Const ERRO_ESTOQUE_PRODUTO_EMPENHO_NAO_CAD = 8578 'Parâmetros: sProduto, iAlmoxarifado
'A associação entre o produto %s e o Almoxarifado %i não está cadastrada. Assim sendo, não é possivel gravar a quantidade empenhada do produto neste almoxarifado.
Public Const ERRO_ITEMOP_SITUACAO_BAIXADA = 8579 'sCodProduto, sCodigoOP
'Não é possível realizar esta operação pois o produto %s da ordem de produção %s possui situação "baixada".
Public Const ERRO_ITEMOP_SITUACAO_DESABILITADA = 8580 'sCodProduto, sCodigoOP
'Não é possível realizar esta operação pois o produto %s da ordem de produção %s possui situação "desabilitada".
Public Const ERRO_SALDO_MAT_OP = 8581 'Parametros sProduto, iAlmoxarifado, dSaldo
'O Saldo do Produto em Ordem de Produção é insuficiente para realizar esta operação. Produto = %s, Almoxarifado = %i, Saldo = %d.
Public Const ERRO_ORDEMPRODUCAO_BAIXADA = 8582 'sCodigoOP
'Não é possível realizar esta operação pois a ordem de produção %s está baixada.
Public Const ERRO_PRODUTO_NAO_PRODUZIVEL1 = 8583 'Parametro sCodProduto
'O Produto %s não pode ser produzido.
Public Const ERRO_NAO_KIT_BASICO = 8584 'Parametro sCodigo
'O produto %s não pode ser um produto básico de um Kit.
Public Const ERRO_OPCODIGO_NAO_CADASTRADO = 8585 'Parametro: sCodigoOP
'A ordem de producao %s não está cadastrada.
Public Const ERRO_OP_NAO_PREENCHIDO1 = 8586 'Parametro iLinhaGrid
'A Ordem de Produção do ítem %i do Grid de Material Produzido não foi preenchido.
Public Const ERRO_PRODUTOOP_NAO_PREENCHIDO_OP = 8587 'Sem Parametros
'O Produto O.P. só pode ser preenchido depois de ser preenchido a Ordem de Produção
Public Const ERRO_UM_NAO_PREENCHIDA = 8588 ' Parametro iLinhaGrid
'A Unidade de Medida do ítem %i do Grid não foi preenchida.
Public Const ERRO_LEITURA_ORDEMPRODUCAOBAIXADA = 8589 'Sem Parametros
'Ocorreu um erro na leitura da Tabela de Ordens de Produção Baixadas.
Public Const ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID = 8591 'iLinhaGrid
'A linha %i do grid já possui o campo produto preenchido.
Public Const ERRO_MOVIMENTO_NAO_PRODUCAO = 8592 'Paramtro lCodigoMov
'Os Movimentos de Estoque com código = %l não incluem entradas de material produzido.
Public Const ERRO_TRANSFERENCIA_MESMO_ESCANINHO = 8593 'Parametro iLinhaGrid
'A linha %i do grid indica uma transferencia envolvendo o mesmo almoxarifado e o mesmo tipo. (TipoOrigem = TipoDestino).
Public Const ERRO_TIPO_ESTOQUE_INVALIDO = 8594 'sTipoEstoque
'Tipo do Estoque inválido. Tipo = %s.
Public Const ERRO_ITEM_INVENTARIO_REPETIDO = 8595 'Parametros iLinhaGrid, iLinhaGrid1
'Os itens %i e %i do grid tem informação sobre o mesmo produto, almoxarifado e tipo.
Public Const ERRO_INVENTARIO_NAO_CADASTRADO = 8596 'Parametro sCodigoInv
'O Inventário com código %s não está cadastrado.
Public Const ERRO_PRODUTO_NAO_COMPONENTE = 8597 'sCodProdutoComponente, sCodProdutoItemOP
'O Produto %s não é componente do produto %s.
Public Const ERRO_MOVIMENTO_NAO_REQPRODUCAO = 8598 'Paramtro lCodigoMov
'Os Movimentos de Estoque com código = %l não incluem requisições de produção.
Public Const ERRO_RESERVA_CODIGO_DIFERENTE = 8599 'lCodReserva
'Existe uma reserva para o item do pedido em questão com código diferente. Código da Reserva Cadastrada = %l.
Public Const ERRO_EXCLUSAO_RECEB_BAIXADO_FORN = 8600 'Parametros lFornecedor, iFilialForn, dtDataEntrada, sSerie, lNumNotaFiscal
'Não é possível excluir o Recebimento de Material com os dados: Fornecedor = %l, Filial Fornecedor = %i, Data de Entrada = %dt, Série = %s, Nota Fiscal = %l. Este recebimento encontra-se com o status = BAIXADO.
Public Const ERRO_EXCLUSAO_RECEB_BAIXADO_CLI = 8601 'Parametros lCliente, iFilialCli, dtDataEntrada, sSerie, lNumNotaFiscal
'Não é possível excluir o Recebimento de Material com os dados: Cliente = %l, Filial Cliente = %i, Data de Entrada = %dt, Série = %s, Nota Fiscal = %l. Este recebimento encontra-se com o status = BAIXADO.
Public Const ERRO_DATA_VIGENCIA_MENOR_DATA_ATUAL = 8602 'Parametros: dtDataAtual
'A Data de Vigência não pode ser menor que a Data Atual que é: %dt.
Public Const ERRO_DATA_VIGENCIA_NAO_PREENCHIDA = 8603 'Sem Parametros
'É obrigatório o preenchimento da Data de Vigência.
Public Const ERRO_TABELAPRECOITEM_INEXISTENTE2 = 8604 'Parametros: iCodTabela, sCodProduto, dtDataVigencia
'O Item de Tabela de Preço com código da Tabela %i, código do Produto %s, Data de Vigência %dt não está cadastrada no Banco de Dados.
Public Const ERRO_QUANTIDADE_BENEF_NAO_PREENCHIDA = 8605 'Parâmetro: iLinhaGrid
'O campo Quantidade da Linha %i do Grid de Beneficiamento não foi preenchido.
Public Const ERRO_ALMOXARIFADO_BENEF_NAO_PREENCHIDO = 8606 'Parâmetro: iLinhaGrid
'O campo Almoxarifado da Linha %i do Grid de Beneficiamento não foi preenchido.
Public Const ERRO_UM_BENEF_NAO_PREENCHIDA = 8607 ' Parametro iLinhaGrid
'O campo Unidade de Medida da Linha %i do Grid de Beneficiamento não foi preenchido.
Public Const ERRO_NFISCAL_EXTERNA = 8608 'Sem parâmetros
'Não é possível gerar um Número para uma Nota Fiscal Externa.
Public Const ERRO_ALMOX_ITEMOP_PRODUCAO = 8609 'sCodProduto, iAlmoxProducao, iAlmoxItemOP
'A produção do produto %s não pode ser gravado pois o almoxarifado que vai estocar o produto = %i não coincide com o almoxarifado da ordem de produção = %i.
Public Const ERRO_DATA_INVENTARIO_MENOR = 8610 'sCodProduto, sAlmoxNomeRed, sDataUltimoInventario, sDataInventario
'A data do ultimo inventario do produto %s no almoxarifado %s é maior do que a data deste inventario. Data do último inventário = %s, Data deste Inventário = %s
Public Const ERRO_PRODUTO_SUBSTITUTO_GERENCIAL = 8611 'Sem Parametros
'Produto substituto não pode ser gerencial.
Public Const ERRO_NFISCAL_DIFERE_PARCELASPAGAR = 8612 'Parametros: dTotalNFMenosIR, dTotalParcelasPag
'O valor total da nota fiscal menos o imposto de renda difere do total de parcelas a pagar. Total da Nota - I.R. = %d, Total de Parcelas a Pagar = %d.
Public Const ERRO_QUANTIDADE_RESERVADA_MAIOR_FATURAR = 8613 'Parametros: dQuantReserva, dQuantFaturar
'A quantidade que está tentando resevar ultrapassa a quantidade que falta para faturar neste pedido. Quantidade a Reservar = %d, Quantidade a Faturar = %d.
Public Const ERRO_QUANTIDADE_RESERVA_NAO_PREENCHIDA = 8614 'Sem Parâmetro
'A quantidade da reserva não foi preenchida.
Public Const ERRO_QUANTIDADE_RESERVADA_NAO_POSITIVA = 8615 'Sem Parametros.
'A quantidade da reserva tem que ser um valor positivo.
Public Const ERRO_DOCUMENTO_ORIGEM_RESERVA = 8616 'Parametros lDocOrigemTela, lDocOrigemReserva
'O pedido de venda que originou esta reserva difere do que foi informado na tela. Documento de Origem Tela = %l, Documento de Origem Reserva = %l.
Public Const ERRO_RESERVA_NAO_CADASTRADA1 = 8617 'Parametros iFilialEmpresa, lCodigo
'A reserva em questão não está cadastrada. FilialEmpresa = %i, Código = %l.
Public Const ERRO_GRID_CATEGORIA_NAO_PREENCHIDA = 8618 'Sem Parametros
'A coluna de valor deve ser preenchida após o preenchimento da coluna categoria.
Public Const ERRO_ALMOXARIFADO_DE_OUTRA_FILIAL = 8619 'Parametros iCodAlmox, iFilialEmpresa
'O Almoxarifado %i pertence a filial %i e só pode ser alterado pela mesma.
Public Const ERRO_LEITURA_PRODUTOSFILIAL2 = 8620 'Sem parâmetros
'Erro na leitura da tabela ProdutosFilial.
Public Const ERRO_LEITURA_SLDMESESTALM1 = 8621 'Parametros iAno, iAlmoxarifado
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm). Ano = %i, Almoxarifado=%i.
Public Const ERRO_EXCLUSAO_SLDMESEST1 = 8622 'Parametros sProduto, iFilialEmpresa
'Erro na exclusão de registro da tabela de saldos mensais de estoque (SldMesEst). Codigo do Produto = %s, Filial = %i.
Public Const ERRO_ABERTURA_NOVOANO_SDLMESESTALM = 8623 'Parametros: iAno, sProduto, iAlmoxarifado
'Não foi possivel abrir o Ano %i para o produto %s no Almoxarifado %i
Public Const ERRO_ABERTURA_NOVOMES_SLDMESESTALM = 8624 'Parametros: iAno, iMes, iAlmoxarifado, sProduto
'Não foi possivel abrir um novo mês com os dados a seguir. Ano = i%, Mes: %i, Almoxarifado = %i, Produto = %s.
Public Const ERRO_ALTERACAO_SLDMESESTALM = 8625 'Parametros: iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na alteração do registro da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm). Ano = %i, Almoxarifado = %i, Produto = %s.
Public Const ERRO_ALTERACAO_SLDMESEST = 8626 'Parametros: iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na alteração do registro da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm). Ano = %i, Filial = %i, Produto = %s.
Public Const ERRO_LEITURA_TIPOSDEPRODUTO1 = 8627 'Sem parâmetros
'Erro na leitura da tabela TiposDeProduto.
Public Const ERRO_CLASSEUM_UTILIZADA_TIPOSDEPRODUTO = 8628 'Parâmetros: iClasse
'A Classe %i está sendo utilizada em Tipos de Produto.
Public Const ERRO_CLASSEUM_E_SIGLAUM_UTILIZADAS_TIPOSDEPRODUTO = 8629 'Parâmetros: iClasse, sSiglaUM
'A Classe %i e a Sigla %s estão sendo utilizadas em Tipos de Produto.
Public Const ERRO_TABELAPRECO_PADRAO = 8630
'Para desmarcar um Almoxarifado Padrão primeiro marque outro Almoxarifado como Padrão.
Public Const ERRO_LIMITE_ALMOX_VLIGHT = 8631 'Parametros : iNumeroMaxAlmoxarifados
'Número máximo de Almoxarifados desta versão é %i.
Public Const ERRO_LEITURA_ALMOXARIFADO2 = 8632 'Sem Parametro
'Erro na leitura da Tabela de Almoxarifados.
Public Const ERRO_LIMITE_CLASSE_UM = 8633 'Parametros : iNumeroMaxClasseUM
'Número máximo de Classes de Unidade de Medida desta Versão é %i.
Public Const ERRO_LIMITE_FORNPRODUTO = 8634 'Parametros: iNumeroMaxFornProduto
'Número máximo de fornecedores por Porduto desta versão é %i.
Public Const ERRO_LIBERAR_QUANTIDADE_RESERVADA = 8635 'Parametros iLinha, dQuantidade
'Na linha %i, a quantidade %d está reservada, é necessário liberar a reserva.
Public Const ERRO_LIBERAR_QUANTIDADE_RESERVADA_CONSIG = 8636 'Parametros: iLinha, dQuantidade
'Na linha %i, a quantidade %d está reservada em consignação, é necessário liberar a reserva.
Public Const ERRO_SELECAO_ALMOXARIFADO = 8637 'Sem parâmetros
'A seleção de "Todos Almoxarifados" deve ser usada apenas como consulta.
Public Const ERRO_NUMAUTO_FORNECEDOR = 8638 'Sem parâmetros
'Não é possível gerar um número automático para a Nota Fiscal de um Fornecedor.
Public Const ERRO_RECEBIMENTO_BAIXADO = 8639 'Parametros: lNumRecebimento
'Recebimento com o Número %l está baixado, não é possível alterar.
Public Const ERRO_RECEBIMENTO_NAO_CADASTRADO = 8640 'Parametros: lNumRecebimento
'Recebimento com o Número %l não está cadastrado, não é possível alterar.
Public Const ERRO_RECEBIMENTO_NAO_PREENCHIDO = 8641 'Sem Parametros
'É obrigatório o preenchimento do Número do Recebimento.
Public Const ERRO_RECEBIMENTO_NFEXT_CADASTRADO_FORNECEDOR = 8642 'Parmetros: lNumRecebimento
'Não é possivel a Inserção. O Recebimento Número %l já possui os Dados: Série= %s, Número = %l, Fornecedor = %l, Filial = %i, Data de Entrada = %dt.
Public Const ERRO_NUMAUTO_CLIENTE = 8643 'Sem parâmetros
'Não é possível gerar um número automático para a Nota Fiscal de um Cliente.
Public Const ERRO_RECEBIMENTO_NFEXT_CADASTRADO_CLIENTE = 8644 'Parametro: lNumRecebimento
'Não é possivel a Inserção. O Recebimento Número %l já possui os Dados: Série= %s, Número = %l, Cliente = %l, Filial = %i, Data de Entrada = %dt.
Public Const ERRO_EXCLUSAO_ESTOQUEPRODUTO1 = 8645 'Sem Parametros
'Ocorreu um erro na tentativa de excluir um registro da tabela de Estoque de Produtos.
Public Const ERRO_LEITURA_MOVIMENTOESTOQUE4 = 8646 'Parametros sCcl
'Ocorreu um erro na leitura de um registro da tabela de movimentos de estoque. Centro de Custo=%s
Public Const ERRO_RECEBIMENTO_NFINT_NAO_PREENCHIDO = 8647 'Sem Parametros
'É obrigatório o preenchimento do Número do Recebimento para Notas Fiscais Internas.
Public Const ERRO_TIPODOCINFO_NAO_CADASTRADO2 = 8648 'Sem Parametros
'O TipoDocInfo com o Código %i não está cadastrado.
Public Const ERRO_RECEBIMENTO_CLIENTE = 8649 'Parametros: lNumRecebimento
'O Recebimento %l é um Recebimento de Cliente
Public Const ERRO_PRODUTO_NOTA_SERVICO = 8650 'Parâmetro: sProduto
'O produto %s não pode ser utilizado. Em uma nota de serviço somente os produto produzidos, vendáveis e não estocáveis podem ser utilizados.
Public Const ERRO_TIPONFISCALORIGINAL_DIFERENTE_TELA = 8651 'lNumNotaFiscal, sSerieNFiscal, lForn/Cli, iFilial, sSiglaNF, sSiglaNFOrig
'Nota Fiscal Número %l, Série %s Forn/Cli %l, Filial %i é uma Nota do tipo %s  e tem devolução do tipo %s.
Public Const ERRO_SERIE_NAO_CONFIGURADA = 8652 'Parametro: sSerie
'É preciso configurar uma Série como Padrão.
Public Const ERRO_QUANTINV_ESTOQUEATUAL = 8653 'Parametro sCodInventario, sCodProduto, iAlmoxarifado, sSiglaUM, dQuantInvEst, dQuantEstoque
'No lançamento do inventário %s referente ao produto %s e almoxarifado %i foi registrada uma quantidade do produto estocada diferente da quantidade atual. Unidade de Medida = %s, Quantidade Encontrada pelo Inventário = %d, Quantidade Atual no Estoque = %d.
Public Const ERRO_LEITURA_INVLOTEPENDENTE1 = 8654 'Sem Parametros
'Erro de Leitura na Tabela InvLotePendente.
Public Const ERRO_IMPOSSIVEL_ATUALIZACAO = 8655 'Parametros iLote , iFilialEmpresa
'Impossível atualizar o Lote %i com Filial %i , que já está sendo atualizado .
Public Const ERRO_UNLOCK_INVLOTEPENDENTE = 8656 'Parametros iLote, iFilialEmpresa
'Erro na tentativa de fazer "unlock" na tabela InvLotePendente para o Lote %i com Filial %i.
Public Const ERRO_INSERCAO_INVLOTE = 8657 'Parametros iLote, iFilialEmpresa
'Erro na tentativa de inserção na tabela InvLote para o Lote %i com Filial %i.
Public Const ERRO_ATUALIZACAO_ESTOQUEPRODUTO1 = 8658  'Sem Parametros
'Ocorreu um erro na tentativa de atualizar a tabela de EstoqueProduto.
Public Const ERRO_UNLOCK_ESTOQUEPRODUTO = 8659  'Sem Parametros
'Ocorreu um erro na tentativa de fazer 'unlock' na tabela de EstoqueProduto.
Public Const ERRO_UNLOCK_MATCONFIG = 8660  'Parametro sCodigo
'Ocorreu um erro na tentativa de fazer 'unlock' na tabela Configuração de Materiais (MATConfig). Codigo=%s.
Public Const ERRO_ESTOQUEMES_CMP_APURADO = 8661 'Parametros iFilialEmpresa, iAno, iMes
'O Custo Médio de Prudução de EstoqueMes já foi apurado. FilialEmpresa = %i, Ano = %i, Mês = %i.
Public Const ERRO_ATUALIZACAO_MOVESTOQUE = 8662 'Parametro: lNumIntDoc
'Erro de atualização na tabela MovimentoEstoque. Número interno = %l.
Public Const ERRO_AUSENCIA_REGISTRO_SALDOMESEST = 8663 'Parametros: iFilialEmpresa, iAno, sProduto
'Ausência na tabela SaldoMesEst de registro com chave, FilialEmpresa = %i, Ano = %i, Produto = %s.
Public Const ERRO_AUSENCIA_MOVTOS_PRODUTOS_PRODUZIDOS = 8664 'Parametros: iFilialEmpresa, iMes, iAno
'Ausência de movimentos de estoque de produtos produzidos. FilialEmpresa = %i, Ano = %i, Mês = %i.
Public Const ERRO_CUSTO_PRODUCAO_NAO_INFORMADO = 8665 'Parametros: sProduto, iMes, iAno
'Custo de Produção do Produto %s no mês %i no ano %i não foi cadastrado.
Public Const ERRO_FALTA_ITEMMOVEST_MOVTOESTOQUE = 8666 'Parametro: lNumIntDoc
'Falta Item de Movimento de Estoque com número interno %l na tabela MovimentoEstoque do Banco de Dados.
Public Const ERRO_MOVESTOQUE_SAIDA_APROPR_CRP = 8667 'Parametro: lNumIntDoc
'Movimento de Estoque com número interno %l é uma saída com apropriação Custo Real de Produção.
Public Const ERRO_LEITURA_SLDDIAEST2 = 8668 'Parametros: iFilialEmpresa
'Erro de leitura na tabela SaldoDiaEst. FilialEmpresa = %i.
Public Const ERRO_AUSENCIA_REGISTRO_SLDDIAEST = 8669 'Parametros: iFilialEmpresa, sProduto, dtData
'Ausência de registro na tabela SaldoDiaEst. Chave: FilialEmpresa = %i, Produto = %s, Data = %dt.
Public Const ERRO_ATUALIZACAO_ESTOQUEMES = 8670 'Parametros: iFilialEmpresa, iAno, iMes
'Erro de atualização na tabela EstoqueMes com chave FilialEmpresa=%i, Ano=%i, Mês=%i.
Public Const ERRO_LEITURA_SLDMESESTALM2 = 8671 'Parametro iAno
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm). Ano = %i.
Public Const ERRO_AUSENCIA_REGISTRO_SLDMESESTALM = 8672 'Parametros: iAno, sProduto, iAlmoxarifado
'Ausência na tabela SaldoMesEst do registro com chave, Ano = %i, Produto = %s, Almoxarifado = %i.
Public Const ERRO_LEITURA_SLDDIAESTALM2 = 8673 'Sem parametros
'Ocorreu erro na leitura da tabela de saldos diários de estoque por almoxarifado.
Public Const ERRO_AUSENCIA_REGISTRO_SLDDIAESTALM = 8674 'Parametros: iAlmoxarifado, sProduto, dtData
'Ausência de registro na tabela SaldoDiaEstAlm. Almoxarifado = %i, Produto = %s, Data = %dt.
Public Const ERRO_CUSTO_STANDARD_NAO_INFORMADO = 8675 'Parametros: sProduto, iMes, iAno
'O custo standard do produto %s para o mês %i/%i não foi informado.
Public Const ERRO_ATUALIZACAO_SLDMESEST1 = 8676  'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i, Produto=%s.
Public Const ERRO_ATUALIZACAO_SLDMESEST2 = 8677  'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i, Produto=%s.
Public Const ERRO_LEITURA_SLDMESESTALM11 = 8678 'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm1). Ano = %i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_LOCK_SLDMESESTALM1 = 8679 'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm1). Ano = %i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_ATUALIZACAO_SLDMESESTALM1 = 8680  'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm1). Ano=%i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_LEITURA_SLDMESESTALM21 = 8681 'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm2). Ano = %i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_LOCK_SLDMESESTALM2 = 8682 'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm2). Ano = %i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_ATUALIZACAO_SLDMESESTALM2 = 8683  'Parametros iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm2). Ano=%i, Almoxarifado=%i, Produto=%s.
Public Const ERRO_INSERCAO_SLDMESEST1 = 8684 'Parâmetros: iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de inserir um registro na tabela SldMesEst1 com Ano %i, Filial %i e Produto %s.
Public Const ERRO_INSERCAO_SLDMESEST2 = 8685 'Parâmetros: iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de inserir um registro na tabela SldMesEst2 com Ano %i, Filial %i e Produto %s.
Public Const ERRO_INSERCAO_SLDMESESTALM1 = 8686 'Parâmetros: iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de inserir um registro na tabela SldMesEstAlm1 com Ano = %i, Almoxarifado = %i, e Produto = %s.
Public Const ERRO_INSERCAO_SLDMESESTALM2 = 8687 'Parâmetros: iAno, iAlmoxarifado, sProduto
'Ocorreu um erro na tentativa de inserir um registro na tabela SldMesEstAlm2 com Ano = %i, Almoxarifado = %i, e Produto = %s.
Public Const ERRO_EXCLUSAO_ORDENSPRODUCAOBAIXADAS = 8688 'sem parametros
'Erro na tentativa de exclusão na tabela de Ordens de Produção Baixadas.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_DISP = 8689 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material disponível difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_CONSERTO = 8690 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em conserto difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_CONSIG = 8691 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em consignação difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_DEMO = 8692 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em demonstração difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_OUTROS = 8693 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material outros difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_BENEF = 8694 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em beneficiamento difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_CONSERTO3 = 8695 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em conserto de terceiros difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_CONSIG3 = 8696 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em consignação de terceiros difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_DEMO3 = 8697 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em demonstração de terceiros difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_OUTROS3 = 8698 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material outros de terceiros difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_QUANT_RASTO_DIF_INICIAL_BENEF3 = 8699 'dQuantInicial, dQuantTotalRastro
'A quantidade inicial para material em beneficiamento de terceiros difere da quantidade rastreada. Quant. Inicial = %d, Quant. Rastreada = %d.
Public Const ERRO_TIPO_NFISCAL_NAO_RASTRO = 8700 'Parametro sTipoNFiscal
'O tipo de Nota Fiscal %s não pode ter rastro associado.
Public Const ERRO_TIPO_NFISCAL_NAO_ENTRADA = 8701 'Parametro iTipoNFiscal
'O tipo de Nota Fiscal %i não é de entrada.
Public Const ERRO_ITEMNF_NAO_SELECIONADO = 8702 'Sem Parametros
'Nenhum item de nota fiscal foi selecionado.
Public Const ERRO_ALMOX_NAO_SELECIONADO = 8703 'Sem Parametros
'Nenhum almoxarifado foi selecionado.
Public Const ERRO_TIPO_NFISCAL_NAO_SAIDA = 8704 'Parametro iTipoNFiscal
'O tipo da Nota Fiscal %i não é de saída.
Public Const ERRO_INSERCAO_MATCONFIG = 8705  'Parametro sCodigo, iFilial
'Ocorreu um erro na tentativa de fazer 'lock' na tabela Configuração de Materiais (MATConfig). Codigo=%s. Filial = %i.
Public Const ERRO_ATUALIZACAO_MATCONFIG1 = 8706  'Parametro sCodigo, iFilial
'Ocorreu um erro na tentativa de atualizar um registro na tabela de Configuração de Materiais (MATConfig). Codigo=%s, Filial = %i.
Public Const ERRO_MOVIMENTOESTOQUE_REPROCESSAMENTO = 8707  'Parametros: iFilialEmpresa, lCodigo, lNumIntDoc, sProduto, sSiglaUM, dQuantidade, iAlmoxarifado, iTipoMov, dtData
'O Erro ocorreu no movimento de estoque com os seguintes parametros: FilialEmpresa=%i, Codigo=%l, NumIntDoc = %l, Produto = %s, SiglaUM = %s, Quantidade = %d, Almoxarifado = %i, Tipo de Movimento = %i, Data = $dt.
Public Const ERRO_ATUALIZACAO_SLDMESEST1_1 = 8708  'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i, Produto=%s.
Public Const ERRO_ATUALIZACAO_SLDMESEST2_1 = 8709  'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i, Produto=%s.
Public Const ERRO_LEITURA_SLDMESESTALM1_1 = 8710 'Parametros iAno, iAlmoxarifado
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm1). Ano = %i, Almoxarifado=%i.
Public Const ERRO_LEITURA_SLDMESESTALM2_1 = 8711 'Parametros iAno, iAlmoxarifado
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque por almoxarifado (SldMesEstAlm2). Ano = %i, Almoxarifado=%i.
Public Const ERRO_LOCK_MATCONFIG1 = 8712  'Parametro sCodigo, iFilialEmpresa
'Ocorreu um erro na tentativa de fazer 'lock' na tabela Configuração de Materiais (MATConfig). Codigo=%s, Filial = %i.
Public Const ERRO_LEITURA_MOVIMENTOESTOQUE5 = 8713 'Parametros iFilialEmpresa, Produto, Data
'Ocorreu um erro na leitura de um registro da tabela de movimentos de estoque. FilialEmpresa = %i, Produto = %s, Data >= %dt.
Public Const ERRO_LEITURA_MATCONFIG1 = 8714 'Parametro sCodigo, iFilial
'Ocorreu um erro na leitura da tabela de Configuração de Materiais (MATConfig). Codigo=%s, Filial = %i.


'fernando
Public Const ERRO_TIPONUMINTDOCORIGEM_NAO_TRATADO = 0 'Parametros: iTipoNumIntDocOrigem
'O Tipo de Documento de Origem do movimento de estoque não é tratado pelo sistema. Tipo = %i.
Public Const ERRO_DATA_REPROC_MENOR_DATA_INICIO = 0 'Parametros: dtDataInicio, dtDataMinimaReprocessamento
'A Data de início do reprocessamento é maior que a data mínima exigida para esse procedimento. Data de início: %dt, Data mínima:  %dt.
Public Const ERRO_EXCLUSAO_MATCONFIG = 0  'Parametro sCodigo, iFilialEmpresa
'Ocorreu um erro na tentativa de excluir um registro da tabela de Configuração de Materiais (MATConfig). Codigo=%s, Filial = %i.



'VEIO DE ERROS TRB
Public Const ERRO_LEITURA_NATOPPADRAO = 7050 'sem parametros
'Erro na leitura da tabela de padrões de natureza de operação


'VEIO DE ERROS PV
Public Const ERRO_LEITURA_BLOQUEIOSPV = 7202 'lCodigo
'Erro na tentativa de leitura dos Bloqueios associados ao Pedido de Venda de Código = %l na Tabela BloqueiosPV.
Public Const ERRO_PEDIDO_VENDA_SEM_ITENS = 7203 'lCodPedido
'O Pedido de Venda com o Código %l não possui Ítens associados à ele.
Public Const ERRO_LEITURA_RESERVAITEMBD = 7208
'Erro na leitura de Reserva Item.


'VEIO DE ERROS FAT
Public Const ERRO_LEITURA_BLOQUEIOS_PV_LIBERACAO = 8003 'Sem parâmetro
'Erro na leitura de Bloqueio de Pedido de Venda.
Public Const ERRO_EXCLUSAO_COMISSOESPEDVENDAS = 8016 'Parametro: lNumIntDoc
'Erro na tentativa de excluir registro da tabela de Comissões de Pedido de Vendas, do pedido de Número %l
Public Const ERRO_LEITURA_PEDIDOSDEVENDA = 8018 'Parametros:lCodigo
'Erro na leitura da tabela de pedidos de venda, do pedido de número %l.
Public Const ERRO_PEDIDO_VENDA_NAO_CADASTRADO = 8019 'Parametros: lCodigo
'O pedido de venda de número %l não está cadastrado.
Public Const ERRO_LOCK_PEDIDOS_DE_VENDA = 8036 'Parametro: lCodigo
'Erro na tentativa de "lock" do Pedido de Venda %l na tabela PedidosDeVenda.
Public Const ERRO_ATUALIZACAO_PEDVENDA = 8040 'Parametro: lCodigo
'Erro na atualização do Pedido de Venda %l na tabela PedidosDeVenda.
Public Const ERRO_INSERCAO_PEDVENDA = 8041 'Parametro: lCodigo
'Erro na inserção do Pedido de Venda %l na tabela PedidosDeVenda.
Public Const ERRO_ATUALIZACAO_BLOQUEIOPV = 8043 'Parametros: iFilialEmpresa, lPedidoDeVendas, iSequencial
'Ocorreu um erro na atualização de um registro da tabela de Bloqueios de Pedido de Venda. Filial=%i, Pedido=%l, Sequencial=%i.
Public Const ERRO_INSERCAO_BLOQUEIOPV = 8044 'Parametros: iFilialEmpresa, lPedidoDeVendas, iSequencial
'Ocorreu um erro na inserção de um registro da tabela de Bloqueios de Pedido de Venda. Filial=%i, Pedido=%l, Sequencial=%i.
Public Const ERRO_CANALVENDA_NAO_CADASTRADO = 8049 'Parametro objCanal.iCodigo
'Canal %i não cadastrado
Public Const ERRO_LEITURA_PREVVENDA = 8070 'Parâmetro sCodigo
'Erro na leitura da Previsão de Venda com código %s na tabela PrevVenda .
Public Const ERRO_PRODUTO_NAO_CADASTRADO = 8076 'Parâmetro sProduto
'Produto %s não cadastrado na tabela Produtos .
Public Const ERRO_INSERCAO_PEDIDODEVENDABAIXADO = 8080 'Parametro lCodPedido
'Ocorreu um erro na tentativa de inserir um registro na tabela de Pedidos de Venda Baixados. Pedido = %l.
Public Const ERRO_EXCLUSAO_PEDIDODEVENDA = 8081 'Parametros = lCodPedido
'Ocorreu um erro na tentativa de excluir um registro da tabela de Pedido de Venda. Pedido = %l.
Public Const ERRO_EXCLUSAO_BLOQUEIOSPV = 8092 'lCodigo
'Erro na tentativa de exclusão dos Bloqueios associados ao Pedido de Venda de Código = %l na Tabela BloqueiosPV.
Public Const ERRO_LEITURA_PEDIDOSDEVENDABAIXADOS = 8111 'lCodigo
'Erro na leitura do Pedido de Venda código = %l da tabela de Pedidos de Venda Baixados



'VEIO DE ERROSCOM
Public Const ERRO_LEITURA_CONCORRENCIA = 12024 'Sem Parâmetros
'Erro de leitura na tabela de concorrências.
Public Const ERRO_USUARIO_NAO_ENCONTRADO = 12061
'Nao existe usuario
Public Const ERRO_LEITURA_ITEMPEDCOTACAO = 12101 'Parametro: lCodPedidoCotacao
'Erro na leitura dos Itens do Pedido de Cotação %l.
Public Const ERRO_LEITURA_ITENSCOTACAO = 12102 'Sem parametros
'Erro na leitura da tabela ItensCotacao.
Public Const ERRO_LEITURA_PEDIDOCOTACAO = 12103 'Parametro: lCódigo
'Erro na leitura do Pedido de Cotação com o código %l.
Public Const ERRO_LEITURA_COTACAO = 12104 'Sem parametros
'Erro na leitura da tabela Cotacao.
Public Const ERRO_LEITURA_ITENSPEDCOMPRA = 12105 'Sem Parametros
'Erro na leitura da tabela ItensPedCompra.
Public Const ERRO_LEITURA_COTACAOITEMCONCORRENCIA = 12106 'Sem parametros
'Erro na leitura da tabela CotacaoItemConcorrencia.
Public Const ERRO_LOCK_ITEMPEDCOTACAO = 12108 'Parametro: lCodPedidoCotacao
'Não conseguiu fazer o lock nos itens do pedido de cotação %l.
Public Const ERRO_LOCK_ITENSCOTACAO = 12109 'Parametro: lCodigo
'Não conseguiu fazer o lock dos itens de cotação do pedido de cotação %l.
Public Const ERRO_EXCLUSAO_ITEMPEDCOTACAO = 12116 'Parametro lCodigo
'Erro na exclusão dos itens do pedido de cotação %l.
Public Const ERRO_EXCLUSAO_ITENSCOTACAO = 12117 'Parametro lCodigo
'Erro na exclusão dos itens de cotação do pedido de cotação %l.
Public Const ERRO_USUARIO_NAO_CADASTRADO2 = 12174 'Parametros: sNomeReduzido
'O Usuario %s não está cadastrado.



'VEIO de ERROSFIS
Public Const ERRO_ATUALIZACAO_LIVREGCADPROD = 13353 'Sem Parametros
'Erro na tentativa de Atualizar LivRegESCadProd.
Public Const ERRO_INSERCAO_LIVREGESCADPROD = 13424
'Erro na Inserção na Tabela de LivRegESCadProd.
Public Const ERRO_LEITURA_LIVREGESCADPROD = 13425
'Erro na leitura da tabela de LivRegESCadProd.


'Códigos de Avisos - Reservado de 5800 até 5899
Public Const AVISO_EXCLUIR_CATEGORIAPRODUTO = 5800 'Parametro: sCategoria
'Confirma exclusão da Categoria %s?
Public Const AVISO_CRIAR_ALMOXARIFADO1 = 5801 'Parametro: iCodigo
'Almoxarifado com código %i não está cadastrado no Banco de Dados. Deseja criar?
Public Const AVISO_CRIAR_ALMOXARIFADO2 = 5802 'Parametro: sCodigo
'Almoxarifado %s não está cadastrado no Banco de Dados. Deseja criar?
Public Const AVISO_EXCLUIR_CLASSEUM = 5803 'Parametro: iClasse
'Confirma exclusão da Classe %i?
Public Const AVISO_CRIAR_CLASSEUM = 5804 'Parâmetro: iClasse
'A Classe %i de Unidade de Medida não existe. Deseja Criá-la?
Public Const AVISO_EXCLUSAO_TIPOPRODUTO = 5805 'Parametro: iTipo
'Confirma exclusão do Tipo de Produto %i?
Public Const AVISO_CRIAR_CATEGORIAPRODUTO = 5806 'sCategoriaProduto
'A Categoria %s não está cadastrada. Deseja criar uma nova Categoria de Produto?
Public Const AVISO_CRIAR_CATEGORIAPRODUTOITEM = 5807 'sItem
'O ítem %s não está cadastrado. Deseja criar um novo Item de Categoria de Produto?
Public Const AVISO_SELECIONAR_ESTRUTURA_PRODUTO = 5808 'Sem parametros
'Um item da árvore de estrutura do produto deve ser selecionado.
Public Const AVISO_TROCAR_PRODUTO = 5809 'Sem parametros
'Para excluir o produto, selecione-o na lista de Estrutura do Produto. O produto que está selecionado é o Produto Raiz e este não pode ser excluído.
Public Const AVISO_PRODUTO_TEM_FILHOS = 5810 'Sem parametros
'O produto selecionado tem filhos. Deseja excluir assim mesmo?
Public Const AVISO_ALMOXARIFADO_INEXISTENTE = 5811 'Parametro sAlmoxarifadoNomeRed
'Almoxarifado com Nome Reduzido %s não está cadastrado no Banco de Dados. Deseja criar novo Almoxarifado?
Public Const AVISO_ALMOXARIFADO_INEXISTENTE1 = 5812 'Parametro iCodAlmoxarifado
'Almoxarifado com Código %i não está cadastrado no Banco de Dados. Deseja criar novo Almoxarifado?
Public Const AVISO_OP_NAO_CADASTRADA = 5814 'Parametro sCodOrdemProducao
'Ordem de Produção %s não está cadastrada no Banco de Dados. Deseja Criar?
Public Const AVISO_OPCODIGO_NAO_CADASTRADO = 5815 ' Parametro sCodigoOP
'Ordem de Produção com Código %s não está cadastrado no Banco de Dados. Deseja Criar?
Public Const AVISO_CONFIRMA_EXCLUSAO_EMPENHO = 5816 'Sem Parâmetro
'Confirma a exclusão do Empenho ?
Public Const AVISO_EXCLUSAO_PRODUTO_FINAL = 5817
'O Produto que está sendo excluido é Final. Confirma a exclusão?
Public Const AVISO_EXCLUSAO_PRODUTO_GERENCIAL_COM_FILHOS = 5818 'Sem parametros
'O Produto que está sendo excluido é sintético e possui Produtos abaixo dele.
'Ao excluir este Produto, seus "filhos" serão também excluidos. Confirma a exclusão?
Public Const AVISO_EXCLUSAO_PRODUTO_GERENCIAL = 5819
'O Produto que está sendo excluido é Gerencial e não possui Produtos abaixo dele.
'Confirma a exclusão?
Public Const AVISO_EXISTENCIA_NOTA_FISCAL = 5820 'Parâmetro: lFornecedor, iFilialForn, sSerie, lNumNotaFiscal, dtDataEmissao
'Já existe uma Nota Fiscal com os Dados: Código do Fornecedor = %s, Código da Filial = %s, Série NF = %s, Número NF = %s, Data Emissão = %s.
'Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número ?
Public Const AVISO_EXISTENCIA_NOTA_FISCAL_BAIXADA = 5821 'Parâmetro: lFornecedor, iFilialForn, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Baixada com os Dados Código do Fornecedor =%l, Código da Filial =%i, Tipo =%i,  Série NF =%s, Número NF =%l, Data Emisão =%dt.
'Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número?
Public Const AVISO_CONFIRMA_EXCLUSAO_NFISCAL_ENTRADA = 5822 ' Parâmetros: iTipoNFiscal, lFornecedor, iFilialForn, sSerie, lNumNotaFiscal, dtDataEmissao
'Confirma a exclusão da Nota Fiscal Entrada do Banco de Dados com dados: Tipo =%i, Código do Fornecedor =%l, Filial Fornecedor = %l, Série =%s, Número Nota Fiscal =%l, Data Emissão =%dt?
Public Const AVISO_CRIAR_NATUREZA_OPERACAO = 5823 'Parâmetro: sNaturezaOp
'A Natureza Operacao com Código %s não está cadastrada no Banco de Dados. Deseja Criar?
Public Const AVISO_EXCLUIR_ALMOXARIFADO = 5824 'Parametro: iCodigo
'Confirma exclusão do Almoxarifado com código %i?
Public Const AVISO_EXCLUIR_CLASSEABC = 5825 'parametro: sCodigo
'Confirma exclusão da Classificação ABC com código %s ?
Public Const AVISO_CRIAR_PRODUTOFILIAL = 5826 'Parâmetro : iFilialEmpresa, sProduto
'Produto %s não está cadastrado na tabela ProdutoFilial com FilialEmpresa = %i. Deseja cadastrar ?
Public Const AVISO_PREENCHER_TELA = 5827 'Parametro lCodigo
'O código %l já existe . Deseja trazê-lo para a tela ?
Public Const AVISO_MOVIMENTO_ESTOQUEPRODUTO = 5828 'Parâmetros: sProduto, iAlmoxarifado
'Não é permitido atualizar a quantidade, valor e data de Estoque Inicial do Produto %s e Almoxarifado %i porque já houve movimento. Se o campo Localização Física e Conta Contabil foi modificado, só este será atualizado.
Public Const AVISO_EXCLUSAO_ESTOQUEINICIAL = 5829 'Parâmetros: sProduto, iAlmoxarifado
'Confirma exclusão de Estoque Inicial do Produto %s e Almoxarifado %i ?
Public Const AVISO_CANCELAR_FECHAMENTO_MES = 5830 'Parâmetro: iMes
'Confirma o cancelamento do fechamento do mes i% ?
Public Const AVISO_TERMINO_FECHAMENTO_MES = 5831 'Sem Parâmetros
'Termino do fechamento do mes
Public Const AVISO_EXCLUSAO_FORNECEDOR_PRODUTO = 5832 'Parâmetros: lFornecedor, sProduto
'Confirma exclusão de Fornecedor %l de Produto %s ?
Public Const AVISO_SAIR_SEM_SALVAR = 5833
'Deseja sair sem gerar a Ordem de Produção atual ?
Public Const AVISO_CONFIRMA_EXCLUSAO_KIT = 5834 'parametro sProdutoRaiz , sVersao
'Confirma a exclusão do Kit com ProdutoRaiz=%s e Versão=%s ?
Public Const AVISO_NF_EXTERNA_DATA_PROXIMA2 = 5835 'Parâmetros: lFornecedor, iFilialForn, lCliente, iFilialCli, sSerie,lNumNotaFiscal,dtDataEmissao
'No Banco de Dados existe Nota Fiscal Externa com os dados Código do Fornecedor = %l, Código da Filial Fornecedor = %i, Cliente = %l, Código da Filial Cliente = %i, Série = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na insercao de nova Nota Fiscal com o mesmo número?
Public Const AVISO_NF_INTERNA_DATA_PROXIMA2 = 5836 'Parâmetros: sSerie,lNumNotaFiscal,dtDataEmissao
'No Banco de Dados existe Nota Fiscal Interna com os dados Série = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número?
Public Const AVISO_NFISCAL_REMESSA_MESMO_NUMERO = 5837 'Parâmetros: lFornecedor, iFilialForn, lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados exite Nota Fiscal de Remessa com os Dados Código do Fornecedor =%l, Código da Filial Fornecedor =%i, Código do Cliente =%l AND Código da Filial Cliente =%i, Tipo =%i, Série NF =%s, Número NF =%l, Data Emissão =%dt.
'Deseja prosseguir na inserção de Nota Fiscal de Remessa com o mesmo número?
Public Const AVISO_NF_EXTERNA_DATA_PROXIMA = 5838 'Parâmetros:lFornecedor, iFilialForn, sSerie, lNumero, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Externa com os dados Código do Fornecedor = %l, Código da Filial = %i, Série = %s, Número = %l, Data Emissão = %dt.
'Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número ?
Public Const AVISO_QUANTIDADE_ITEMPEDIDO = 5839
'A quantidade ordenada é diferente da quantidade do item de pedido !
Public Const AVISO_CONFIRMA_EXCLUSAO_OP = 5840 ' Parâmetros: iCodigor, iFilialEmpresa
'Confirma a exclusão da Ordem de Produção do Banco de Dados com dados: Código da OP =%l, Filial Empresa = %i ?
Public Const AVISO_QUANTIDADE_ESTOQUEMAXIMO = 5841 'Parametro: objItemOP.sProduto
'A soma da quantidade ordenada mais a quantidade disponível é maior que a quantidade de estoque máximo do produto %s. Confirma a quantidade informada?
Public Const AVISO_QUANTIDADE_PEDIDA_A_MAIOR = 5842 'Parametros dQuantidadePedida ,sProduto,dQuantidadeFaltaProduzir,sCodigoOP
'A quantidade %d do Produto %s para gerar requisições é maior do que a quantidade que falta produzir %d na Ordem de Produção %s. Deseja prosseguir substituindo na geração a quantidade pedida pela que falta produzir?
Public Const AVISO_APAGAR_GRID = 5843
'Deseja limpar o Grid de Material Requisitado ?
Public Const AVISO_RECEBMATERIALC_MESMO_NUMERO = 5844 ' Parametros lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEntrada
'No Banco de Dados existe Recebimento de Material com os dados Código do Cliente %l, Código da Filial %i, Tipo %i, Série NF %s, Número NF %l, DataEntrada %dt. Deseja prosseguir na inserção de novo Recebimento de Material com o mesmo número de Nota Fiscal?
Public Const AVISO_CONFIRMA_EXCLUSAO_RECEBIMENTO1 = 5845 'Parametro: lCodigo
'Confirma a exclusão de Recebimento de Material de Cliente com Número %l ?
Public Const AVISO_RECEBMATERIALF_MESMO_NUMERO = 5846 ' Parametros lFornecedor, iFilialForn, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEntrada
'No Banco de Dados existe Recebimento de Material com os dados Código do Fornecedor %l, Código da Filial %i, Tipo %i, Série NF %s, Número NF %l, DataEntrada %dt. Deseja prosseguir na inserção de novo Recebimento de Material com o mesmo número de Nota Fiscal?
Public Const AVISO_CONFIRMA_EXCLUSAO_RECEBIMENTO = 5847 'Parametro: lCodigo
'Confirma a exclusão de Recebimento de Material de Fornecedor do Banco de Dados com Código %l ?
Public Const AVISO_LINHAS_NOVAS = 5848
'Existem novos movimentos junto com os movimentos de estorno. Deseja seguir gravando somente os movimentos de estorno?
Public Const AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO = 5849
'Deseja substituir o Acréscimo Financeiro pelo Acréscimo Financeiro da Condição de Pagamento?
Public Const AVISO_EXCLUSAO_RESERVA = 5850 'Parametro lCodReserva
'Confirma exclusão de Reserva %l do Banco de Dados ?
Public Const AVISO_ALTERAR_RESERVA = 5851 'Parametrso lCodReserva, lCodPedidoVenda, iNumItemPedido, sCodProduto,iCodAlmoxarifado
'Reserva com Código=%l , Pedido de Venda =%l , Item do Pedido= %i , Produto= %s , Almoxarifado = %i , está cadastrada no Banco de Dados .Deseja adicionar quantidade Reservada a esta Reserva ?
Public Const AVISO_CRIAR_TIPOTRIBUTACAO = 5852
'Deseja criar novo Tipo de Tributação ?
Public Const AVISO_CONFIRMA_EXCLUSAO_PEDIDO_VENDA = 5853 'lCodigo
'Confirma a exclusão do Pedido de Venda com o código %l?
Public Const AVISO_BLOQUEIO_TOTAL2 = 5854
'Existe Bloqueio Total no Banco de Dados, o que implica em cancelamento de todas as reservas feitas. Deseja prosseguir?
Public Const AVISO_BLOQUEIO_TOTAL = 5855
'Bloqueio Total implica em cancelamento de todas as reservas feitas. Deseja Prosseguir?
Public Const AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS = 5856 'dValorDesconto, dValorProdutos
'Desconto = %d não pode ultrapassar o Valor dos Produtos = %d. Desconto será zerado.
Public Const AVISO_CRIAR_ALMOXARIFADO = 5857 'Parâmetro: iCodigo
'Almoxarifado %i não está cadastrado no Banco de Dados. Deseja cadastrar?
Public Const AVISO_ALMOXARIFADO_TELA_PADRAO = 5859  'Parametro : sAlmoxarifado
'O almoxarifado %s da tela será o Almoxarifado Padrão.
Public Const AVISO_EXCLUSAO_ESTOQUE_PRODUTO = 5860 'Parametros: Produto e Almoxarifado
'Confirma exclusão da associação do Produto %s com o Almoxarifado %s ?
Public Const AVISO_ALMOXARIFADO_TELA_PADRAO1 = 5861
'Este Almoxarifado passa a ser o padrão por ser o primeiro a ser cadastrado esse Produto.
Public Const AVISO_ATUALIZACAO_LOTE2 = 5862 'iLote, iNumeroItensAtual
'O Número total de lotes calculados em InvetarioLotePendente para o lote %i difere
'do total exibido. Número de Lotes calculados = iNumeroItensAtual. Deseja alterar o
'Total exibidos para que fique compatível com o Total em InventarioLote?
Public Const AVISO_IGUALDADE_TOTAIS2 = 5863 'Sem parametros
'O Numero de Itens calculados através da leitura dos lotes de inventario Pendente
'são iguais aos exibidos na sua tela.
Public Const AVISO_CRIAR_LOTEINV = 5864 'Parâmetro: iLote
'O Lote de Inventário %i não está cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_TIPOPRODUTO = 5865 'Parâmetro: iTipo
'O Tipo de Produto %i não existe. Deseja criá-lo?
Public Const AVISO_CRIAR_SERIE = 5866 'Parâmetro: sSerie
'A Série %s não existe. Deseja Criá-la?
Public Const AVISO_TABELA_TELA_PADRAO = 5867
'Esta Tabela de preços passa a ser a padrão por ser a primeira a ser cadastrada para esse Produto.
Public Const AVISO_INFORMA_NUMERO_RECEBIMENTO_GRAVADO = 5868 'Parametros: lNumRecebimento
'O Recebimento foi gravado com o Número %l.
Public Const AVISO_CANCELAR_ATUALIZACAO_INVLOTE = 5869
'Confirma o cancelamento da atualização dos Inventários em Lote ?
Public Const AVISO_CUSTO_STANDARD_DIFERENTE_ALMOXARIFADO = 5870   'Parametros sProduto, dCustoStandardLido, iAlmoxarifado, dCustoStandardInserir
'O Produto %s tem Custo Standard %s, o estoque inicial a ser inserido para o almoxarifado %s tem Custo Standard %s está alteração implicará na revalorização dos estoques iniciais já cadastrados. Deseja continuar ?
Public Const AVISO_MOVIMENTO_ESTOQUE_ALTERACAO_CAMPOS = 5872 'Sem parametros
'Todos os campos com exceção do número do lote, Horas Máquina e Inicio produção não serão alterados. Deseja proseguir na alteração ?
Public Const AVISO_LOTE_PRODUTO_FILIALOP_INEXISTENTE = 5873 'Parâmetros: sLote, sProduto, iFilial
'Não existe lote %s para o produto %s da FilialOP = %i. Deseja cadastrá-lo?
Public Const AVISO_CRIAR_EMBALAGEM = 5874 'parametro:objEmbalagem.iCodigo
'A embalagem com código %s não está cadastrada. Deseja criá-la agora?
Public Const AVISO_NAO_EXISTE_REQPRODUCAO_OP = 5875 'Parametro: sCodigoOP
'Não existe Requisição de Produção para a Ordem de Produção %s. Deseja Prosseguir ?
Public Const AVISO_EXCLUIR_EMBALAGEM = 5876 'parametro:icodigo
'Confirma exclusão da Embalagem com código %i?
Public Const AVISO_REATIVACAO_OP = 5877 'parametro: sCodigo
'Deseja reativar a ordem de produção %s?
Public Const AVISO_TERMINO_ABERTURA_MES = 5878 'Sem Parâmetros
'Término da abertura do mês.


'fernando
Public Const AVISO_REPROC_MES_ESTOQUE_FECHADO = 0 'Parametros: iMes, iAno
'O Mês %i/%i do estoque está fechado. Confirma o reprocessamento?
Public Const AVISO_CANCELAR_REPROC_MOVESTOQUE = 0
'Confirma o cancelamento do reprocessamento dos movimentos de estoque ?


'VEIO DE ERROSCOM
Public Const AVISO_CONFIGURACAO_GRAVADA = 15212
'Configuração gravada.

