Attribute VB_Name = "ErrosMAT2"
''Option Explicit

'''C�digos de Erro  RESERVADO de 11200 a 11399
''Public Const ERRO_PRODUTO_SEM_TIPO = 11200 'Parametros sCodigo
'''Produto %s n�o tem Tipo de Produto associado.
''Public Const ERRO_PRODUTO_MESMA_DESCRICAO = 11201 'sDescricaoProduto
'''J� existe um Produto cadastrado com a Descri��o = %s
''Public Const ERRO_LEITURA_FORNECEDORPRODUTOFF = 11202 'Sem Parametro
'''Erro na Leitura da Tabela FornecedorProdutoFF.
''Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_PEDCOMPRA = 11204 'Par�metros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'''N�o � poss�vel excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'''filial %i est�o sendo utilizadas no Pedido de Compra de c�digo %l.
''Public Const ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF = 11205 'Par�metros: lFornecedor, iFilial, sProduto
'''Erro na tentativa de atualizar registro na tabela FornecedorProdutoFF com Fornecedor %l, Filial %i e Produto %s.
''Public Const ERRO_INSERCAO_FORNECEDORPRODUTOFF = 11206 'Par�metros: lFornecedor, sProduto
'''Erro na tentativa de inserir registro na tabela FornecedorProdutoFF com Fornecedor %l e Produto %s.
''Public Const ERRO_LOCK_FORNECEDORPRODUTOFF = 11207 'Par�metros: lFornecedor, sProduto
'''Erro na tentativa de "lock" na tabela FornecedorProdutoFF com Fornecedor %l e Produto %s.
''Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_REQUISICAOCOMPRA = 11208 'Par�metros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'''N�o � poss�vel excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'''filial %i est�o sendo utilizadas na Requisic�o de Compra de c�digo %l.
''Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_CONCORRENCIA = 11209 'Par�metros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'''N�o � poss�vel excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'''filial %i est�o sendo utilizadas na Concorr�ncia de c�digo %l.
''Public Const ERRO_PRODUTO_SEM_FORNECEDOR = 11210 'Parametro: sCodigo
'''O Produto %s n�o tem Fornecedores cadastrados nessa Filial da Empresa.
''Public Const ERRO_FORNECEDORPRODUTOFF_NAO_ENCONTRADO = 11212 'Par�metros: lFornecedor, sProduto
'''O Fornecedor %l do Produto %s n�o est� cadastrado no Banco de Dados.
''Public Const ERRO_INSERCAO_ITEMPEDCOTACAOBAIXADO = 11213 'Sem par�metros
'''Erro na tentativa de inserir registros na tabela Item Pedido Cota��o Baixado.
''Public Const ERRO_ITEMCOTACAO_VINCULADO_COTACAOITEMCONCORRENCIA = 11214 'Par�metros: lFornecedor, iFilialForn, sProduto
'''N�o � poss�vel excluir o registro Fornecedor %l, Filial %i e Produto %s pois eles est�o vinculados com
'''Cota��o Item Concor�ncia.
''Public Const ERRO_INSERCAO_ITENSCOTACAOBAIXADOS = 11215 'Sem par�metros
'''Erro na tentativa de inserir registros na tabela Itens Cota��o Baixados.
''Public Const ERRO_LOCK_PEDIDOCOTACAO = 11216 'Sem par�metros
'''Erro na tentativa de fazer "lock" na tabela Pedido Cota��o.
''Public Const ERRO_INSERCAO_PEDIDOCOTACAOBAIXADO = 11217 'Sem par�metros
'''Erro na tentativa de inserir registros na tabela Pedido Cota��o Baixado.
''Public Const ERRO_EXCLUSAO_PEDIDOCOTACAO = 11218 'Sem par�metros
'''Erro na exclus�o de registro na tabela Pedido Cota��o.
''Public Const ERRO_LEITURA_COTACAOPRODUTO = 11219 'Sem par�metros
'''Erro na leitura da tabela Cota��o Produto.
''Public Const ERRO_LOCK_COTACAO = 11220 'Sem par�metros
'''Erro na tantativa de fazer "lock" na tabela Cota��o.
''Public Const ERRO_INSERCAO_COTACAOPRODUTOBAIXADO = 11221 'Sem par�metros
'''Erro na tentativa de inserir registros na tabelaCota��o Produto Baixado.
''Public Const ERRO_LOCK_COTACAOPRODUTO = 11222 'Sem par�metros
'''Erro na tentativa de fazer "lock" na tabela Cota��o Produto.
''Public Const ERRO_EXCLUSAO_COTACAOPRODUTO = 11223 'Sem par�metros
'''Erro na tentativa de excluir registros da tabela Cota��o Produto.
''Public Const ERRO_INSERCAO_COTACAOBAIXADA = 11224 'Sem par�metros
'''Erro na tentativa de inserir registros na tabela Cota��o Baixada.
''Public Const ERRO_EXCLUSAO_COTACAO = 11225 'Sem par�metros
'''Erro na exclus�o de registro na tabela Cota��o.
''Public Const ERRO_EXCLUSAO_FORNECEDORPRODUTOFF = 11226 'Par�metros: sProduto
'''Erro na tentativa de excluir registros de FornecedorProdutoFF com Produto = %s.
''Public Const ERRO_PRODUTO_FORNECEDORPRODUTOFF = 11227 'sem parametros
'''O Produto em quest�o n�o pode ser exclu�do pois est� relacionado a Fornecedor ProdutoFF.
''Public Const ERRO_LEITURA_SLDDIAFORN = 11228 'Sem parametros
'''Ocorreu um erro na leitura da tabela de saldos di�rios de fornecedor.
''Public Const ERRO_LEITURA_SLDMESFORN = 11229 'Sem parametros
'''Ocorreu um erro na leitura da tabela de saldos mensais de fornecedor.
''Public Const ERRO_PRODUTO_SLDMESFORN = 11230 'sem parametros
'''O Produto em quest�o n�o pode ser exclu�do pois est� relacionado a um Saldo Mensal de Fornecedor.
''Public Const ERRO_PRODUTO_SLDDIAFORN = 11231 'sem parametros
'''O Produto em quest�o n�o pode ser exclu�do pois est� relacionado a um Saldo Di�rio de Fornecedor.
''Public Const ERRO_PRODUTO_ITENSPEDCOMPRA = 11232 'sem parametros
'''O Produto em quest�o n�o pode ser exclu�do pois est� relacionado a um Item de Pedido de Compra.
''Public Const ERRO_PRODUTO_ITENSREQCOMPRA = 11233 'sem parametros
'''O Produto em quest�o n�o pode ser exclu�do pois est� relacionado a um Item de Requisi��o de Compra.
''Public Const ERRO_LEITURA_ITENSCONCORRENCIA1 = 11234 'sem parametros
'''Erro na leitura da tabela ItensConcorrencia
''Public Const ERRO_PRODUTO_ITENSCONCORRENCIA = 11235 'sem parametros
'''O Produto em quest�o n�o pode ser exclu�do pois est� relacionado a um Item de Concorr�ncia.
''Public Const ERRO_PRODUTO_COTACAOPRODUTO = 11236 'sem parametros
'''O Produto em quest�o n�o pode ser exclu�do pois est� relacionado a Cota��o Produto.
''
''
''
''
'''VEIO DE ERROSCOM
''Public Const ERRO_LEITURA_CONCORRENCIA = 12024 'Sem Par�metros
'''Erro de leitura na tabela de concorr�ncias.
''Public Const ERRO_LEITURA_ITEMPEDCOTACAO = 12101 'Parametro: lCodPedidoCotacao
'''Erro na leitura dos Itens do Pedido de Cota��o %l.
''Public Const ERRO_LEITURA_ITENSCOTACAO = 12102 'Sem parametros
'''Erro na leitura da tabela ItensCotacao.
''Public Const ERRO_LEITURA_PEDIDOCOTACAO = 12103 'Parametro: lC�digo
'''Erro na leitura do Pedido de Cota��o com o c�digo %l.
''Public Const ERRO_LEITURA_COTACAO = 12104 'Sem parametros
'''Erro na leitura da tabela Cotacao.
''Public Const ERRO_LEITURA_ITENSPEDCOMPRA = 12105 'Sem Parametros
'''Erro na leitura da tabela ItensPedCompra.
''Public Const ERRO_LEITURA_COTACAOITEMCONCORRENCIA = 12106 'Sem parametros
'''Erro na leitura da tabela CotacaoItemConcorrencia.
''Public Const ERRO_LOCK_ITEMPEDCOTACAO = 12108 'Parametro: lCodPedidoCotacao
'''N�o conseguiu fazer o lock nos itens do pedido de cota��o %l.
''Public Const ERRO_LOCK_ITENSCOTACAO = 12109 'Parametro: lCodigo
'''N�o conseguiu fazer o lock dos itens de cota��o do pedido de cota��o %l.
''Public Const ERRO_EXCLUSAO_ITEMPEDCOTACAO = 12116 'Parametro lCodigo
'''Erro na exclus�o dos itens do pedido de cota��o %l.
''Public Const ERRO_EXCLUSAO_ITENSCOTACAO = 12117 'Parametro lCodigo
'''Erro na exclus�o dos itens de cota��o do pedido de cota��o %l.
''
