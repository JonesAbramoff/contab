Attribute VB_Name = "ErrosMAT2"
''Option Explicit

'''Códigos de Erro  RESERVADO de 11200 a 11399
''Public Const ERRO_PRODUTO_SEM_TIPO = 11200 'Parametros sCodigo
'''Produto %s não tem Tipo de Produto associado.
''Public Const ERRO_PRODUTO_MESMA_DESCRICAO = 11201 'sDescricaoProduto
'''Já existe um Produto cadastrado com a Descrição = %s
''Public Const ERRO_LEITURA_FORNECEDORPRODUTOFF = 11202 'Sem Parametro
'''Erro na Leitura da Tabela FornecedorProdutoFF.
''Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_PEDCOMPRA = 11204 'Parâmetros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'''Não é possível excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'''filial %i estão sendo utilizadas no Pedido de Compra de código %l.
''Public Const ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF = 11205 'Parâmetros: lFornecedor, iFilial, sProduto
'''Erro na tentativa de atualizar registro na tabela FornecedorProdutoFF com Fornecedor %l, Filial %i e Produto %s.
''Public Const ERRO_INSERCAO_FORNECEDORPRODUTOFF = 11206 'Parâmetros: lFornecedor, sProduto
'''Erro na tentativa de inserir registro na tabela FornecedorProdutoFF com Fornecedor %l e Produto %s.
''Public Const ERRO_LOCK_FORNECEDORPRODUTOFF = 11207 'Parâmetros: lFornecedor, sProduto
'''Erro na tentativa de "lock" na tabela FornecedorProdutoFF com Fornecedor %l e Produto %s.
''Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_REQUISICAOCOMPRA = 11208 'Parâmetros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'''Não é possível excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'''filial %i estão sendo utilizadas na Requisicão de Compra de código %l.
''Public Const ERRO_FORNECEDORPRODUTOFF_UTILIZADO_CONCORRENCIA = 11209 'Parâmetros: sProduto, lFornecedor, iFilialForn, lCodPedidoCompra
'''Não é possível excluir o Fornecedor Filial Produto pois o produto %s, o fornecedor %l e a
'''filial %i estão sendo utilizadas na Concorrência de código %l.
''Public Const ERRO_PRODUTO_SEM_FORNECEDOR = 11210 'Parametro: sCodigo
'''O Produto %s não tem Fornecedores cadastrados nessa Filial da Empresa.
''Public Const ERRO_FORNECEDORPRODUTOFF_NAO_ENCONTRADO = 11212 'Parâmetros: lFornecedor, sProduto
'''O Fornecedor %l do Produto %s não está cadastrado no Banco de Dados.
''Public Const ERRO_INSERCAO_ITEMPEDCOTACAOBAIXADO = 11213 'Sem parâmetros
'''Erro na tentativa de inserir registros na tabela Item Pedido Cotação Baixado.
''Public Const ERRO_ITEMCOTACAO_VINCULADO_COTACAOITEMCONCORRENCIA = 11214 'Parâmetros: lFornecedor, iFilialForn, sProduto
'''Não é possível excluir o registro Fornecedor %l, Filial %i e Produto %s pois eles estão vinculados com
'''Cotação Item Concorência.
''Public Const ERRO_INSERCAO_ITENSCOTACAOBAIXADOS = 11215 'Sem parâmetros
'''Erro na tentativa de inserir registros na tabela Itens Cotação Baixados.
''Public Const ERRO_LOCK_PEDIDOCOTACAO = 11216 'Sem parâmetros
'''Erro na tentativa de fazer "lock" na tabela Pedido Cotação.
''Public Const ERRO_INSERCAO_PEDIDOCOTACAOBAIXADO = 11217 'Sem parâmetros
'''Erro na tentativa de inserir registros na tabela Pedido Cotação Baixado.
''Public Const ERRO_EXCLUSAO_PEDIDOCOTACAO = 11218 'Sem parâmetros
'''Erro na exclusão de registro na tabela Pedido Cotação.
''Public Const ERRO_LEITURA_COTACAOPRODUTO = 11219 'Sem parâmetros
'''Erro na leitura da tabela Cotação Produto.
''Public Const ERRO_LOCK_COTACAO = 11220 'Sem parâmetros
'''Erro na tantativa de fazer "lock" na tabela Cotação.
''Public Const ERRO_INSERCAO_COTACAOPRODUTOBAIXADO = 11221 'Sem parâmetros
'''Erro na tentativa de inserir registros na tabelaCotação Produto Baixado.
''Public Const ERRO_LOCK_COTACAOPRODUTO = 11222 'Sem parâmetros
'''Erro na tentativa de fazer "lock" na tabela Cotação Produto.
''Public Const ERRO_EXCLUSAO_COTACAOPRODUTO = 11223 'Sem parâmetros
'''Erro na tentativa de excluir registros da tabela Cotação Produto.
''Public Const ERRO_INSERCAO_COTACAOBAIXADA = 11224 'Sem parâmetros
'''Erro na tentativa de inserir registros na tabela Cotação Baixada.
''Public Const ERRO_EXCLUSAO_COTACAO = 11225 'Sem parâmetros
'''Erro na exclusão de registro na tabela Cotação.
''Public Const ERRO_EXCLUSAO_FORNECEDORPRODUTOFF = 11226 'Parâmetros: sProduto
'''Erro na tentativa de excluir registros de FornecedorProdutoFF com Produto = %s.
''Public Const ERRO_PRODUTO_FORNECEDORPRODUTOFF = 11227 'sem parametros
'''O Produto em questão não pode ser excluído pois está relacionado a Fornecedor ProdutoFF.
''Public Const ERRO_LEITURA_SLDDIAFORN = 11228 'Sem parametros
'''Ocorreu um erro na leitura da tabela de saldos diários de fornecedor.
''Public Const ERRO_LEITURA_SLDMESFORN = 11229 'Sem parametros
'''Ocorreu um erro na leitura da tabela de saldos mensais de fornecedor.
''Public Const ERRO_PRODUTO_SLDMESFORN = 11230 'sem parametros
'''O Produto em questão não pode ser excluído pois está relacionado a um Saldo Mensal de Fornecedor.
''Public Const ERRO_PRODUTO_SLDDIAFORN = 11231 'sem parametros
'''O Produto em questão não pode ser excluído pois está relacionado a um Saldo Diário de Fornecedor.
''Public Const ERRO_PRODUTO_ITENSPEDCOMPRA = 11232 'sem parametros
'''O Produto em questão não pode ser excluído pois está relacionado a um Item de Pedido de Compra.
''Public Const ERRO_PRODUTO_ITENSREQCOMPRA = 11233 'sem parametros
'''O Produto em questão não pode ser excluído pois está relacionado a um Item de Requisição de Compra.
''Public Const ERRO_LEITURA_ITENSCONCORRENCIA1 = 11234 'sem parametros
'''Erro na leitura da tabela ItensConcorrencia
''Public Const ERRO_PRODUTO_ITENSCONCORRENCIA = 11235 'sem parametros
'''O Produto em questão não pode ser excluído pois está relacionado a um Item de Concorrência.
''Public Const ERRO_PRODUTO_COTACAOPRODUTO = 11236 'sem parametros
'''O Produto em questão não pode ser excluído pois está relacionado a Cotação Produto.
''
''
''
''
'''VEIO DE ERROSCOM
''Public Const ERRO_LEITURA_CONCORRENCIA = 12024 'Sem Parâmetros
'''Erro de leitura na tabela de concorrências.
''Public Const ERRO_LEITURA_ITEMPEDCOTACAO = 12101 'Parametro: lCodPedidoCotacao
'''Erro na leitura dos Itens do Pedido de Cotação %l.
''Public Const ERRO_LEITURA_ITENSCOTACAO = 12102 'Sem parametros
'''Erro na leitura da tabela ItensCotacao.
''Public Const ERRO_LEITURA_PEDIDOCOTACAO = 12103 'Parametro: lCódigo
'''Erro na leitura do Pedido de Cotação com o código %l.
''Public Const ERRO_LEITURA_COTACAO = 12104 'Sem parametros
'''Erro na leitura da tabela Cotacao.
''Public Const ERRO_LEITURA_ITENSPEDCOMPRA = 12105 'Sem Parametros
'''Erro na leitura da tabela ItensPedCompra.
''Public Const ERRO_LEITURA_COTACAOITEMCONCORRENCIA = 12106 'Sem parametros
'''Erro na leitura da tabela CotacaoItemConcorrencia.
''Public Const ERRO_LOCK_ITEMPEDCOTACAO = 12108 'Parametro: lCodPedidoCotacao
'''Não conseguiu fazer o lock nos itens do pedido de cotação %l.
''Public Const ERRO_LOCK_ITENSCOTACAO = 12109 'Parametro: lCodigo
'''Não conseguiu fazer o lock dos itens de cotação do pedido de cotação %l.
''Public Const ERRO_EXCLUSAO_ITEMPEDCOTACAO = 12116 'Parametro lCodigo
'''Erro na exclusão dos itens do pedido de cotação %l.
''Public Const ERRO_EXCLUSAO_ITENSCOTACAO = 12117 'Parametro lCodigo
'''Erro na exclusão dos itens de cotação do pedido de cotação %l.
''
