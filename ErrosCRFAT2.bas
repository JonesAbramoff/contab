Attribute VB_Name = "ErrosCRFAT2"
''Option Explicit

'''Códigos de Erros - Reservado de 13100 a 13299
''Public Const ERRO_LEITURA_CONTRATOFORNECIMENTO = 13100 'Sem parametros
'''Erro na leitura da tabela de ContratoFornecimento.
''Public Const ERRO_LEITURA_ITENS_CONTRATO = 13101 'Sem parametros
'''Erro na leitura da tabela de ItensContrato.
''Public Const ERRO_FORNECEDOR_REL_PEDIDOCOMPRA = 13102 'Parâmetros: lCodigoForn, lCodigoPedidoCompra
'''O Fornecedor %l está relacionado com o Pedido de Compra %l.
''Public Const ERRO_FORNECEDOR_REL_CONTRATOFORNECIMENTO = 13103 'Parâmetros: lCodigoForn, lCodigoContrato
'''O Fornecedor %l está relacionado com o Contrato de Fornecimento %l.
''Public Const ERRO_FORNECEDOR_REL_ITEMCONCORRENCIA = 13104 'Parâmetros: lCodigoForn
'''O Fornecedor %l está relacionado com Item de Concorrência.
''Public Const ERRO_FORNECEDOR_REL_CONCORRENCIA = 13105 'Parâmetros: lCodigoForn, lCodigoConcorrencia
'''O Fornecedor %l está relacionado com a Concorrência %l.
''Public Const ERRO_FORNECEDOR_REL_REQCOMPRA = 13106 'Parametro: lCodFornecedor, lCodRequisicaoCompra
'''O Fornecedor %l está relacionado com a Requisição de Compra %l.
''Public Const ERRO_FORNECEDOR_REL_PEDIDOCOTACAO = 13107 'Parametro: lCodFornecedor, lCodPedidoCotacao
'''O Fornecedor %l está relacionado com o Pedido de Cotação %l.
''Public Const ERRO_FORNECEDOR_REL_COTACAO = 13108 'Parametro: lCodFornecedor, lCodCotacao
'''O Fornecedor %l está relacionado com a Cotação %l.
''Public Const ERRO_FORNECEDOR_REL_ITEMREQCOMPRA = 13109 'Parametro: lCodFornecedor
'''O Fornecedor %l está relacionado com um item de Requisição de Compra.
''Public Const ERRO_FORNECEDOR_REL_REQMODELO = 13110 'Parametro: lCodFornecedor, lCodRequisicaoModelo
'''O Fornecedor %l está relacionado com a Requisição Modelo %l.
''Public Const ERRO_FORNECEDOR_REL_COTACAOPRODUTO = 13111 'Parametro: lCodFornecedor
'''O Fornecedor %l está relacionado com Cotação Produto.
''Public Const ERRO_LEITURA_ITENSREQMODELO2 = 13112 'sem parâmetros
'''Erro na leitura da tabela ItensReqModelo.
''Public Const ERRO_FORNECEDOR_REL_ITENSREQMODELO = 13113 'Parametro: lCodFornecedor
'''O Fornecedor %l está relacionado com um item de Requisição Modelo.
''Public Const ERRO_LEITURA_COTACAOPRODUTOTODAS = 13114 'Sem parametros
'''Erro na leitura da tabela CotacaoProdutoTodas.
''Public Const ERRO_LEITURA_CONCORRENCIATODAS = 13115 'Sem Parâmetros
'''Erro de leitura na tabela de ConcorrenciaTodas.
''Public Const ERRO_FORNECEDOR_REL_FORNECEDOR_PRODUTOFF = 13116 'Parâmetros: lCodFornecedor, sProduto
'''O Fornecedor %l está relacionado com o Produto %s.
''
''
''
''
'''Veio de ErrosCOM
''Public Const ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO = 12004 'Parametros sCodigo,iFilialEmpresa
'''Registro na tabela ComprasConfig com Código=%s e FilialEmpresa=%i não foi encontrado.
''Public Const ERRO_LEITURA_COMPRASCONFIG = 12005 'Parametro sCodigo
'''Erro na leitura de %s na tabela de ComprasConfig.
''
