Attribute VB_Name = "ErrosCRFAT2"
''Option Explicit

'''C�digos de Erros - Reservado de 13100 a 13299
''Public Const ERRO_LEITURA_CONTRATOFORNECIMENTO = 13100 'Sem parametros
'''Erro na leitura da tabela de ContratoFornecimento.
''Public Const ERRO_LEITURA_ITENS_CONTRATO = 13101 'Sem parametros
'''Erro na leitura da tabela de ItensContrato.
''Public Const ERRO_FORNECEDOR_REL_PEDIDOCOMPRA = 13102 'Par�metros: lCodigoForn, lCodigoPedidoCompra
'''O Fornecedor %l est� relacionado com o Pedido de Compra %l.
''Public Const ERRO_FORNECEDOR_REL_CONTRATOFORNECIMENTO = 13103 'Par�metros: lCodigoForn, lCodigoContrato
'''O Fornecedor %l est� relacionado com o Contrato de Fornecimento %l.
''Public Const ERRO_FORNECEDOR_REL_ITEMCONCORRENCIA = 13104 'Par�metros: lCodigoForn
'''O Fornecedor %l est� relacionado com Item de Concorr�ncia.
''Public Const ERRO_FORNECEDOR_REL_CONCORRENCIA = 13105 'Par�metros: lCodigoForn, lCodigoConcorrencia
'''O Fornecedor %l est� relacionado com a Concorr�ncia %l.
''Public Const ERRO_FORNECEDOR_REL_REQCOMPRA = 13106 'Parametro: lCodFornecedor, lCodRequisicaoCompra
'''O Fornecedor %l est� relacionado com a Requisi��o de Compra %l.
''Public Const ERRO_FORNECEDOR_REL_PEDIDOCOTACAO = 13107 'Parametro: lCodFornecedor, lCodPedidoCotacao
'''O Fornecedor %l est� relacionado com o Pedido de Cota��o %l.
''Public Const ERRO_FORNECEDOR_REL_COTACAO = 13108 'Parametro: lCodFornecedor, lCodCotacao
'''O Fornecedor %l est� relacionado com a Cota��o %l.
''Public Const ERRO_FORNECEDOR_REL_ITEMREQCOMPRA = 13109 'Parametro: lCodFornecedor
'''O Fornecedor %l est� relacionado com um item de Requisi��o de Compra.
''Public Const ERRO_FORNECEDOR_REL_REQMODELO = 13110 'Parametro: lCodFornecedor, lCodRequisicaoModelo
'''O Fornecedor %l est� relacionado com a Requisi��o Modelo %l.
''Public Const ERRO_FORNECEDOR_REL_COTACAOPRODUTO = 13111 'Parametro: lCodFornecedor
'''O Fornecedor %l est� relacionado com Cota��o Produto.
''Public Const ERRO_LEITURA_ITENSREQMODELO2 = 13112 'sem par�metros
'''Erro na leitura da tabela ItensReqModelo.
''Public Const ERRO_FORNECEDOR_REL_ITENSREQMODELO = 13113 'Parametro: lCodFornecedor
'''O Fornecedor %l est� relacionado com um item de Requisi��o Modelo.
''Public Const ERRO_LEITURA_COTACAOPRODUTOTODAS = 13114 'Sem parametros
'''Erro na leitura da tabela CotacaoProdutoTodas.
''Public Const ERRO_LEITURA_CONCORRENCIATODAS = 13115 'Sem Par�metros
'''Erro de leitura na tabela de ConcorrenciaTodas.
''Public Const ERRO_FORNECEDOR_REL_FORNECEDOR_PRODUTOFF = 13116 'Par�metros: lCodFornecedor, sProduto
'''O Fornecedor %l est� relacionado com o Produto %s.
''
''
''
''
'''Veio de ErrosCOM
''Public Const ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO = 12004 'Parametros sCodigo,iFilialEmpresa
'''Registro na tabela ComprasConfig com C�digo=%s e FilialEmpresa=%i n�o foi encontrado.
''Public Const ERRO_LEITURA_COMPRASCONFIG = 12005 'Parametro sCodigo
'''Erro na leitura de %s na tabela de ComprasConfig.
''
