Attribute VB_Name = "ErrosPV"
Option Explicit

'Códigos de Erros - Reservado de 7200 até 7299
Public Const ERRO_LEITURA_ITENSPEDIDODEVENDA2 = 7200 'lCodPedido
'''Erro na tentativa de leitura dos Item do Pedido de Venda de Código = %l na Tabela ItensPedidoDeVenda.
Public Const ERRO_LEITURA_PARCELASPEDIDODEVENDA = 7201 'lCodigo
'Erro na tentativa de leitura das Parcelas do Pedido de Venda de Código = %l na Tabela ParcelaPedidoDeVenda.
Public Const ERRO_PRODUTO_QUANTIDADE_FATURADA = 7204
'Não é possível alterar Produto que tenha quantidade faturada.
Public Const ERRO_PRODUTO_NAO_PODE_SER_VENDIDO = 7205
'Produto não pode ser vendido.
Public Const ERRO_LEITURA_NFSREC_PV = 7206 'lCodPedVenda
'Erro na tentativade leitura de Nota Fiscal associada ao Pedido de Venda de Número %l na tabela NFsRec.
Public Const ERRO_COMPL_PV_TIPO_INVALIDO = 7207 'sem parametros
'registro na tabela TributacaoComplPV com tipo inválido
Public Const ERRO_INSERCAO_ITENSPEDIDODEVENDABAIXADOS = 7209 'Parametros = lCodPedido
'Ocorreu um erro na tentativa de inserir um registro na tabela de Itens de Pedido de Venda Baixados. Pedido = %l.
Public Const ERRO_EXCLUSAO_ITENSPEDIDODEVENDA = 7210 'Parametros = lCodPedido
'Ocorreu um erro na tentativa de excluir um registro da tabela de Itens de Pedido de Venda. Pedido = %l.
Public Const ERRO_NIVEL_NAO_INFORMADO = 7211
'O Nivel deve ser informado.
Public Const ERRO_ALMOXARIFADO_INICIAL_MAIOR = 7212
'O Almoxarifado Inicial é maior que o Final.'
Public Const ERRO_PRODUTO_INICIAL_MAIOR = 7213
'O Produto Inicial é maior que o Final.
Public Const ERRO_FILIALEMPRESA_INICIAL_MAIOR = 7214
'A FilialEmpresa Inicial é maior que a Final.'
Public Const ERRO_INSERCAO_SLDMESFAT = 7215 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de faturamento. Ano=%i, FilialEmpresa=%i, Produto=%s.
Public Const ERRO_INSERIR_FILIALCLIENTEFILEMP = 7217 'Sem Parametros
'Erro na tentativa de inserir na tabela FilialClienteFilEmp.
Public Const ERRO_FORNECEDOR_BENEF_NAO_PREENCHIDA = 7219
'Os dados do Fornecedor que irá beneficiar o material em questão precisa ser preenchido.
Public Const ERRO_FILIAL_FORNECEDOR_BENEF_NAO_PREENCHIDA = 7220
'É obrigatório o preenchimento da Filial do Fornecedor Beneficiado.
Public Const ERRO_DESCRICAOITEM_NAO_PREENCHIDA = 7221 'Parametro iIndice
'Descrição do item %i do Grid Itens não foi preenchida.
Public Const ERRO_ITEM_TIPO_INVALIDO = 7222
'Erro interno: o item %d não foi tratado.
Public Const ERRO_ITEM_MNEMONICO_INVALIDO = 7223
'Erro interno: o Mnemônico %s não foi tratado.
Public Const ERRO_LOCK_TIPODEBLOQUEIO = 7224 'iTipoBloqueio
'Erro na tentativa de faxer lock no tipo de bloqueio %i na tabela de tipos de bloqueio.
Public Const ERRO_LEITURA_TITULOSRECBAIXADOS1 = 7225 'Parâmetros: lCodCliente, iFilialEmpresa
'Erro na leitura da tabela TitulosRecBaixados com cliente %l e Filial Empresa %i.
Public Const ERRO_LEITURA_MOVCC_BAIXAPR_TITREC = 7226 'Sem parâmetros
'Erro na leitura da tabela MovCC_BaixaPR_TitRec.
Public Const ERRO_LEITURA_PARCELASPEDIDODEVENDA_BAIXADAS = 7227 'lCodigo
'Erro na tentativa de leitura das Parcelas Baixadas do Pedido de Venda de Código = %l na Tabela ParcelasPedidoDeVendaBaixado.
Public Const ERRO_ITEM_PV_QUANT_FAT_POSITIVA = 7228 'Parametro: iItemAtual
'Item %i do Pedido de Venda tem quantidade faturada, não pode ser excluído.
Public Const ERRO_NAO_HA_RESERVAS_PARA_LIBERAR = 7229 'Sem Parametros
'Não há reservas neste pedido para serem liberadas.
Public Const ERRO_PRODUTOFILIAL_INEXISTENTE_FILIALFATURAMENTO = 7230 'Parametros: sProduto, iFilialEmpresaFaturamento
'O Produto %s não está cadastrado na Filial Faturamento %i.
Public Const ERRO_EXCLUSAO_PARCELASPEDIDODEVENDA_BAIXADAS = 7231 'Parametro: lCodPedido
'Ocorreu um erro na tentativa de excluir um registro da tabela de Parcelas de Pedido de Venda Baixado. Pedido = %l.
Public Const ERRO_LEITURA_COMISSOESPEDVENDASBAIXADOS = 7232 'Sem parâmetro
'Erro na leitura de registros de comissões de pedidos de venda baixadas.
Public Const ERRO_EXCLUSAO_COMISSOESPEDVENDASBAIXADOS = 7233 'Parametro: iFilialEmpresa, lPedidoVenda
'Erro na tentativa de excluir registro da tabela de ComissoesPedVendasBaixados da Filial %i e Pedido %l.
Public Const ERRO_LEITURA_BLOQUEIOSPVBAIXADOS = 7234 'lCodigo
'Erro na tentativa de leitura dos Bloqueios associados ao Pedido de Venda de Código = %l na Tabela BloqueiosPV.
Public Const ERRO_EXCLUSAO_BLOQUEIOSPVBAIXADOS = 7235 'lCodigo
'Erro na tentativa de exclusão dos Bloqueios associados ao Pedido de Venda de Código = %l na Tabela BloqueiosPV.
Public Const ERRO_PV_NAOEDITAVEL_FILIALEMPRESA_DIFERENTE = 7236 'Parametros: lPedido, iFilialPedido
'O Pedido de Venda %s não é editavel pois ele pertence a Filial Empresa %s.
Public Const ERRO_STATUS_OP_ABERTO = 7237 'sem parametros
'A ordem de produção foi alterada e seu novo Status = Aberto. É preciso
'realizar uma nova consulta para produzir as alterações desejadas.
Public Const ERRO_STATUS_OP_BAIXADO = 7238 'sem parametros
'A ordem de produção foi alterada e seu novo Status = Baixado. É preciso
'realizar uma nova consulta para produzir as alterações desejadas.
Public Const ERRO_NF_SEM_DOCCPR_VINCULADO = 7239 'Sem parâmetros
'Essa Nota Fiscal não possui documentos do Contas a pagar / Contas a Receber vinculados a ela.




'VEIO Erros Fat
Public Const ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO = 8057 'Parametro objTipo.iCodigo
'Tipo de Bloqueio %i não cadastrado
Public Const ERRO_LEITURA_TIPOSDEBLOQUEIO = 8085 'Sem Parâmetros
'Erro de leitura na Tabela TiposDeBloqueio
Public Const ERRO_LEITURA_SLDMESFAT = 8201 'Parametros  iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de faturamento (SldMesFat). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LOCK_SLDMESFAT = 8202 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de faturamento (SldMesFat). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_ATUALIZACAO_SLDMESFAT = 8203  'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos mensais de faturamento. Ano=%i, FilialEmpresa=%i, Produto=%s.
Public Const ERRO_LEITURA_SLDDIAFAT = 8204 'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na leitura da tabela de saldos diários de faturamento. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_INSERCAO_SLDDIAFAT = 8205 'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na inclusão de registro na tabela de saldos diários de faturamento. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_LOCK_SLDDIAFAT = 8206  'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos diários de faturamento. FilialEmpresa=%i, Produto=%s, Data=%s.
Public Const ERRO_ATUALIZACAO_SLDDIAFAT = 8207  'Parametros iFilialProduto, sProduto, sData
'Ocorreu um erro na tentativa de atualizar um registro na tabela de saldos diários de faturamento. FilialEmpresa=%i, Produto=%s, Data=%s.



'Códigos de Avisos - Reservado de 5700 até 5799
Public Const AVISO_CONFIRMA_LIBERACAO_RESERVAS = 5700 'Sem Parametros
'As reservas deste pedido serão apagadas. Confirma?

