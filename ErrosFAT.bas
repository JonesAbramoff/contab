Attribute VB_Name = "ErrosFAT"
Option Explicit

'Códigos de Erro - Reservado de 8000 a 8299
Public Const ERRO_LEITURA_FATCONFIG2 = 8000 '%s chave %d FilialEmpresa
'Erro na leitura da tabela FATConfig. Codigo = %s Filial = %i
Public Const ERRO_FATCONFIG_INEXISTENTE = 8001 '%s chave %d FilialEmpresa
'Não foi encontrado registro em FATConfig. Codigo = %s Filial = %i
Public Const ERRO_ATUALIZACAO_FATCONFIG = 8002 '%s chave %d FilialEmpresa
'Erro na gravação da tabela FATConfig. Codigo = %s Filial = %i
Public Const ERRO_SEM_BLOQUEIOS_PV_SEL = 8004 'Sem parâmetro
'Não há bloqueios dentro dos critérios de seleção informados.
Public Const ERRO_LEITURA_PEDIDOS_VENDA_GERACAO_NF = 8005 'Sem parâmetro
'Erro na leitura de pedidos de venda para a geração de notas fiscais
Public Const ERRO_SEM_PEDIDOS_VENDA_ENCONTRADOS = 8006 'Sem parâmetro
'Não há pedidos de venda dentro dos critérios de seleção informados.
Public Const ERRO_SEM_NFSREC_GERACAO_FATURA = 8007 'Sem parâmetro
'Não há notas fiscais a serem faturadas dentro dos critérios de seleção informados.
Public Const ERRO_LEITURA_NFSREC_GERACAO_FATURA = 8008 'Sem parâmetro
'Erro na leitura de notas fiscais para geração de fatura
Public Const ERRO_SERIE_NAO_PREENCHIDA = 8012
'A Serie deve ser preenchida.
Public Const ERRO_ATUALIZACAO_COMISSOESPEDVENDAS = 8013 'Parametro: lNumIntDoc
'Erro na tentativa de atualizar as Comissões do Pedido de Vendas de Número Interno %l
Public Const ERRO_VALOREMISSAO_COMISSAO_NAO_INFORMADO = 8014 'iComissao
'O Valor de Emissão da Comissao %i do Título não foi informado.
Public Const ERRO_INCLUSAO_COMISSOESPEDVENDAS = 8015 'Sem parametro
'Erro na tentativa incluir registro na tabela de Comissões de Pedido de Vendas.
Public Const ERRO_VALOR_EMISSAO_MAIOR = 8017 'Sem Parâmetros
'Valor da comissão na emissão não pode ser maior que o valor total da comissão.
Public Const ERRO_PERCENTUAL_EMISSAO_NAO_INFORMADO = 8020 'iComissao
'O Percentual de Emissão da Comissão %i não foi informado.
Public Const ERRO_PEDIDO_VENDA_BAIXADO_MODIFICACAO = 8021 'Parametro: lCodigo
'O Pedido de Venda número %l já foi faturado. Não pode ser modificado.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI = 8022
'O item de código não existe na tabela de Tipos de Tributação para IPI.
Public Const ERRO_FUNDAMENTACAO_NAO_PREENCHIDA = 8023 'Sem parâmetro
'O preenchimento da Fundamentação é obrigatório.
Public Const ERRO_DESTINO_NAO_PREENCHIDO = 8024 'Sem parâmetro
'O preenchimento do destino é obrigatório.
Public Const ERRO_LEITURA_CATEGORIAPRODUTOITEM1 = 8025 'Parâmetro: sCategoria
'Erro na leitura dos registros da categoria %s da tabela Categoria Produto.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS = 8028 'Parâmetro: iCodigo
'O item de código %i não existe na tabela de Tipos de Tributação para ICMS.
Public Const ERRO_LOCK_ESTADOS = 8029  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de itens das Categorias de Clientes.
Public Const ERRO_PEDVENDA_FATURADO_ALTERACAO_CLIENTE = 8037 'Parametros: lCodigoPV, lClienteBD, lClienteTela
'Pedido de Venda %l tem ítens faturados. Não é possível alterar Cliente %l no Banco de Dados para %l da Tela.
Public Const ERRO_PEDVENDA_FATURADO_ALTERACAO_FILIAL_CLIENTE = 8038 'Parametros: lCodigoPV, iFilialBD, iFilialTela
'Pedido de Venda %l tem ítens faturados. Não é possível alterar Filial Cliente %i no Banco de Dados para %i da Tela.
Public Const ERRO_PEDVENDA_FATURADO_ALTERACAO_NATUREZA = 8039 'Parâmetros: lCodigoPV, sNaturezaBD, sNaturezaTela
'Pedido de Venda %l tem ítens faturados. Não é possível alterar NaturezaFilial Cliente %i no Banco de Dados para %i da Tela.
Public Const ERRO_LEITURA_FATCONFIG = 8042 'Parametro Codigo
'Erro na leitura da tabela FATConfig. Codigo = %s.
Public Const ERRO_LEITURA_BLOQUEIOPV = 8045 'Parametros: iFilialEmpresa, lPedidoDeVendas, iTipoDeBloqueio
'Ocorreu um erro na leitura de um registro da tabela de Bloqueios de Pedido de Venda. Filial=%i, Pedido=%l, Tipo de Bloqueio=%i.
Public Const ERRO_LOCK_BLOQUEIOPV = 8046 'Parametros: iFilialEmpresa, lPedidoDeVendas, iSequencial
'Ocorreu um erro na tentativa de fazer um "lock" de um registro da tabela de Bloqueios de Pedido de Venda. Filial=%i, Pedido=%l, Sequencial=%i.
Public Const ERRO_LEITURA_CANALVENDA = 8047 'Sem parametros
'Erro de Leitura na tabela CanalVenda
Public Const ERRO_NOME_REDUZIDO_CANALVENDA_REPETIDO = 8050 'Sem parametros
'Erro Nome reduzido já é utilizado por canal com outro Código
Public Const ERRO_INSERCAO_CANALVENDA = 8051 'Parametro objCanal.iCodigo
'Erro na inserção do Canal %i
Public Const ERRO_ATUALIZACAO_CANALVENDA = 8052 'Parametro objCanal.iCodigo
'Erro na Atualização do Canal %i
Public Const ERRO_EXCLUSAO_CANALVENDA = 8053 'Parametro objCanal.iCodigo
'Erro exclusão Canal de Venda % i
Public Const ERRO_DESCRICAONF_NAO_PREENCHIDA = 8054 'Sem parametros
'O preenchimento da Descricão da Nota Fiscal é obrigatório.
Public Const ERRO_LEITURA_TIPODEBLOQUEIO = 8055 'Sem parametros
'Erro de Leitura na tabela TipodeBloqueio
Public Const ERRO_LEITURA_TIPODEBLOQUEIO1 = 8056 'Parametros objTipo.iCodigo
'Erro na leitura do Tipo %i da tabela Tipo de Bloqueio
Public Const ERRO_NOME_REDUZIDO_TIPODEBLOQUEIO_REPETIDO = 8058 'Sem parametros
'Erro Nome reduzido já é utilizado por Tipo com outro Código
Public Const ERRO_INSERCAO_TIPODEBLOQUEIO = 8059 'Parametro objTipo.iCodigo
'Erro na inserção do Tipo %i
Public Const ERRO_ATUALIZACAO_TIPODEBLOQUEIO = 8060 'Parametro objTipo.iCodigo
'Erro na Atualização do Tipo %i
Public Const ERRO_EXCLUSAO_TIPODEBLOQUEIO = 8061 'Parametro objTipo.iCodigo
'Erro na Exclusão do Tipo %i
Public Const ERRO_LEITURA_TRIBNFISCAL = 8065
'Erro na leitura da tabela de Tributação de Nota Fiscal.
Public Const ERRO_LEITURA_TRIBCOMPLNFISCAL = 8066
'Erro na leitura da tabela de Tributação de Complemento de Nota Fiscal.
Public Const ERRO_LEITURA_TRIBITEMNFISCAL = 8067
'Erro na leitura da tabela de Tributação de Itens de Nota Fiscal.
Public Const ERRO_INCONSISTENCIA_TRIBITEMNFISCAL = 8068
'Tributação para o Item não confere com o existente.
Public Const ERRO_IR_FONTE_MAIOR_VALOR_TOTAL = 8069 'dValorIRRF, dValorTotal
'IR Fonte = %d não pode ultrapassar Valor Total = %d.
Public Const ERRO_ATUALIZACAO_PREVVENDA = 8071 'Parâmetro sCodigo
'Erro na atualização da Previsão de Venda com o código %s na tabela PrevVenda .
Public Const ERRO_INSERCAO_PREVVENDA = 8072 'Parâmetro sCodigo
'Erro na inserção da Previsão de Venda com o código %s na tabela PrevVenda .
Public Const ERRO_EXCLUSAO_PREVVENDA = 8073 'Parâmetro sCodigo
'Erro na exclusão da Previsão de Venda com o código %s na tabela PrevVenda .
Public Const ERRO_PREVVENDA_NAO_CADASTRADA = 8074 'Parâmetro sCodigo
'Previsão de Venda %s não cadastrada na tabela PrevVenda .
Public Const ERRO_LOCK_PREVVENDA = 8075 'Parâmetro sCodigo
'Não foi possível fazer o Lock da Previsão de Venda %s da tabela PrevVenda .
Public Const ERRO_CODIGO_PREVVENDA_NAO_PREENCHIDO = 8077
'O Código da Previsão de Venda não está preenchido.
Public Const ERRO_DATAPREVISAO_NAO_PREENCHIDA = 8078
'O campo Data de Previsão não está preenchido .
Public Const ERRO_PEDIDODEVENDA_INEXISTENTE = 8082 'Parametros = lCodPedido
'O Pedido de Venda %l não está cadastrado.
Public Const ERRO_LOCKEXCLUSIVE_TIPODEBLOQUEIO = 8083 'Sem Parâmetros
'Não foi possível fazer LockExclusive na tabela TiposDeBloqueio
Public Const ERRO_TIPODEBLOQUEIO_USADO = 8084  'Parâmetro, o código do Tipo de Bloqueio
'Não é permitido excluir o Tipo de Bloqueio %i, pois ele é utilizado em Bloqueios de Pedido de Venda. Pedido de Venda = %l, Filial Empresa = %i e Sequencial = %l.
Public Const ERRO_LEITURA_BLOQUEIOSPV_BLOQUEIOSPVBAIXADOS = 8086 'Sem Parâmetros
'Erro de Leitura na tabela BloqueiosPV ou BloqueiosPVBaixados
Public Const ERRO_LEITURA_PEDIDOSDEVENDA_PEDIDOSDEVENDABAIXADOS = 8087 'Sem Parâmetros
'Erro de Leitura na Tabela PedidosDeVenda ou PedidosDeVendaBaixados
Public Const ERRO_LEITURA_NFISCAL_NFISCALBAIXADAS = 8088 'Sem parâmtros
'Erro de Leitura na Tabela NFiscal ou NFiscalBaixadas
Public Const ERRO_LOCKEXCLUSIVE_CANALVENDA = 8089 'Sem Parâmetros
'Não foi possível fazer LockExclusive na Tabel CanalVenda
Public Const ERRO_CANAL_EM_PV = 8090 'Parâmetro Código do Canal de Venda
'Não é permitido excluir o Canal de Venda %i, pois ele é utilizado no Pedido de Venda %l da Filial Empresa %i.
Public Const ERRO_CANAL_EM_NF = 8091 'Parâmetro Código do Canal de Venda
'Não é permitido excluir o Canal de Venda %i, pois ele está sendo utilizado na Nota Fiscal com Número %l, Série %s da Filial Empresa %i.
Public Const ERRO_LIBERACAOCREDITO_INEXISTENTE = 8093 'Parametros = sCodUsuario
'O usuário %s não tem autorização para liberar por crédito este bloqueio.
Public Const ERRO_LIBERACAOCREDITO_LIMITEOPERACAO = 8094 'Parametros = sCodUsuario
'O usuário %s não tem autorização para liberar por crédito este bloqueio (Limite por Operação Excedido).
Public Const ERRO_LIBERACAOCREDITO_LIMITEMENSAL = 8095 'Parametros = sCodUsuario
'O usuário %s não tem autorização para liberar por crédito este bloqueio (Limite Mensal Excedido).
Public Const ERRO_ALTERACAO_CLIENTES = 8096 'Parametro lCodigo
'Erro na alteracao da tabela de Clientes. Código do Cliente = %l.
Public Const ERRO_LEITURA_LIBERACAOCREDITO = 8097 'Parametro: sCodUsuario
'Ocorreu um erro na leitura de um registro da tabela de Liberação de Crédito. Usuário = %s.
Public Const ERRO_LEITURA_VALORLIBERADOCREDITO = 8098 'Parametro: sCodUsuario, iAno
'Ocorreu um erro na leitura de um registro da tabela ValorLiberadoCredito. Usuário = %s, Ano = %i.
Public Const ERRO_EXCLUSAO_VALORLIBERADOCREDITO = 8099
'Erro na tentativa de excluir registro da tabela ValorLiberadoCredito.
Public Const ERRO_LOCK_LIBERACAOCREDITO = 8100
'Erro na tentativa de fazer 'lock' na tabela LiberacaoCredito.
Public Const ERRO_LIMITES = 8101
'Limite de Operação precisa ser menor ou igual ao Limite Mensal.
Public Const ERRO_ATUALIZACAO_LIBERACAOCREDITO = 8102
'Erro na tentativa de atualizar a tabela LiberacaoCredito.
Public Const ERRO_INSERCAO_LIBERACAOCREDITO = 8103
'Erro na tentiva de inserir um registro na tabela LiberacaoCredito.
Public Const ERRO_EXCLUSAO_LIBERACAOCREDITO = 8104 'Parametro codigo usuario
'Erro na tentativa de exclusão da alçada do usuário %s da tabela LiberacaoCredito.
Public Const ERRO_VINCULADO_PEDIDOSVENDA = 8105 'Parametro codigo usuario.
'A alçada não pode ser excluída pois o usuário %s está vinculado a um Pedido de Venda.
Public Const ERRO_VINCULADO_VALORLIBERADOCREDITO = 8106 'Parametro codigo usuario
'A alçada não pode ser excluída pois o usuário %s está vinculado a um registro na tabela de Liberação de Créditos.
Public Const ERRO_LIBERACAOCREDITO_VAZIA = 8109
'Erro tabela LiberacaoCredito está Vazia.
Public Const ERRO_LEITURA_TABELA_ALMOXARIFADOS = 8113 'Sem parametros
'Erro de leitura na Tabela Almoxarifados
Public Const ERRO_UNIDADE_NAO_CADASTRADA = 8114 'Sem parametros
'Unidade de medida não cadastrada
Public Const ERRO_TOTAL_RESERVADO_SEM_PREENCHIMENTO = 8116 'Sem parametors
'Total Reservado não foi informado
Public Const ERRO_NENHUM_ALMOXARIFADO_BD = 8117 'Sem parametros
'Nemhum almoxarifado está cadastrado no Banco de Dados
Public Const ERRO_ITEM_NAO_EXISTE = 8118 'Parametro sCodItem
'Item com código %s não existe
Public Const ERRO_PRODUTO_NAO_DISPONIVEL = 8119 'Parametro objItemPedido.sProduto
'Não existe disponibilidade do produto %s
Public Const ERRO_QUANTRESERVADA_MAIOR = 8121 'Parametro dQuantReservada
'Quantidade reservada não pode ser maior do que a quantidade disponível!
Public Const ERRO_TOTAL_RESERVADO_MAIOR = 8122 'Parametro dTotalReservado
'O Total reservado não pode ser maior que o total a reservar!
Public Const ERRO_TABELA_NAO_MARCADA = 8123 'Sem parâmetros
'Deve ser selecionada pelo menos uma Tabela para atualização.
Public Const ERRO_TABELAPRECO_INEXISTENTE1 = 8125 'Sem Parametro
'Nenhuma Tabela de Preço Cadastrada no Banco de Dados
Public Const ERRO_PERCENTUAL_NAOPREENCHIDO = 8126 'Sem parametro
'O valor do percentual de reajuste deve estar preenchido.
Public Const TELA_AUTCRED_CHAMADA_SEM_PARAMETRO = 8127 'Sem Parametro
'A tela Autorização de Crédito foi chamada sem a passagem do parametro necessário.
Public Const ERRO_EMPRESA_INVALIDA = 8128
'Utilização Empresa inválida
Public Const ERRO_BLOQUEIO_LIBERACAO_NAO_MARCADO = 8130 'Sem parâmetro
'Pelo menos um Bloqueio deve estar marcado para Liberação.
Public Const ERRO_BLOQUEIO_EDICAO_NAO_MARCADO = 8131 'Sem parâmetro
'Deve haver um Bloqueio marcado para Edição.
Public Const ERRO_TIPOBLOQUEIO_NAO_MARCADO = 8132 'Sem parâmetro
'Pelo menos 1 Tipo de Bloqueio deve estar marcado para esta operação.
Public Const ERRO_AUSENCIA_ALMOXARIFADO_FILIAL = 8134 'iFilialEmpresa
'Nunhum almoxarifado da Filial %i está cadastrado no Banco de Dados.
Public Const ERRO_ITEM_INEXISTENTE = 8135
'Ítem não existe
Public Const ERRO_QUANT_ALOCADA_MAIOR_DISPONIVEL = 8136
'Quantidade Alocada não pode ser maior do que a Quantidade Disponível
Public Const ERRO_TOTAL_ALOCACAO_SUPERIOR_ALOCAR = 8137
'Total Alocado não pode ser maior do que a Quantidade a Alocar.
Public Const ERRO_TIPODOC_DIFERENTE_NF_SAIDA_DEVOLUCAO = 8159 'iTipoNFiscal
'Tipo de Documento %i não é Nota Fiscal de Saida de Devolução.
Public Const ERRO_NF_EXTERNA_NAO_CADASTRADA = 8161 'Parâmetros: sSerie, lNumNotaFiscal, lFornecedor, iFilial
'Nota Fiscal Externa com série %s, número %l, fornecedor = %l, Cliente = %l e Filial = %i não está cadastrada no Banco de Dados.
Public Const ERRO_DATASAIDA_NAO_PREENCHIDA = 8162
'A Data de Saída não foi preenchida.
Public Const ERRO_LEITURA_NOTA_FISCAL = 8163 'Parâmetros: iTipoNFiscal, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscal na Nota Fiscal com Tipo = %i, Serie = %s e Número = %l.
Public Const ERRO_ALTERACAO_NFISCAL_SAIDA2 = 8164 'Parâmetros: iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal com os dados Tipo = %i, Serie = %s, Número NF = %l, Data Emissao = %dt está cadastrada no Banco de Dados. Não é possível alterar.
Public Const ERRO_PEDIDO_VENDA_BLOQUEIO_CREDITO = 8165 'Parametro: lCodPedido
'Não é possível gerar Nota Fiscal a partir do pedido %l pois ele possui um Bloqueio de Crédito.
Public Const ERRO_QUANTIDADE_FATURAR_MENOR = 8166 'Parametro: dQuantFaturar
'A quantidade do item não pode ultrapassar a quantidade a ser faturada que é %d.
Public Const ERRO_PEDIDO_VENDA_BLOQUEIO_TOTAL = 8167 'Parametro: lCodPedido
'Não é possível gerar Nota Fiscal a partir do pedido %l pois ele possui um Bloqueio Total.
Public Const ERRO_ALTERACAO_NFISCALPEDIDO = 8168
'Não é possível alterar nota fiscal gerada a partir de pedido.
Public Const ERRO_CODPEDIDO_NAO_INFORMADO = 8169
'O Código do Pedido deve ser informado.
Public Const ERRO_FILIALPEDIDO_NAO_INFORMADA = 8170
'A Filial do Pedido deve ser informada.
Public Const ERRO_PEDIDOVENDA_BAIXADO = 8171 'Parâmetros:lCodPedido, iFilialPedido
'O Pedido de Venda com o código %l da Filial Empresa %i já está baixado
Public Const ERRO_PEDIDO_VENDA_NAO_CADASTRADO1 = 8172 'Parâmetros:lCodPedido, iFilialPedido
'O Pedido de Venda com o código %l da Filial Empresa %i não está cadastrado no Banco de Dados.
Public Const ERRO_FILIALFATURAMENTO_DIFERENTE = 8173 'Parametros: iFilialFaturamento,lCodPedido, iFilialEmpresa
'A Filial de Faturamento %i do Pedido de Venda %l, é diferente da Filial empresa Atual %i.
Public Const ERRO_NENHUM_PEDIDO_TRAZIDO = 8174
'Nenhum pedido de venda foi trazido para a tela.
Public Const ERRO_PEDIDO_VENDA_FATURA_INTEGRAL = 8175
'O Pedido de Venda deve ser faturado integralmente.
Public Const ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO = 8176
'Não pode haver substituicao de produto.
Public Const ERRO_CANALVENDA_NAO_ENCONTRADO = 8177 '%sCanalVenda
'O Canal de Venda %s não foi encontrado.
Public Const ERRO_TIPODOC_DIFERENTE_NF_SAIDA_REMESSA = 8178 'Parâmetro: iTipoNFiscal
'Tipo de Documento %i não é Nota Fiscal de Saida de Remessa.
Public Const ERRO_MODIFICACAO_SERIE = 8179 'Sem parametro
'Erro na modificação da Serie.
Public Const ERRO_INSERCAO_SERIE = 8180 'Sem parametro
'Erro na inserção da Serie.
Public Const ERRO_EXCLUSAO_SERIE = 8181 'Parametro: sSerie
'Erro na tentativa de excluir Serie.
Public Const ERRO_PROXNUMNFISCAL_NAO_PREENCHIDA = 8182
'O Proximo Número para Nota Fiscal deve ser preenchido.
Public Const ERRO_SERIE_REL_NF_REC = 8183 'Parametro: sSerie
'Erro na exclusão da Serie %s, relacionado com Nota Fiscal a Receber.
Public Const ERRO_SERIE_REL_NF_REC_BAIXADA = 8184 'Parametro: sSerie
'Erro na exclusão da Serie %s, relacionado com Nota Fiscal a Receber Baixada.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_SERIE = 8185
'Erro na leitura da tabela de Séries de Notas Fiscais.
Public Const ERRO_EXISTEM_NOTAS_FISCAIS_SERIE = 8186
'A série não pode ser excluída, pois está vinculada com a Nota Fiscal %l da Filial Empresa %i.
Public Const ERRO_NAO_EXISTE_RESERVAS1 = 8187 ' Parametros lCodNFiscal, sCodProduto
'Não existe Reservas para a NFiscal  %l do Produto %s.
Public Const ERRO_PRODUTO_SEM_SUBSTITUTOS = 8189 'Sem parametros
'Produto não tem Produtos Substitutos associados
Public Const ERRO_PRODUTOS_SUBSTITUTOS_INVALIDOS1 = 8190 'Sem parametros
'Produtos substitutos já fazem parte da Nota Fiscal ou estão Inativos ou não participam do Faturamento
Public Const ERRO_PEDIDO_EDICAO_NAO_MARCADA = 8195 'Sem parâmetros
'Deve haver uma Pedido marcado para Edição.
Public Const ERRO_DATA_FATURA_NAO_PREENCHIDA = 8197 'Sem parametros
'A data da fatura não foi preenchida.
Public Const ERRO_FATURA_SEM_NOTASFISCAIS = 8198 'Sêm Parâmetros
'Não existem Parcelas no Grid de Parcelas Para faturar.
Public Const ERRO_NOTASFISCAIS_NAO_SELECIONADAS = 8199 'Sem Parâmetros
'Não existe nenhuma Nota Fiscal selecionada no Grid.
Public Const ERRO_TIPODEBLOQUEIO_EXCLUSAO = 8200 'Parâmetro: iCodigo
'O Tipo De Bloqueio %i não pode ser excluído.
Public Const ERRO_PEDIDO_FATURAMENTO_PARCIAL = 8208 'Parametro CodigoPV
'O pedido %l não pode ser faturado totalmente. Use a transação de faturamento de nota fiscal a partir de um pedido para faturar este pedido.
Public Const ERRO_PEDIDO_JA_FATURADO_PARCIALMENTE = 8209 'Parametro CodigoPV
'O pedido %l já foi faturado parcialmente. Use a transação de faturamento de nota fiscal a partir de um pedido para faturar este pedido.
Public Const ERRO_PEDIDO_SEM_COBRANCA = 8210 'Parametro: lCodPedido
'Foi solicitada a geração de uma nota fiscal fatura para o Pedido %l que não possui informações de cobrança.
Public Const ERRO_INSERCAO_VALORLIBERADOCREDITO = 8211 'Parametros: sCodUsuario, iAno
'Ocorreu um erro na inserção de um registro na tabela ValorLiberadoCredito. Usuário = %s, Ano = %i.
Public Const ERRO_ATUALIZACAO_VALORLIBERADOCREDITO = 8212 'Parametros: sCodUsuario, iAno
'Ocorreu um erro na atualização de um registro da tabela ValorLiberadoCredito. Usuário = %s, Ano = %i.
Public Const ERRO_CLIENTE_SEM_CREDITO = 8213 'Parametros: sCodUsuario
'O cliente %s não possui crédito para efetuar esta operação.
Public Const ERRO_LEITURA_NFISCALBAIXADA3 = 8214 'Parâmetros: iTipoNFiscal, lCliente, iFilialCli, sSerie, lNumNotaFiscal
'Ocorreu um erro na leitura da tabela NFiscalBaixadas na Nota Fiscal com Tipo = %i, Cliente = %l, Filial = %i, Serie = %s e Número = %l.
Public Const ERRO_EXCLUSAO_COMISSOESPEDVENDAS1 = 8215 'Parametro: iFilialEmpresa, lPedidoVenda
'Erro na tentativa de excluir registro da tabela de ComissoesPedVendas da Filial %i e Pedido %l.
Public Const ERRO_PRODUTOS_SUBSTITUTOS_INVALIDOS = 8218 'Sem parametros
'Produtos substitutos já fazem parte do pedido de venda ou estão Inativos ou não participam do Faturamento
Public Const ERRO_VARIAS_NFISCAL_EDICAO = 8219 'Sem parâmetro
'Só deve haver uma Nota Fiscal marcada para Edição.
Public Const ERRO_FILIALDE_MAIOR_FILIALATE = 8220 'Sem parâmetros
'Cliente De não pode ser maior do que o Cliente Até.
Public Const ERRO_TIPO_BLOQUEIO_PRE_DEFINIDO = 8221 'sem parametros
'Um tipo de bloqueio pré-definido não pode ser alterado
Public Const ERRO_PEDIDO_VENDA_INICIAL_MAIOR = 8225
'O Pedido de Venda Inicial é maior que o Final.
Public Const ERRO_VENDEDOR_NAO_CADASTRADO2 = 8226
'O Vendedor não está cadastrado.'
Public Const ERRO_TRANSPORTADORA_INICIAL_MAIOR = 8227
'A Transportadora Inicial é maior do que a Final.'
Public Const ERRO_PERIODO_PREENCHIDO_INCORRETO = 8228
'Os Periodos devem estar preenchidos em ordem, sem espaços entre eles.
Public Const ERRO_PERIODO_ANTERIOR_MENOR = 8229
'O Periodo anterior não pode ser maior que o seguinte.
Public Const ERRO_INSERCAO_RELFATPRAZOPAG = 8230 'Parâmetro: lCodigo
'Erro na Inserção do Relatorio com o Código %l
Public Const ERRO_NUM_MAXIMO_BLOQUEIO_MAIOR_LIMITE = 8231 'Sem parâmetros
'O número de Bloqueios ultrapassa ao limite máximo que é de 1000.
Public Const ERRO_NATUREZAOP_SAIDA = 8232 'parametro: natop
'Informe uma natureza de operaÇÃo de saída (>= 500).
Public Const ERRO_NATUREZAOP_ENTRADA = 8233 'parametro: natop
'Informe uma natureza de operacao de entrada (<500)
Public Const ERRO_NATUREZAOP_ITEM_TRIBUTACAO_NAO_PREENCHIDA = 8234 'iItem
'A Natureza de Operação da tributação do item %i não foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_ITEM_NAO_PREENCHIDO = 8235 'iItem
'O Tipo de Tributacao do item %i não foi preenchido.
Public Const ERRO_NATUREZAOP_DESCONTO_NAO_PRENCHIDA = 8236
'A Natureza de Operação da tributação do Valor Desconto não foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_DESCONTO_NAO_PREENCHIDO = 8237
'O Tipo de Tributação do Valor Desconto não foi preenchido.
Public Const ERRO_NATUREZAOP_FRETE_NAO_PRENCHIDA = 8238
'A Natureza de Operação da tributação do Valor Frete não foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_FRETE_NAO_PREENCHIDO = 8239
'O Tipo de Tributação do Valor Frete não foi preenchido.
Public Const ERRO_NATUREZAOP_DESPESAS_NAO_PRENCHIDA = 8240
'A Natureza de Operação da tributação do Valor Despesas não foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_DESPESAS_NAO_PREENCHIDO = 8241
'O Tipo de Tributação do Valor Despesas não foi preenchido.
Public Const ERRO_NATUREZAOP_SEGURO_NAO_PRENCHIDA = 8242
'A Natureza de Operação da tributação do Valor Seguro não foi preenchida.
Public Const ERRO_TIPO_TRIBUTACAO_SEGURO_NAO_PREENCHIDO = 8243
'O Tipo de Tributação do Valor Seguro não foi preenchido.
Public Const ERRO_VARIOS_PEDIDO_EDICAO = 8244 'Sem parâmetro
'Só deve haver um Pedido marcado para Edição.
Public Const ERRO_NOTAFISCAL_NAO_EDITAVEL = 8245 'Sem parâmetros
'A Nota Fiscal não pode ser editável.
Public Const ERRO_PEDIDOVENDA_NAO_INFORMADO = 8246 'Sem parâmetros
'O número do Pedido de Venda deve ser preenchido.
Public Const ERRO_NFISCAL_NAO_INFORMADA = 8247 'Sem parâmetros
'O número e a série da Nota Fiscal devem estar preenchidos.
Public Const ERRO_LEITURA_NFISCAL5 = 8248 'Parâmetros: iFilialEmpresa, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscal com Filial %i, Série %s e Número da Nota %l.
Public Const ERRO_SERIE_MAIOR_LIMITE_MAXIMO = 8249 'Sem parâmetros
'Ultrapassou o limite máximo do tamanho da Série.
Public Const ERRO_VALORBASE_MAIOR_VALORDOC = 8250 'Parâmetros: dValorBase, dValorDoc
'O Valor Base não pode ser maior que o Valor do Documento.
Public Const ERRO_DATA_DESCONTO_PARCELA_SUPERIOR_DATA_VENCIMENTO = 8251 'Parâmetros: dtDataDesconto, iDesconto, iParcela
'A data %dt do desconto %i da parcela %i é superior a data de vencimento.
Public Const ERRO_DATAINICIAL_NAOPREENCHIDA = 8252 'Sem Parâmetros
'Erro a Data Inicial não foi preenchida.
Public Const ERRO_DATAFINAL_NAOPREENCHIDA = 8253 'Sem Parâmetros
'Erro a Data Final não foi preenchida.
Public Const ERRO_FILIALENTREGA_NAO_INFORMADA = 8255
'A Filial de Entrega do Cliente não foi preenchida.
Public Const ERRO_NATUREZAOP_INICIAL_MAIOR = 8256
'A Natureza de Operação Inicial é maior que a Final.
Public Const ERRO_PEDIDO_VENDA_BLOQUEIO = 8257 'Parâmetro: lCodPedido
'Não é possível gerar Nota Fiscal a partir do pedido %l pois ele possui bloqueios que impedem o seu faturamento.
Public Const ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2 = 8258 'Parametro: sProduto
'O Produto %s não pode ser vendido.
Public Const ERRO_LOCK_COMISSOESNF = 8259 'Sem Parametros
'Ocorreu um erro na tentativa de fazer um lock de um registro da tabela de Comissões de Notas Fiscais.
Public Const ERRO_EXCLUSAO_COMISSOESNF2 = 8260 'Sem Parametros
'Erro na tentativa de excluir registro da tabela de Comissões de Notas Fiscais.
Public Const ERRO_EXCLUSAO_RESERVA1 = 8261 'Sem parâmetros
'Erro na tentativa de excluir registro da tabela de Reservas.
Public Const ERRO_ATUALIZACAO_RESERVA1 = 8262 'Sem Parâmetros
'Erro na atualização de resgitsro na tabela de reservas.
Public Const ERRO_LEITURA_USUARIOFILIALEMPRESA_DIC = 8263 'Sem Parametros
'Ocorreu um erro na leitura da Query UsuarioFilialEmpresa do Dicionario de Dados.
Public Const ERRO_LEITURA_PEDIDOS_VENDA_BAIXA_PV = 8264 'Sem parametros
'Erro na leitura de Pedidos de Venda para baixa.
Public Const ERRO_NOTA_FISCAL_NAO_TEM_COMISSAO = 8265 'Parametros: lNumeroNFiscal, sSerie
'A Nota Fiscal %l da Série %s não pode ter comissão ou não está cadastrada.
Public Const ERRO_NENHUM_PEDIDO_SELECIONADO = 8267
'Para editar um Pedido é necessário que uma das linhas do grid seja selecionada.
Public Const ERRO_FILIALPEDIDO_DIFERENTE_FILIALEMPRESA = 8268 'Parâmetros: lCodPedido, iFilialEmpresaPedido, giFilialEmpresa
'Não é possível editar o Pedido de Venda %l da Filial Empresa %i pois estamos na Filial Empresa %i.
Public Const ERRO_PRODUTO_SEM_ALMOX_PADRAO1 = 8269 'sProduto, giFilialEmpresa
'O Produto %s não está relacionado com nenhum Almoxarifado da Filial Empresa %i.
Public Const ERRO_QUANTIDADE_PREVVENDA_NAO_PREENCHIDA = 8270 'Sem Parametro
'A Quantidade da previsão de Venda não foi preenchida.
Public Const ERRO_CODIGO_DESABILITADO_IMUTAVEL = 8271
'Não é possível alterar código do Pedido de Vendas.
Public Const ERRO_FILIALENTREGA_FORNECEDOR_NAO_INFORMADA = 8272
'A Filial de Entrega do Fornecedor não foi preenchida.
Public Const ERRO_CONDICAO_PAGTO_ALTERADA_NUM_PARC = 8273 'parametro: codigo da condicao
'A condição de pagamento %s teve alterado o seu número de parcelas.
Public Const ERRO_LEITURA_NFISCALTIPODOCINFO = 8274 'Sem parâmetros
'Erro na leitura da tabela NfiscalTipoDocInfo.
Public Const ERRO_QUANT_FATURADA_MAIOR_QUANT_A_FATURAR = 8275 'parametro lNumIntItemPV
'Quantidade faturada é maior do que quantidade a faturar de ítemPV com número interno %l.
Public Const ERRO_TIPODOC_DIFERENTE_NF_FATURA_PEDIDO = 8276 'Parâmetro: iTipoDocInfo
'Tipo de Documento %i não é Nota Fiscal Fatura Pedido.
Public Const ERRO_TIPODOC_DIFERENTE_NF_PEDIDO = 8277 'Parâmetro: iTipoDocInfo
'Tipo de Documento %i não é Nota Fiscal Pedido.
Public Const ERRO_PRODUTO_SEM_TABELA_PADRAO = 8278 'Parametro: sCodProduto
'Produto com código %s não tem Tabela de Preço Padrão associada.
Public Const ERRO_TABELAPRECOITEM_INEXISTENTE3 = 8279 'Parametros: sCodProduto, dtDataFinal
'Inexiste no Banco de Dados Item de Tabela de Preço Padrão para Produto %s com Data de Vigência anterior ou igual a %dt.
Public Const ERRO_VALOR_EMISSAO_GRID_NAO_PREENCHIDO = 8280 'Parametro: iLinha
'O valor da comissão na emissão na linha %i do Grid de Comissões não foi preenchido.
Public Const ERRO_PERCENTAGEM_EMISSAO_GRID_NAO_PREENCHIDA = 8281
'A porcentagem da comissão na emissão na linha %i do Grid de Comissões não foi preenchida.
Public Const ERRO_PERCENTUAL_COMISSAO_NULO = 8282
'A porcentagem de comissão não pode ser nula.
Public Const ERRO_LIBERACAOCREDITO_LIMITEOPERACAO1 = 8283 'Parametros: lPedido, sCodUsuario
'O Pedido %l não teve o crédito liberado pois o usuário %s ultrapassou o limite por operação.
Public Const ERRO_LIBERACAOCREDITO_INEXISTENTE1 = 8284 'Parametros: sCodUsuario, lPedido
'O usuário %s não tem autorização para liberar por crédito o bloqueio do Pedido %l.
Public Const ERRO_LIBERACAOCREDITO_LIMITEMENSAL1 = 8285 'Parametros: lPedido, sCodUsuario
'O Pedido %l não teve o crédito liberado, pois o usuário %s ultrapassou o limite mensal.
Public Const ERRO_LEITURA_ITEMPV = 8286 'Parametro: lNumIntDoc
'Erro na leitura da tabela ItensPedidoDeVenda, registro com número interno %l.
Public Const ERRO_LOCK_ITEMPEDIDODEVENDA = 8287 'Parametro: lNumIntDoc
'Erro de lock na tabela ItensPedidoDeVenda, registro com número interno %l.
Public Const ERRO_QUANT_LIBERADA_MAIOR_QUANT_RESERVADA = 8288 'Parametro: lNumIntDoc
'Quantidade de reserva liberada de ItemPV com número interno %l superior a quantidade reservada do ítem.
Public Const ERRO_ATUALIZACAO_ITEMPV = 8289 'Parametro: lNumIntDoc
'Erro na atualização de ItemPedidoDeVenda com número interno %l.
Public Const ERRO_NUMERO_DE_NAO_PREENCHIDO = 8291 'Sem Parametros
'O número de Nota Fiscal De não foi preenchido.
Public Const ERRO_NUMERO_ATE_NAO_PREENCHIDO = 8292 'Sem Parametros
'O número de Nota Fiscal Até não foi preenchido.
Public Const ERRO_LOCK_SERIE_IMPRESSAO_NF = 8293 'Parametros : sSérie
'A Série %s está lockada para impressão.
Public Const ERRO_LEITURA_TRIBUTACAOPVBAIXADO = 8294
'Erro na leitura da tabela de TributacaoPVBaixado.
Public Const ERRO_COMPL_PV_BAIXADO_TIPO_INVALIDO = 8295 'sem parametros
'registro na tabela TributacaoComplPVBaixado com tipo inválido
Public Const ERRO_SERIE_SEM_PADRAO = 8298 'Parametro: sSerie
'Para desmarcar a Serie %s como Padrão marque outra Série como Padrão.
Public Const ERRO_TIPO_NOTA_FISCAL_NAO_FATURA_PEDIDO = 8299 'Parametro: iCodigo
'Código %i não corresponde a Nota Fiscal Fatura de Venda a partir de Pedido.
Public Const ERRO_NFISCALFATURA_SEM_TITULO_RECEBER = 8300 'Parametro: lNumNotaFiscal
'Nota Fiscal Fatura com número %l não tem Título a Receber associado.
Public Const ERRO_TITULO_RECEBER_NAO_CADASTRADO = 8301 'Parametro: lNumIntDoc
'Título Receber com número interno %l não está cadastrado.
Public Const ERRO_FATURAPAG_NAO_EXCLUIDA = 8302
'A nota fiscal não pode ser cancelada/excluída por estar vinculada a uma fatura a pagar.
'Para realizar o cancelamento da nota fiscal é preciso, antes, excluir a Fatura a Pagar.
Public Const ERRO_FATURAREC_NAO_EXCLUIDA = 8303 'Parâmetro : lNumFatura
'A nota fiscal não pode ser cancelada/excluída por estar vinculada a fatura a Receber %l.
'Para realizar o cancelamento da nota fiscal é preciso, antes, excluir a Fatura a Receber.
Public Const ERRO_COMISSAO_BAIXADA_CANC_NFISCAL = 8304 'sem parametros
'Não pode cancelar uma nota fiscal que teve a comissao já baixada(paga).
Public Const ERRO_EXCLUSAO_FILIALCLIENTEFILEMP = 8305 'Sem Parametros
'Erro na Exclusão da tabela FilialClienteFilEmp.
Public Const ERRO_FATURADE_NAO_PREENCHIDA = 8306 'Sem Parametros
'É obrigatório o preenchimento do campo Fatura De.
Public Const ERRO_FATURAATE_NAO_PREENCHIDA = 8307 'Sem Parametros
'É obrigatório o preenchimento do campo Fatura Até.
Public Const ERRO_TIPOPRODUTO_INICIAL_MAIOR = 8308 'Sem parâmetros
'O tipo de produto inicial não pode ser maior que o tipo de produto final.
Public Const ERRO_INCLUSAO_NFISCAL_NUMAUTO = 8309 'sem parametros
'Para criar uma nota fiscal o numero da mesma tem que estar em branco
Public Const ERRO_NF_NAO_CADASTRADA2 = 8310 'Parâmetros: lNumNotaFiscal, sSerie, iTipoNFiscal, lCliente, iFilialCli, dtDataEmissao
'A nota fiscal com os dados abaixo não está cadastrada. Número: %s, Série: %s, Tipo: %s, Cliente: %s, Filial: %s e Emissão em: %s.
Public Const ERRO_NF_NAO_CADASTRADA4 = 8311 'Parâmetros: lNumNotaFiscal, sSerie, iTipoNFiscal, lFornecedor, iFilialForn, dtDataEmissao
'A nota fiscal com os dados abaixo não está cadastrada. Número: %s, Série: %s, Tipo: %s, Fornecedor: %s, Filial: %s e Emissão em: %s.
Public Const ERRO_VALORICMS_MAIOR_TOTAL = 8312 'Parametros: dValorICMS , dValorTotal
'O Valor do Imposto %s não pode ser maior que o Valor Total  %s.
Public Const ERRO_DESTINATARIO_NAO_PREENCHIDO = 8313
' O Preenchimento do Destinatário é obrigatório
Public Const ERRO_BASECALCULO_NAO_PREENCHIDA = 8314
'O Preenchimento da base de cálculo é obrigatório.
Public Const ERRO_VALORFRETE_NAO_PREENCHIDO = 8315
'O Preenchimento do Valor do Frete é obrigatório.
Public Const ERRO_LEITURA_CONHECIMENTO_FRETE = 8316
'Erro de leitura na tabela de Conhecimento de Frete.
Public Const ERRO_CONHECIMENTOFRETE_NAO_CADASTRADO = 8317 'Parametros: sSerie, lNumNotaFiscal
'O Conhecimento de Transporte com a série %s e Número %s não está cadastrado.
Public Const ERRO_REMETENTE_NAO_PREENCHIDO = 8318
'O Preenchimento do Remetente é obrigatório.
Public Const ERRO_VALORBASE_MENOR_SUBTOTAL = 8319 'Parametros: dValorBase , dValorTotal
'O Valor Base %s não pode ser menos que a soma dos valores que é %s.
Public Const ERRO_INCLUSAO_NFISCAL = 8320 'Parâmetros: sSerie, lNumNotaFuscal
'Erro na tentativa de incluir a Nota Fiscal  Serie %s Número %s.
Public Const ERRO_INCLUSAO_CONHECIMENTOFRETE = 8321 'Parâmetros: sSerie, lNumNotaFuscal
'Erro na tentativa de incluir a Nota Fiscal  Serie %s Número %s.




'Códigos de Avisos - Reservado de 5900 até 5999
Public Const AVISO_EXCLUSAO_NATUREZAOP = 5900 'Parametro: sCodigo
'Confirma a exclusão da Natureza de Operação %s ?
Public Const AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS = 5901 'Sem Parametros
'Não é possível selecionar "Todos" para produtos e clientes simultaneamente.
Public Const AVISO_CONFIRMA_EXCLUSAO_CANAL = 5902 'Parametro objCanal.iCodigo
'Confirma exclusão do Canal %i ?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPODEBLOQUEIO = 5903 'Parametro objTipo.iCodigo
'Confirma exclusão do Tipo de Bloqueio %i ?
Public Const AVISO_EXCLUSAO_PREVVENDA = 5904 'Parametro sCodigo
'Confirma exclusão da Previsão %s do Banco de Dados ?
Public Const AVISO_NUM_MAX_BLOQUEIOS_LIBERACAO = 5905
'O número máximo de Bloqueios possíveis de exibição foi atingido . Ainda existem mais Bloqueios para Liberação .
Public Const AVISO_CONFIRMA_EXCLUSAO_ALCADAFAT_USUARIO = 5906 'Parâmetro: Código do Usuário
'Confirma a exclusão da alçada do usuário com código %s?
Public Const AVISO_ITEM_ANTERIOR_ALTERADO = 5907 'Sem parametros
'As reservas do Item anterior foram alteradas!Deseja salvar alterações?
Public Const AVISO_CANCELAR_NFISCAL = 5908 'Parâmetro: lNumNotaFiscal
'Deseja realmente cancelar a Nota Fiscal de Saída %l ?.
Public Const AVISO_ALOCADO_MENOR_ALOCAR = 5909 'dTotalAlocado, dQuantAlocar
'A quantidade alocada %d é inferior à quantidade a alocar %d. Deseja prosseguir?
Public Const AVISO_NFISCAL_SAIDA_DEVOLUCAO_MESMO_NUMERO = 5910 'Parâmetros: sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados exite Nota Fiscal  com os Dados Série NF =%s, Número NF =%l, Data Emissão =%dt.
'Deseja prosseguir na inserção de Nota Fiscal com o mesmo número?
Public Const AVISO_PEDIDO_TELA_NAO_UTILIZADO = 5911 'Parametros: lCodPedido, iFilialPedido
'A Nota Fiscal da tela vai ser gerada com base nos dados do Pedido %l da Filial %i trazidos anteriormente para a tela.
'O codigo e filial que estão na tela serão ignorados. Deseja prosseguir na gravação da nota fiscal?
Public Const AVISO_CRIAR_CANALVENDA = 5912 'iCanalVenda
'O Canal Venda %i não existe. Deseja Criá-lo?
Public Const AVISO_NFISCAL_SAIDA_REMESSA_MESMO_NUMERO = 5913 'PARÂMETROS: sSerie, lNumero, dtDataEmissao
'No Banco de Dados exite Nota Fiscal  com os Dados Série NF =%s, Número NF =%l, Data Emissão =%dt.
'Deseja prosseguir na inserção de Nota Fiscal com o mesmo número?
Public Const AVISO_EXCLUIR_SERIE = 5914 'Parametro: sSerie
'Confirma exclusão da Série %s?
Public Const AVISO_EXISTE_NF_MAIOR_OU_IGUAL = 5915 ' Parametros: sSerie, lNumNotaFiscal, dtDataEmissao, iFilialEmpresa
'Existe Nota Fiscal com Número maior ou igual à Nota: Série = %s, Número = %l, Data de Emissão %dt e FilialEmpresa = %i. Confirma a Gravação ?
Public Const AVISO_NF_ULTIMA_GRAVADA = 5916 ' Parametros: sSerie, lNumNotaFiscal, dtDataEmissao, iFilialEmpresa
'A última Nota Fiscal gravada foi: Série = %s, Número = %l, Data de Emissao %dt e FilialEmpresa = %i.
Public Const AVISO_ALTERACAO_NFISCAL_SAIDA_CONTAB = 5918 'Parâmetros: iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal com os dados Tipo = %i, Serie = %s, Número NF = %l, Data Emissao = %dt está cadastrada no Banco de Dados, só poderá ser alterado os dados relativos a contabilidade. Deseja proseguir na alteração?
Public Const AVISO_NFISCAL_REIMPRESSA = 5919 'Parametros : lNumeroNFInicial, lNumNotaFinal
'Confirma a reimpressão das Notas de número %l até %l ?
Public Const AVISO_NFISCAL_LOCKADA = 5920 'Sem Parametros
'A Impressão da Nota Fiscal está bloqueada. Está havendo uma impressão ou houve erro anterior.
'Continue somente em caso de Erro anterior. Deseja Continuar?
Public Const AVISO_SERIE_GRAVADA_PADRAO = 5921 'Parametro: sSerie
'A Série %s será a Série padrão.



