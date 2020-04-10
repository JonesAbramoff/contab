Attribute VB_Name = "ErrosInpal"
Option Explicit

Public Const ERRO_DESPESAFINANCEIRA_NAO_PREENCHIDA = 500000 'Sem parâmetros
'Para Despesa Financeira específica, é obrigatório preencher seu valor.
Public Const ERRO_JUROS_NAO_PREENCHIDO = 500001 'Sem parâmetros
'Para Juros específico, é obrigatório preencher seu valor.
Public Const ERRO_PEDIDO_PROGRAMADO_COM_RESERVA = 500002 'Sem parâmetros
'Para um Pedido de Vendas programado, as Reservas não devem
'ser preenchidas.
Public Const ERRO_TRANSPORTADORA_NAO_PREENCHIDA = 500003 'Sem parâmetros
'Para frete por conta do destinatário é obrigatório o preenchimento
'da Transportadora.
Public Const ERRO_VALORBASE_ITEM_NAO_PREENCHIDO = 500004 'Parâmetros: iLinha
'Valor Base do Item %i do Grid Itens não foi preenchido.
Public Const ERRO_VALORORIGINALPARC_NAO_INFORMADO = 500005 'Parâmetros: iLinha
'O Valor Original da parcela %s não foi informado.
Public Const ERRO_MOTIVODIFERENCAPARC_NAO_INFORMADO = 500006 'Parâmetros: iLinha
'O Motivo de Diferença da parcela %s não foi informado.
Public Const ERRO_SOMA_VALORORIGINAL_DIFERENTE_VALORTOTAL = 500007 'Parâmetros: dSomaValorOriginal, dValorTotal
'A soma de valores originais das parcelas %s é diferente do valor total que é %s.
Public Const ERRO_MOTIVODIFERENCA_NAO_ENCONTRADO = 500008 'Parâmetros: sMotivo
'O motivo de diferença %s não está cadastrado.
Public Const ERRO_SOMA_PARCELAS_DIFERENTE_TOTAL = 500009 'Parâmetros: dValorParcelas, dValorTotal, dValorDiferenca
'A soma das parcelas %s é diferente do valor total do título %s mais o valor de diferença %s.
Public Const ERRO_MOTIVODIFERENCA_INFORMADO_ERRADO = 500010 'Parâmetros: iParcela
'Foi preenchido um Motivo de Diferença para parcela %s, não tem diferença de valor.
Public Const ERRO_PRECOUNITARIO_NAO_PREENCHIDO = 500011 'Sem parâmetros
'É obrigatório o preenchimento de Preço Unitário.
Public Const ERRO_VALOR_DIFERENTE_QUANTPRECOUNITARIO = 500012 'Parâmetros: dValor, dValorReal
'O valor %s é diferente da quantidade vezes o preço unitário, que é %s.
Public Const ERRO_LEITURA_MATERIAL = 500013 'Parametro: iCodigo
'Erro ocorrido na leitura do material %i.
Public Const ERRO_MATERIAL_MESMA_DESCRICAO = 500014 'Parametro :sDescricao
'Já existe material com a descrição %s.
Public Const ERRO_LOCK_MATERIAl = 500015 'Parametro: iCodigo
'Erro na tentativa de lock do material %i.
Public Const ERRO_ATUALIZACAO_MATERIAL = 500016 'Parametro: iCodigo
'Erro na tentativa de atualizar o material com código %i.
Public Const ERRO_INSERCAO_MATERIAL = 500017 'Parametro: iCodigo
'Erro na tentativa de inserir material com código %i.
Public Const ERRO_CODIGO_MATERIAL_ZERADO = 500018 'Sem Parametro
'O código do material não pode ser igual a zero.
Public Const ERRO_MATERIAL_NAO_ENCONTRADO = 500019 'Parametro: iCodigo
'O material com o código %i não está cadastrado
Public Const ERRO_EXCLUSAO_MATERIAL = 500020 'Parametro: iCodigo
'Erro ocorrido na exclusão do material %i.
Public Const ERRO_LEITURA_MATERIAL_COTACAOVENDA = 500021 'Parametro: iCodigo
'Erro ocorrido na exclusão de material
Public Const ERRO_MATERIAL_UTILIZADO_COTACAOVENDA = 500022 'Parametro: iCodigo
'Erro de material utilizado na tabela de cotação de vendas ocorrido na exclusão de material
'Não é possível excluir o material com o código %s, pois ele está sendo utilizado em um cotaçãio de transportadora
Public Const ERRO_MATERIAL_COTACAOVENDA_NAO_ENCONTRADO = 500023 'Parametro: iCodigo
'Erro de material utilizado não encontrado na tabela de cotação de vendas ocorrido na exclusão de material
'Public Const ERRO_SIGLA_NAO_PREENCHIDA = 500024
'Preenchimento da sigla é obrigatório.
Public Const ERRO_LEITURA_MOTIVOPERDA = 500025 'Parametro iCodigo
'Erro na leitura da tabela 'Motivo Perda' - Codigo = %s
Public Const ERRO_MOTIVOPERDA_NAO_CADASTRADO = 500026 'Parametro iCodigo
'Erro, Motivo Perda não cadastrado - Codigo = %s
Public Const ERRO_LEITURA_COTACAOVENDA = 500028 'Parametro iCodigo
'Erro, na leitura de Cotacao Venda - Codigo= %s
Public Const ERRO_MOTIVOPERDA_MESMA_SIGLA = 500029 'Parametro iCodigo
'Erro, Sigla já existente para outro código - iCodigo = %s
Public Const ERRO_MOTIVOPERDA_MESMA_DESCRICAO = 500030 'Parametro iCodigo
'Erro, descrição já existente para outro código - iCodigo = %s
Public Const ERRO_MOTIVOPERDA_INSERCAO = 500031 'Parametro iCodigo
'Erro na Inserção do 'Motivo Perda' código - iCodigo = %s
Public Const ERRO_MOTIVOPERDA_ATUALIZACAO = 500032 'Parametro iCodigo = %s
'Erro na Atualização do 'Motivo Perda' código - iCodigo%s
Public Const ERRO_MOTIVOPERDA_NAO_ENCONTRADO = 500033 'Parametro iCodigo
'O Motivo com o código %s não está cadastrado
Public Const ERRO_LOCK_MOTIVOPERDA = 500034 'Parametro iCodigo
'Erro de 'LOK'na tabela 'Motivo Perda' Código %s - iCodigo
Public Const ERRO_EXCLUSAO_MOTIVOPERDA = 500035 'Parametro iCodigo
'Erro na exclusão do 'Motivo Perda' Código %s - iCodigo
Public Const ERRO_PREVVENDAS_JA_CADASTRADA = 500036 'Parâmetros: sCodigo
'Já existe no banco de dados a previsão de código %s com Datas do período diferentes.
Public Const ERRO_CODIGOPROPOSTA_NAO_NUMERICO = 500037 'Parametro:sCodigo
'%s não é um valor numérico.
Public Const ERRO_FORMATO_CODIGO_PROPOSTA = 500038 'Sem parametros
'O código da Proposta não está no formato correto.
Public Const ERRO_LEITURA_PROPOSTAVENDA = 500039 'Sem parametros
'Erro na leitura da tabela PropostaVenda.
Public Const ERRO_ATUALIZACAO_PROPOSTAVENDA = 500040 'Parametro:sCodigo
'Erro na atualização da Proposta de Venda %s.
Public Const ERRO_LOCK_COTVENDA_ANALISEECONOMICA = 500041 'Sem parametros
'Erro na tentativa de lock na tabela CotVendaAnáliseEconômica.
Public Const ERRO_INSERCAO_PROPOSTAVENDA = 500042 'Parametro:sCodigo
'Erro na inserção da Proposta de Venda %s.
Public Const ERRO_CODIGO_PROPOSTAVENDA_NAO_PREENCHIDO = 500043 'Sem parametros
'O Código da Proposta de Venda não está preenchido.
Public Const ERRO_EXCLUSAO_PROPOSTAVENDA = 500044 'Parametro: sCodigoProposta
'Erro na tentativa de excluir a Proposta de Venda %s.
Public Const ERRO_PROPOSTAVENDA_NAO_CADASTRADA = 500045 'Sem parametros
'A Proposta de Venda não está cadastrada.
Public Const ERRO_LOCK_PROPOSTAVENDA = 500046 'Parametro:sCodigo
'Erro na tentativa de lock na Proposta de Venda %s.
Public Const ERRO_MOTIVOPERDA_ASSOCIADO_COTVENDA = 500047 'Parametro: iCodigo
'O Motivo de Perda %i não pode ser excluído pois está vinculado a uma Cotação de Venda.
Public Const ERRO_LEITURA_COTVENDA_ANALISEECONOMICA = 500048 'Sem parametros
'Erro na leitura da tabela CotVendaAnáliseEconômica.
Public Const ERRO_INSERCAO_COTVENDA_OUTROSSERVICOS = 500049 'Sem parametros
'Erro na inserção na tabela CotVendaOutrosServiços.
Public Const ERRO_LEITURA_COTVENDA_OUTROSSERVICOS = 500050 'Sem parametros
'Erro na leitura da tabela CotVendaOutrosServiços.
Public Const ERRO_EXCLUSAO_COTVENDA_OUTROSSERVICOS = 500051 'Sem parametros
'Erro na exclusão na tabela CotVendaOutrosServiços.
Public Const ERRO_INSERCAO_COTVENDA_ANALISEECONOMICA = 500052 'Sem parametros
'Erro na inserção na tabela CotVendaAnáliseEconômica.
Public Const ERRO_EXCLUSAO_COTVENDA_ANALISEECONOMICA = 500053 'Sem parametros
'Erro na exclusão na tabela CotVendaAnáliseEconômica.
Public Const ERRO_ATUALIZACAO_COTVENDA_ANALISEECONOMICA = 500054 'Sem parametros
'Erro na atualização da tabela CotVendaAnáliseEconômica.
Public Const ERRO_COTACAOVENDA_ANALISEECONOMICA_NAO_CADASTRADA = 500055 'Sem parametros
'A Análise Econômica não está cadastrada.
Public Const ERRO_TERMORESPONSABILIDADE_NAO_PREENCHIDO = 500056 'sem parametros
'Não é possível imprimir o termo de responsabilidade porque ele não está preenchido.
Public Const ERRO_LEITURA_COTVENDA_CARGAS = 500057 'sem parametros
'Erro na leitura da tabela CotVendaCargas.
Public Const ERRO_EXCLUSAO_COTVENDA_CARGAS = 500058 'sem parametros
'Erro na tentativa de exclusão na tabela CotVendaCargas.
Public Const ERRO_LOCK_COTVENDA_ANALISETECNICA = 500059 'sem parametros
'Erro na tentativa de lock na Análise Técnica.
Public Const ERRO_EXCLUSAO_COTVENDA_ANALISETECNICA = 500060 'Sem parametros
'Erro na exclusão na tabela CotVendaAnáliseTécnica.
Public Const ERRO_COTACAOVENDA_ANALISETECNICA_NAO_CADASTRADA = 500061 'Sem parametros
'A análise técnica não está cadastrada.
Public Const ERRO_ANALISETECNICA_INEXISTENTE = 500062 'Parametro: lCodigoCotacao
'Não existe Análise Técnica para Cotação de Venda com código %l.
Public Const ERRO_LEITURA_COTVENDA_ANALISETECNICA = 500063 'Sem parametros
'Erro na leitura da tabela CotVendaAnáliseTécnica.
Public Const ERRO_INSERCAO_COTVENDA_ANALISETECNICA = 500064 'sem parametros
'Erro na inserção na tabela CotVendaAnáliseTécnica.
Public Const ERRO_ATUALIZACAO_COTVENDA_ANALISETECNICA = 500065 'Sem parametros
'Erro na atualização da tabela CotVendaAnáliseTécnica.
Public Const ERRO_INSERCAO_COTVENDA_CARGAS = 500066 'Sem parametros
'Erro na tentativa de inserção na tabela CotVendaCargas.
Public Const ERRO_EQUIPAMENTO_NAO_CADASTRADO = 500067 'Parametro sSigla
'O Equipamento com a sigla %s não está cadastrado
Public Const ERRO_EQUIPAMENTO_MESMA_DESCRICAO = 500068 'Sem parametros
'Existe um outro Equipamento cadastrado com a mesma descrição
Public Const ERRO_EXCLUSAO_EQUIPAMENTO = 500069 'Parametro:sSigla
'Ocorreu um erro na exclusão do equipamento %s.
Public Const ERRO_LEITURA_EQUIPAMENTO = 500070 'Sem Parametro
'Ocorreu um erro na leitura da tabela Equipamentos
Public Const ERRO_INSERCAO_EQUIPAMENTO = 500071 'Parâmetro: sSigla
'Erro na inserção do Equipamento %s.
Public Const ERRO_LOCK_Equipamento = 500072 'Sem parâmetros
'Ocorreu um erro ao tentar fazer o lock de um registro na tabela de Equipamentos.
Public Const ERRO_ATUALIZACAO_EQUIPAMENTO = 500073 'Parâmetro: sSigla
'Erro na atualização do Equipamento %s.
Public Const ERRO_EQUIPAMENTO_VINCULADO_COTACAOVENDAEQUIPAMENTOS = 500074 'parametro sSigla
'O Equipamento %s não pode ser excluído, pois está vinculado à Cotação Vendas Equipamentos.
Public Const ERRO_LEITURA_EQUIPAMENTOS = 500075 'sem parametros
'Erro na leitura da tabela de Equipamentos.
Public Const ERRO_SITUACAO_INEXISTENTE = 500076 'Parametro: sSituacao
'A Situação %s não existe.
Public Const ERRO_LOCK_EQUIPAMENTOS = 500077 'Parametro: sSigla
'Erro na tentativa de lock na tabela Equipamentos para o Equipamento %s.
Public Const ERRO_INSERCAO_COTVENDAEQUIPAMENTOS = 500078 'Sem parametros
'Erro na tentativa de inserção na tabela CotVendaEquipamentos.
Public Const ERRO_LEITURA_MATERIAL2 = 500079 'parametro:sDescricao
'Erro na leitura do material %s na tabela de Materiais.
Public Const ERRO_EXCLUSAO_COTACAOVENDA = 500080 'Parametro: lCodigoCotVenda
'Erro na tentativa de exclusão da Cotação de Venda de código %l.
Public Const ERRO_INSERCAO_COTVENDA_EQUIPAMENTOS = 500081 'Sem parametros
'Erro na tentativa de inserção na tabela CotVendaEquipamentos.
Public Const ERRO_INSERCAO_COTVENDA_CONTATOS = 500082 'Sem parametros
'Erro na tentativa de inserção na tabela CotVendaContatos.
Public Const ERRO_LEITURA_UNIDADEVALOREQUIPAMENTO = 500083 'Sem parametros
'Erro na leitura da tabela UnidadeValorEquipamento.
Public Const ERRO_EXCLUSAO_COTVENDA_CONTATOS = 500084 'Parametro: lCodigoCotVenda
'Erro na tentativa de exclusão do Contato da Cotação de Venda de código %l.
Public Const ERRO_ALTERACAO_COTACAOVENDA = 500085 'Parametro: lCodigoCotVenda
'Erro na tentativa de alterar a Cotação de Venda de código %l.
Public Const ERRO_INSERCAO_COTACAOVENDA = 500086 'parametros: lCodigo
'Erro na tentativa de inserção da Cotação de Venda %l.
Public Const ERRO_LOCK_COTACAOVENDA = 500087 'Parametro: lCodigoCotVenda
'Erro na tentativa de lock na Cotação de Venda com código %l.
Public Const ERRO_EXCLUSAO_COTVENDAEQUIPAMENTOS = 500088 'Sem Parametro
'Erro na tentativa de exclusão na tabela CotVendaEquipamentos.
Public Const ERRO_MATERIAL_NAO_CADASTRADO = 500089 'Parametro: iCodigoMaterial
'O Material com código %i não está cadastrado.
Public Const ERRO_LEITURA_COTVENDAEQUIPAMENTOS = 500090 'Sem parametros
'Erro na leitura da tabela CotVendaEquipamentos.
Public Const ERRO_LEITURA_COTVENDA_CONTATOS = 500091 'sem parametros
'Erro na leitura da tabela CotVendaContatos.
Public Const ERRO_MATERIAL_NAO_CADASTRADO2 = 500092 'Parametro: sDescricaoMaterial
'O Material %s não está cadastrado.
Public Const ERRO_CODIGO_COTACAOVENDA_NAOPREENCHIDO = 500093 'Sem parametros
'O código da Cotação não está preenchido.
Public Const ERRO_EQUIPAMENTO_NAO_EXISTENTE = 500094 'Parametro: sSiglaEquipamento
'O Equipamento com sigla %s não está cadastrado.
Public Const ERRO_EQUIPAMENTO_JA_EXISTENTE_GRIDEQUIPAMENTO = 500095 'Parametros: sSiglaEquipamento, iLinhaGrid
'O Equipamento com a sigla %s já existe na linha %i do Grid.
Public Const ERRO_CODIGO_COTACAOVENDA_NAO_PREENCHIDO = 500096 'Sem Parametros
'O código da Cotação não está preenchido.
Public Const ERRO_DATACOTACAO_NAO_PREENCHIDA = 500097 'Sem parametros
'A Data da Cotação não foi preenchida.
Public Const ERRO_COTACAOVENDA_NAO_CADASTRADO = 500098 'Parametro: lCodigo
'A Cotação com código %l não está cadastrada.
Public Const ERRO_PRODUTO_JA_EXISTENTE_GRIDEQUIPAMENTO = 500099 'Parametros: sSiglaEquipamento, iLinhaGrid
'O Equipamento com a sigla %s já existe na linha %i do Grid.











Public Const AVISO_EXCLUSAO_MOTIVOPERDA = 500500 'Parametro iCodigo
'Confirma a exclusão do 'Motivo de Perda' código %s - iCodigo=%s
Public Const AVISO_CONFIRMA_EXCLUSAO_PROPOSTAVENDA = 500501 'Parametro: sCodigoProposta
'Confirma a exclusão da Proposta de Venda %s?
Public Const AVISO_CONFIRMA_EXCLUSAO_COTVENDA_ANALISEECONOMICA = 500502 'Parametro: lCodCotacao
'Confirma a exclusão da Análise Econômica da Cotação %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_COTVENDA_ANALISETECNICA = 500503 'Parametro:lCodigoCotacao
'Confirma a exclusão da Análise Técnica da Cotação %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_EQUIPAMENTO = 500504 ' parametro: sSigla
'Deseja realmente excluir o equipamento com a sigla %s?
Public Const AVISO_CRIAR_MATERIAL2 = 500505 'Parametro: sDescricao
'Material com Descrição %s não está cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_MATERIAL = 500506 'Parametro: iCOdigo
'Material com Código %i não está cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_MOTIVOPERDA = 500507 'Parametro:iCodigo
'O Motivo de Perda com código %i não existe. Deseja criá-lo?
Public Const AVISO_CONFIRMA_EXCLUSAO_COTACAOVENDA = 500508 'Parametro: lCodigo
'Confirma exclusão da Cotação com o código %l?



