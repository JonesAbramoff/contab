Attribute VB_Name = "ErrosInpal"
Option Explicit

Public Const ERRO_DESPESAFINANCEIRA_NAO_PREENCHIDA = 500000 'Sem par�metros
'Para Despesa Financeira espec�fica, � obrigat�rio preencher seu valor.
Public Const ERRO_JUROS_NAO_PREENCHIDO = 500001 'Sem par�metros
'Para Juros espec�fico, � obrigat�rio preencher seu valor.
Public Const ERRO_PEDIDO_PROGRAMADO_COM_RESERVA = 500002 'Sem par�metros
'Para um Pedido de Vendas programado, as Reservas n�o devem
'ser preenchidas.
Public Const ERRO_TRANSPORTADORA_NAO_PREENCHIDA = 500003 'Sem par�metros
'Para frete por conta do destinat�rio � obrigat�rio o preenchimento
'da Transportadora.
Public Const ERRO_VALORBASE_ITEM_NAO_PREENCHIDO = 500004 'Par�metros: iLinha
'Valor Base do Item %i do Grid Itens n�o foi preenchido.
Public Const ERRO_VALORORIGINALPARC_NAO_INFORMADO = 500005 'Par�metros: iLinha
'O Valor Original da parcela %s n�o foi informado.
Public Const ERRO_MOTIVODIFERENCAPARC_NAO_INFORMADO = 500006 'Par�metros: iLinha
'O Motivo de Diferen�a da parcela %s n�o foi informado.
Public Const ERRO_SOMA_VALORORIGINAL_DIFERENTE_VALORTOTAL = 500007 'Par�metros: dSomaValorOriginal, dValorTotal
'A soma de valores originais das parcelas %s � diferente do valor total que � %s.
Public Const ERRO_MOTIVODIFERENCA_NAO_ENCONTRADO = 500008 'Par�metros: sMotivo
'O motivo de diferen�a %s n�o est� cadastrado.
Public Const ERRO_SOMA_PARCELAS_DIFERENTE_TOTAL = 500009 'Par�metros: dValorParcelas, dValorTotal, dValorDiferenca
'A soma das parcelas %s � diferente do valor total do t�tulo %s mais o valor de diferen�a %s.
Public Const ERRO_MOTIVODIFERENCA_INFORMADO_ERRADO = 500010 'Par�metros: iParcela
'Foi preenchido um Motivo de Diferen�a para parcela %s, n�o tem diferen�a de valor.
Public Const ERRO_PRECOUNITARIO_NAO_PREENCHIDO = 500011 'Sem par�metros
'� obrigat�rio o preenchimento de Pre�o Unit�rio.
Public Const ERRO_VALOR_DIFERENTE_QUANTPRECOUNITARIO = 500012 'Par�metros: dValor, dValorReal
'O valor %s � diferente da quantidade vezes o pre�o unit�rio, que � %s.
Public Const ERRO_LEITURA_MATERIAL = 500013 'Parametro: iCodigo
'Erro ocorrido na leitura do material %i.
Public Const ERRO_MATERIAL_MESMA_DESCRICAO = 500014 'Parametro :sDescricao
'J� existe material com a descri��o %s.
Public Const ERRO_LOCK_MATERIAl = 500015 'Parametro: iCodigo
'Erro na tentativa de lock do material %i.
Public Const ERRO_ATUALIZACAO_MATERIAL = 500016 'Parametro: iCodigo
'Erro na tentativa de atualizar o material com c�digo %i.
Public Const ERRO_INSERCAO_MATERIAL = 500017 'Parametro: iCodigo
'Erro na tentativa de inserir material com c�digo %i.
Public Const ERRO_CODIGO_MATERIAL_ZERADO = 500018 'Sem Parametro
'O c�digo do material n�o pode ser igual a zero.
Public Const ERRO_MATERIAL_NAO_ENCONTRADO = 500019 'Parametro: iCodigo
'O material com o c�digo %i n�o est� cadastrado
Public Const ERRO_EXCLUSAO_MATERIAL = 500020 'Parametro: iCodigo
'Erro ocorrido na exclus�o do material %i.
Public Const ERRO_LEITURA_MATERIAL_COTACAOVENDA = 500021 'Parametro: iCodigo
'Erro ocorrido na exclus�o de material
Public Const ERRO_MATERIAL_UTILIZADO_COTACAOVENDA = 500022 'Parametro: iCodigo
'Erro de material utilizado na tabela de cota��o de vendas ocorrido na exclus�o de material
'N�o � poss�vel excluir o material com o c�digo %s, pois ele est� sendo utilizado em um cota��io de transportadora
Public Const ERRO_MATERIAL_COTACAOVENDA_NAO_ENCONTRADO = 500023 'Parametro: iCodigo
'Erro de material utilizado n�o encontrado na tabela de cota��o de vendas ocorrido na exclus�o de material
'Public Const ERRO_SIGLA_NAO_PREENCHIDA = 500024
'Preenchimento da sigla � obrigat�rio.
Public Const ERRO_LEITURA_MOTIVOPERDA = 500025 'Parametro iCodigo
'Erro na leitura da tabela 'Motivo Perda' - Codigo = %s
Public Const ERRO_MOTIVOPERDA_NAO_CADASTRADO = 500026 'Parametro iCodigo
'Erro, Motivo Perda n�o cadastrado - Codigo = %s
Public Const ERRO_LEITURA_COTACAOVENDA = 500028 'Parametro iCodigo
'Erro, na leitura de Cotacao Venda - Codigo= %s
Public Const ERRO_MOTIVOPERDA_MESMA_SIGLA = 500029 'Parametro iCodigo
'Erro, Sigla j� existente para outro c�digo - iCodigo = %s
Public Const ERRO_MOTIVOPERDA_MESMA_DESCRICAO = 500030 'Parametro iCodigo
'Erro, descri��o j� existente para outro c�digo - iCodigo = %s
Public Const ERRO_MOTIVOPERDA_INSERCAO = 500031 'Parametro iCodigo
'Erro na Inser��o do 'Motivo Perda' c�digo - iCodigo = %s
Public Const ERRO_MOTIVOPERDA_ATUALIZACAO = 500032 'Parametro iCodigo = %s
'Erro na Atualiza��o do 'Motivo Perda' c�digo - iCodigo%s
Public Const ERRO_MOTIVOPERDA_NAO_ENCONTRADO = 500033 'Parametro iCodigo
'O Motivo com o c�digo %s n�o est� cadastrado
Public Const ERRO_LOCK_MOTIVOPERDA = 500034 'Parametro iCodigo
'Erro de 'LOK'na tabela 'Motivo Perda' C�digo %s - iCodigo
Public Const ERRO_EXCLUSAO_MOTIVOPERDA = 500035 'Parametro iCodigo
'Erro na exclus�o do 'Motivo Perda' C�digo %s - iCodigo
Public Const ERRO_PREVVENDAS_JA_CADASTRADA = 500036 'Par�metros: sCodigo
'J� existe no banco de dados a previs�o de c�digo %s com Datas do per�odo diferentes.
Public Const ERRO_CODIGOPROPOSTA_NAO_NUMERICO = 500037 'Parametro:sCodigo
'%s n�o � um valor num�rico.
Public Const ERRO_FORMATO_CODIGO_PROPOSTA = 500038 'Sem parametros
'O c�digo da Proposta n�o est� no formato correto.
Public Const ERRO_LEITURA_PROPOSTAVENDA = 500039 'Sem parametros
'Erro na leitura da tabela PropostaVenda.
Public Const ERRO_ATUALIZACAO_PROPOSTAVENDA = 500040 'Parametro:sCodigo
'Erro na atualiza��o da Proposta de Venda %s.
Public Const ERRO_LOCK_COTVENDA_ANALISEECONOMICA = 500041 'Sem parametros
'Erro na tentativa de lock na tabela CotVendaAn�liseEcon�mica.
Public Const ERRO_INSERCAO_PROPOSTAVENDA = 500042 'Parametro:sCodigo
'Erro na inser��o da Proposta de Venda %s.
Public Const ERRO_CODIGO_PROPOSTAVENDA_NAO_PREENCHIDO = 500043 'Sem parametros
'O C�digo da Proposta de Venda n�o est� preenchido.
Public Const ERRO_EXCLUSAO_PROPOSTAVENDA = 500044 'Parametro: sCodigoProposta
'Erro na tentativa de excluir a Proposta de Venda %s.
Public Const ERRO_PROPOSTAVENDA_NAO_CADASTRADA = 500045 'Sem parametros
'A Proposta de Venda n�o est� cadastrada.
Public Const ERRO_LOCK_PROPOSTAVENDA = 500046 'Parametro:sCodigo
'Erro na tentativa de lock na Proposta de Venda %s.
Public Const ERRO_MOTIVOPERDA_ASSOCIADO_COTVENDA = 500047 'Parametro: iCodigo
'O Motivo de Perda %i n�o pode ser exclu�do pois est� vinculado a uma Cota��o de Venda.
Public Const ERRO_LEITURA_COTVENDA_ANALISEECONOMICA = 500048 'Sem parametros
'Erro na leitura da tabela CotVendaAn�liseEcon�mica.
Public Const ERRO_INSERCAO_COTVENDA_OUTROSSERVICOS = 500049 'Sem parametros
'Erro na inser��o na tabela CotVendaOutrosServi�os.
Public Const ERRO_LEITURA_COTVENDA_OUTROSSERVICOS = 500050 'Sem parametros
'Erro na leitura da tabela CotVendaOutrosServi�os.
Public Const ERRO_EXCLUSAO_COTVENDA_OUTROSSERVICOS = 500051 'Sem parametros
'Erro na exclus�o na tabela CotVendaOutrosServi�os.
Public Const ERRO_INSERCAO_COTVENDA_ANALISEECONOMICA = 500052 'Sem parametros
'Erro na inser��o na tabela CotVendaAn�liseEcon�mica.
Public Const ERRO_EXCLUSAO_COTVENDA_ANALISEECONOMICA = 500053 'Sem parametros
'Erro na exclus�o na tabela CotVendaAn�liseEcon�mica.
Public Const ERRO_ATUALIZACAO_COTVENDA_ANALISEECONOMICA = 500054 'Sem parametros
'Erro na atualiza��o da tabela CotVendaAn�liseEcon�mica.
Public Const ERRO_COTACAOVENDA_ANALISEECONOMICA_NAO_CADASTRADA = 500055 'Sem parametros
'A An�lise Econ�mica n�o est� cadastrada.
Public Const ERRO_TERMORESPONSABILIDADE_NAO_PREENCHIDO = 500056 'sem parametros
'N�o � poss�vel imprimir o termo de responsabilidade porque ele n�o est� preenchido.
Public Const ERRO_LEITURA_COTVENDA_CARGAS = 500057 'sem parametros
'Erro na leitura da tabela CotVendaCargas.
Public Const ERRO_EXCLUSAO_COTVENDA_CARGAS = 500058 'sem parametros
'Erro na tentativa de exclus�o na tabela CotVendaCargas.
Public Const ERRO_LOCK_COTVENDA_ANALISETECNICA = 500059 'sem parametros
'Erro na tentativa de lock na An�lise T�cnica.
Public Const ERRO_EXCLUSAO_COTVENDA_ANALISETECNICA = 500060 'Sem parametros
'Erro na exclus�o na tabela CotVendaAn�liseT�cnica.
Public Const ERRO_COTACAOVENDA_ANALISETECNICA_NAO_CADASTRADA = 500061 'Sem parametros
'A an�lise t�cnica n�o est� cadastrada.
Public Const ERRO_ANALISETECNICA_INEXISTENTE = 500062 'Parametro: lCodigoCotacao
'N�o existe An�lise T�cnica para Cota��o de Venda com c�digo %l.
Public Const ERRO_LEITURA_COTVENDA_ANALISETECNICA = 500063 'Sem parametros
'Erro na leitura da tabela CotVendaAn�liseT�cnica.
Public Const ERRO_INSERCAO_COTVENDA_ANALISETECNICA = 500064 'sem parametros
'Erro na inser��o na tabela CotVendaAn�liseT�cnica.
Public Const ERRO_ATUALIZACAO_COTVENDA_ANALISETECNICA = 500065 'Sem parametros
'Erro na atualiza��o da tabela CotVendaAn�liseT�cnica.
Public Const ERRO_INSERCAO_COTVENDA_CARGAS = 500066 'Sem parametros
'Erro na tentativa de inser��o na tabela CotVendaCargas.
Public Const ERRO_EQUIPAMENTO_NAO_CADASTRADO = 500067 'Parametro sSigla
'O Equipamento com a sigla %s n�o est� cadastrado
Public Const ERRO_EQUIPAMENTO_MESMA_DESCRICAO = 500068 'Sem parametros
'Existe um outro Equipamento cadastrado com a mesma descri��o
Public Const ERRO_EXCLUSAO_EQUIPAMENTO = 500069 'Parametro:sSigla
'Ocorreu um erro na exclus�o do equipamento %s.
Public Const ERRO_LEITURA_EQUIPAMENTO = 500070 'Sem Parametro
'Ocorreu um erro na leitura da tabela Equipamentos
Public Const ERRO_INSERCAO_EQUIPAMENTO = 500071 'Par�metro: sSigla
'Erro na inser��o do Equipamento %s.
Public Const ERRO_LOCK_Equipamento = 500072 'Sem par�metros
'Ocorreu um erro ao tentar fazer o lock de um registro na tabela de Equipamentos.
Public Const ERRO_ATUALIZACAO_EQUIPAMENTO = 500073 'Par�metro: sSigla
'Erro na atualiza��o do Equipamento %s.
Public Const ERRO_EQUIPAMENTO_VINCULADO_COTACAOVENDAEQUIPAMENTOS = 500074 'parametro sSigla
'O Equipamento %s n�o pode ser exclu�do, pois est� vinculado � Cota��o Vendas Equipamentos.
Public Const ERRO_LEITURA_EQUIPAMENTOS = 500075 'sem parametros
'Erro na leitura da tabela de Equipamentos.
Public Const ERRO_SITUACAO_INEXISTENTE = 500076 'Parametro: sSituacao
'A Situa��o %s n�o existe.
Public Const ERRO_LOCK_EQUIPAMENTOS = 500077 'Parametro: sSigla
'Erro na tentativa de lock na tabela Equipamentos para o Equipamento %s.
Public Const ERRO_INSERCAO_COTVENDAEQUIPAMENTOS = 500078 'Sem parametros
'Erro na tentativa de inser��o na tabela CotVendaEquipamentos.
Public Const ERRO_LEITURA_MATERIAL2 = 500079 'parametro:sDescricao
'Erro na leitura do material %s na tabela de Materiais.
Public Const ERRO_EXCLUSAO_COTACAOVENDA = 500080 'Parametro: lCodigoCotVenda
'Erro na tentativa de exclus�o da Cota��o de Venda de c�digo %l.
Public Const ERRO_INSERCAO_COTVENDA_EQUIPAMENTOS = 500081 'Sem parametros
'Erro na tentativa de inser��o na tabela CotVendaEquipamentos.
Public Const ERRO_INSERCAO_COTVENDA_CONTATOS = 500082 'Sem parametros
'Erro na tentativa de inser��o na tabela CotVendaContatos.
Public Const ERRO_LEITURA_UNIDADEVALOREQUIPAMENTO = 500083 'Sem parametros
'Erro na leitura da tabela UnidadeValorEquipamento.
Public Const ERRO_EXCLUSAO_COTVENDA_CONTATOS = 500084 'Parametro: lCodigoCotVenda
'Erro na tentativa de exclus�o do Contato da Cota��o de Venda de c�digo %l.
Public Const ERRO_ALTERACAO_COTACAOVENDA = 500085 'Parametro: lCodigoCotVenda
'Erro na tentativa de alterar a Cota��o de Venda de c�digo %l.
Public Const ERRO_INSERCAO_COTACAOVENDA = 500086 'parametros: lCodigo
'Erro na tentativa de inser��o da Cota��o de Venda %l.
Public Const ERRO_LOCK_COTACAOVENDA = 500087 'Parametro: lCodigoCotVenda
'Erro na tentativa de lock na Cota��o de Venda com c�digo %l.
Public Const ERRO_EXCLUSAO_COTVENDAEQUIPAMENTOS = 500088 'Sem Parametro
'Erro na tentativa de exclus�o na tabela CotVendaEquipamentos.
Public Const ERRO_MATERIAL_NAO_CADASTRADO = 500089 'Parametro: iCodigoMaterial
'O Material com c�digo %i n�o est� cadastrado.
Public Const ERRO_LEITURA_COTVENDAEQUIPAMENTOS = 500090 'Sem parametros
'Erro na leitura da tabela CotVendaEquipamentos.
Public Const ERRO_LEITURA_COTVENDA_CONTATOS = 500091 'sem parametros
'Erro na leitura da tabela CotVendaContatos.
Public Const ERRO_MATERIAL_NAO_CADASTRADO2 = 500092 'Parametro: sDescricaoMaterial
'O Material %s n�o est� cadastrado.
Public Const ERRO_CODIGO_COTACAOVENDA_NAOPREENCHIDO = 500093 'Sem parametros
'O c�digo da Cota��o n�o est� preenchido.
Public Const ERRO_EQUIPAMENTO_NAO_EXISTENTE = 500094 'Parametro: sSiglaEquipamento
'O Equipamento com sigla %s n�o est� cadastrado.
Public Const ERRO_EQUIPAMENTO_JA_EXISTENTE_GRIDEQUIPAMENTO = 500095 'Parametros: sSiglaEquipamento, iLinhaGrid
'O Equipamento com a sigla %s j� existe na linha %i do Grid.
Public Const ERRO_CODIGO_COTACAOVENDA_NAO_PREENCHIDO = 500096 'Sem Parametros
'O c�digo da Cota��o n�o est� preenchido.
Public Const ERRO_DATACOTACAO_NAO_PREENCHIDA = 500097 'Sem parametros
'A Data da Cota��o n�o foi preenchida.
Public Const ERRO_COTACAOVENDA_NAO_CADASTRADO = 500098 'Parametro: lCodigo
'A Cota��o com c�digo %l n�o est� cadastrada.
Public Const ERRO_PRODUTO_JA_EXISTENTE_GRIDEQUIPAMENTO = 500099 'Parametros: sSiglaEquipamento, iLinhaGrid
'O Equipamento com a sigla %s j� existe na linha %i do Grid.











Public Const AVISO_EXCLUSAO_MOTIVOPERDA = 500500 'Parametro iCodigo
'Confirma a exclus�o do 'Motivo de Perda' c�digo %s - iCodigo=%s
Public Const AVISO_CONFIRMA_EXCLUSAO_PROPOSTAVENDA = 500501 'Parametro: sCodigoProposta
'Confirma a exclus�o da Proposta de Venda %s?
Public Const AVISO_CONFIRMA_EXCLUSAO_COTVENDA_ANALISEECONOMICA = 500502 'Parametro: lCodCotacao
'Confirma a exclus�o da An�lise Econ�mica da Cota��o %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_COTVENDA_ANALISETECNICA = 500503 'Parametro:lCodigoCotacao
'Confirma a exclus�o da An�lise T�cnica da Cota��o %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_EQUIPAMENTO = 500504 ' parametro: sSigla
'Deseja realmente excluir o equipamento com a sigla %s?
Public Const AVISO_CRIAR_MATERIAL2 = 500505 'Parametro: sDescricao
'Material com Descri��o %s n�o est� cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_MATERIAL = 500506 'Parametro: iCOdigo
'Material com C�digo %i n�o est� cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_MOTIVOPERDA = 500507 'Parametro:iCodigo
'O Motivo de Perda com c�digo %i n�o existe. Deseja cri�-lo?
Public Const AVISO_CONFIRMA_EXCLUSAO_COTACAOVENDA = 500508 'Parametro: lCodigo
'Confirma exclus�o da Cota��o com o c�digo %l?



