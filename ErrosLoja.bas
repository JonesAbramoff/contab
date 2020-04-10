Attribute VB_Name = "ErrosLoja"
Option Explicit

'C�digos de Erro - Reservado de 13500 a 13699
Public Const ERRO_LEITURA_ALIQUOTAICMS = 13500
'Erro na Leitura da tabela Aliquota ICMS.
Public Const ERRO_LEITURA_LOJACONFIG = 13501 'Parametro: sCodigo
'Erro na leitura da tabela Loja Config.
Public Const ERRO_ATUALIZACAO_LOJACONFIG = 13502 'Parametro: sCodigo
'Erro na atualiza��o da tabela Loja Config.
'Public Const ERRO_CODIGO_NAO_EXISTE = 13503 'Parametro - sCodigo
'O %s n�o existe na tabela Loja Config - sCodigo
Public Const ERRO_REGISTRO_LOJA_CONFIG_NAO_ENCONTRADO = 13504 'Parametro: sCodigo
'N�o foi encontrado o registro com o c�digo %s na tabela LojaConfig
Public Const ERRO_ESPACOENTRELINHAS_NAO_PREENCHIDO = 13505
'O espa�o entre linhas n�o foi informado.
Public Const ERRO_LINHASENTRECUPONS_NAO_PREENCHIDO = 13506
'O campo linhas entre cupons n�o foi informado.
Public Const ERRO_INSERCAO_ALIQUOTASICMS = 13507
'Erro na Inclus�o de Registro na Tabela de AliquotasICMS
Public Const ERRO_EXCLUSAO_ALIQUOTASICMS = 13508
'Erro na exclus�o de Registro na Tabela AliquotasICMS
Public Const ERRO_LEITURA_CAIXAS = 13509 'Par�metro: iCodigo
'Ocorreu um erro na leitura da tabela de Caixas. Caixa= %s
Public Const ERRO_LEITURA_MOVIMENTOS_CAIXA = 13510 'Sem Par�metros
'Ocorreu um erro na leitura da tabela de Movimentos de Caixa (MovimentosCaixa)
Public Const ERRO_LEITURA_ECF = 13511 'Sem par�metros
'Ocorreu um erro na leitura da tabela de ECF's (ECF).
Public Const ERRO_LEITURA_SESSAO = 13512 'Sem par�metros
'Ocorreu um erro na leitura da tabela de Sess�es (Sessao).
Public Const ERRO_EXCLUSAO_CAIXA = 13513 'Par�metros: iCodigo, sNomeReduzido
'Ocorreu um na exclus�o do Caixa %s - %s.
Public Const ERRO_LOCK_CAIXA = 13514 'Sem par�metros
'Ocorreu um erro ao tentar fazer o lock de um registro da tabela de Caixas(Caixa).
Public Const ERRO_CAIXA_NAO_CADASTRADO = 13515 'Par�metro: iCodigo
'O caixa %s n�o est� cadastrado.
Public Const ERRO_NOME_REDUZIDO_CAIXA_REPETIDO = 13516 'Par�metros: iCodigo, sNomeReduzido
'O caixa %s j� utiliza o nome reduzido %s
Public Const ERRO_DATAINICIALCAIXA_MAIOR_DATAMOVIMENTOCAIXA = 13517 'Par�metros: dtDataInicial, dtDataMovimento
'A data de inicializa��o do Caixa (%s) n�o pode ser maior do que a data do primeiro movimento registrado para o mesmo Caixa (%s).
Public Const ERRO_PRIMEIRO_CAIXA_DEVE_SER_CENTRAL = 13518 'Sem par�metros
'O primeiro Caixa de uma filial deve ser configurado como Caixa Central.
Public Const ERRO_CAIXA_CENTRAL_NAO_PODE_SER_ALTERADO = 13519 'Par�metros: iCodigo, sNomeReduzido, giFilialEmpresa
'O Caixa %s - %s est� configurado como Caixa Central da filial %s e n�o pode ser alterado para Caixa Comum.
Public Const ERRO_EXCLUSAO_CAIXA_CENTRAL = 13520 'Par�metros: iCodigo, sNomeReduzido
'O Caixa %s - %s est� configurado como Central e n�o pode ser exclu�do.
Public Const ERRO_INSERCAO_CAIXA = 13521 'Par�metro: iCodigo, sNomeReduzido
'Erro na inser��o do Caixa %s - %s
Public Const ERRO_ATUALIZACAO_CAIXA = 13522 'Par�metro: iCodigo, sNomeReduzido
'Erro na atualiza��o do Caixa %s - %s
Public Const ERRO_CAIXA_VINCULADO_MOVIMENTOCAIXA = 13523 'Par�metro: iCodigo, sNomeReduzido
'O caixa %s - %s n�o pode ser exclu�do, pois existem movimentos de caixa vinculados a ele.
Public Const ERRO_CAIXA_VINCULADO_ECF = 13524 'Par�metro: iCodigo, sNomeReduzido
'O caixa %s - %s n�o pode ser exclu�do, pois existem Emissoras de Cupom Fiscal vinculadas a ele.
Public Const ERRO_CAIXA_VINCULADO_SESSAO = 13525 'Par�metros:iCodigo, sNomeReduzido
'O caixa %s - %s n�o pode ser exclu�do, pois existem Sess�es vinculadas a ele.
Public Const ERRO_ALTERACAO_CAIXA_OUTRA_FILIAL = 13526 'Par�metros: iCodigo, sNomeReduzido, iFilialEmpresa
'O Caixa %s - %s n�o pode ser alterado, pois pertence � filial %s.
Public Const ERRO_EXCLUSAO_CAIXA_OUTRA_FILIAL = 13527 'Par�metros: iCodigo, sNomeReduzido, iFilialEmpresa
'O Caixa %s - %s n�o pode ser exclui�do, pois pertence � filial %s.
Public Const ERRO_SIGLA_EXISTE = 13528 'SIGLA
'A sigla %s j� existe
Public Const ERRO_ATUALIZACAO_PRODUTOSFILIAL1 = 13529 'Par�metros: objLojaConfig.iTabelaPreco
'Erro na tentativa de atualizar registro na tabela ProdutosFilial com Tabela de Pre�o %s.
Public Const ERRO_OPERADOR_NAO_CADASTRADO = 13530 'Parametro: iCodigo
'O operador com o c�digo %i n�o esta cadastrado
Public Const ERRO_LEITURA_OPERADOR2 = 13531
'Erro de leitura na tabela de Operadores.
Public Const ERRO_LEITURA_OPERADOR1 = 13532 'Parametro: sCodUsuario
'Erro de leitura do operador com o c�digo de usu�rio %s.
Public Const ERRO_OPERADOR_USUARIO = 13533 'Parametros: iCodOperador,sCodUsuario
'O c�digo de Operador %i correspondente ao Usu�rio %s no Bando de Dados n�o confere com o Operador %i da Tela.
Public Const ERRO_LOCK_OPERADOR = 13534
'Erro na tentativa de fazer 'lock' na tabela Operador.
Public Const ERRO_LEITURA_OPERADOR = 13535 'Parametro: iCodigo
'Erro de leitura do Operador com o c�digo %i na tabela de operadores.
Public Const ERRO_ATUALIZACAO_OPERADOR = 13536 'Parametro: iCodOperador
'Erro na tentativa de atualizar o Operador %i a tabela Operador.
Public Const ERRO_INSERCAO_OPERADOR = 13537 'Parametro: iCodOperador
'Erro na tentiva de inserir o Operador %i na tabela Operador.
Public Const ERRO_USUARIO_OPERADOR_NAO_ALTERAVEL = 13538 'Parametro: iCodigo
'O Operador com o c�digo %i nao esta cadastrado.
Public Const ERRO_EXCLUSAO_OPERADOR = 13539 ' Parametro: iCodOperador
'Erro na tentativa de excluir o operador %i na tabela Operador.
Public Const ERRO_OPERADOR_VINCULADO_BOLETO = 13540 'Parametro: icodigo
'Operador %i n�o pode ser exclu�do pois est� vinculado a um registro na tabela Boleto.
Public Const ERRO_OPERADOR_VINCULADO_VALETICKET = 13541 'Parametro: icodigo
'Operador %i n�o pode ser exclu�do pois est� vinculado a um registro na tabela ValeTicket.
Public Const ERRO_CATEGORIAPRODUTOITEM_EXISTE = 13542 'Parametro: CategoriaProdutoItem, CategoriaProduto
'O item %s e a categoria %s j� foram definidas no grid.
Public Const ERRO_PRODUTO_CODBARRAS_NAO_PREENCHIDO = 13543
'O c�digo de barras do produto n�o foi preenchido
Public Const ERRO_PRODUTO_REFERENCIA_NAO_PREENCHIDA = 13544
'A refer�ncia do produto n�o foi preenchida
Public Const ERRO_PRODUTO_SEM_TABELAPRECO_PADRAO = 13545
'N�o existe pre�o cadastrado para o produto em quest�o
Public Const ERRO_PRECO_PRODUTO_NAO_CADASTRADO = 13546 'Parametro sCodProduto
'o produto %s n�o tem pre�o cadastrado
Public Const ERRO_GERENTE_NAO_CADASTRADO = 13547 'Parametro: iCodigo
'O gerente com o c�digo %i n�o esta cadastrado.
Public Const ERRO_LEITURA_GERENTE2 = 13548
'Erro de leitura na tabela de Gerentes.
Public Const ERRO_LEITURA_GERENTE1 = 13549 'Parametro: sCodUsuario
'Erro de leitura do gerente com o c�digo de usu�rio %s.
Public Const ERRO_GERENTE_USUARIO = 13550 'Parametros: iCodGerente,sCodUsuario
'O c�digo de Gerente %i correspondente ao Usu�rio %s no Bando de Dados n�o confere com o Gerente %i da Tela.
Public Const ERRO_LOCK_GERENTE = 13551
'Erro na tentativa de fazer 'lock' na tabela Gerente.
Public Const ERRO_LEITURA_GERENTE = 13552 'Parametro: iCodigo
'Erro de leitura do Gerente com o c�digo %i na tabela de gerentes.
Public Const ERRO_ATUALIZACAO_GERENTE = 13553 'Parametro: iCodGerente
'Erro na tentativa de atualizar o Gerente %i a tabela Gerente.
Public Const ERRO_INSERCAO_GERENTE = 13554 'Parametro: iCodGerente
'Erro na tentiva de inserir o Gerente %i na tabela Gerente.
Public Const ERRO_USUARIO_GERENTE_NAO_ALTERAVEL = 13555 'Parametro: iCodigo
'O Gerente com o c�digo %i n�o esta cadastrado.
Public Const ERRO_EXCLUSAO_GERENTE = 13556 ' Parametro: iCodGerente
'Erro na tentativa de excluir o gerente %i na tabela Gerente.
Public Const ERRO_GERENTE_VINCULADO_SESSAO = 13557 'Parametro: icodigo
'Gerente %i n�o pode ser exclu�do pois est� vinculado a um registro na tabela Sess�o.
Public Const ERRO_GERENTE_VINCULADO_CUPOMFISCAL = 13558 'Parametro: sCodUsuario
'Gerente %i n�o pode ser exclu�do pois esta vinculado a um registro na tabela Cupom Fiscal.
Public Const ERRO_POS_NAO_CADASTRADO = 13559 'Parametro:sCodigo
'O POS %s n�o est� cadastrado.
Public Const ERRO_LEITURA_POS = 13560 'Parametro:sCodigo
'Erro na leitura da tabela de POS. POS C�digo %s.
Public Const ERRO_LEITURA_POS1 = 13561 'Sem Parametros
'Erro na leitura da tabela de POS.
Public Const ERRO_REDE_NAO_PREENCHIDA = 13562 'Sem Parametro
'A rede deve ser preenchida.
Public Const ERRO_INSERCAO_POS = 13563 'Parametro:sCodigo
'Erro na tentativa de inserir o POS %s na tabela de POS
Public Const ERRO_ATUALIZACAO_POS = 13564 'Parametro:sCodigo
'Erro na tentativa de atualizar o POS %s na tabela de POS
Public Const ERRO_LOCK_POS = 13565 'Parametro:sCodigo
'Erro na tentativa de fazer "lock" na tabela de POS. POS %s.
Public Const ERRO_EXCLUSAO_POS = 13566 'Parametro:sCodigo
'Erro na tentativa de excluir um registro na tabela de POS. POS %s.
Public Const ERRO_POS_VINCULADO_ECF = 13567 'Parametro:sCodigo
'O POS %s n�o pode ser exclu�do pois est� vinculado � ECF
Public Const ERRO_LEITURA_BOLETO = 13568 'Sem Par�metros
'Erro na leitura da tabela de Boleto.
Public Const ERRO_POS_VINCULADO_BOLETO = 13569 'Parametro:sCodigo
'O POS %s n�o pode ser exclu�do pois est� vinculado � Boleto
Public Const ERRO_POS_OUTRA_FILIALEMPRESA = 13570 'Parametro:sCodigo, iFilialEmpresaPOS
'O POS c�digo %s pertence a filial %i.
Public Const ERRO_LEITURA_REDE = 13571 'Par�metro: iRede
'Erro na leitura da Rede %s na tabela de Redes
Public Const ERRO_REDE_NAO_ENCONTRADA = 13572 'Par�metro: Rede
'A Rede %s n�o foi encontrada.
Public Const ERRO_INTERVALOVARIAVEL_EXISTE = 13573 'CondicoesPagto, IntervalosVariaveis
'A Condi��o de Pagamento %s com o intervalo %s j� existem
Public Const ERRO_TAXA_NAO_PREENCHIDA = 13574
'Deve ser informada se a taxa de pagamento � � vista ou � prazo
Public Const ERRO_LEITURA_VALETICKET = 13575
'Erro na leitura da tabela de Vale Ticket
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_VALETICKET = 13576 'Parametro:sCodigo
'A Administradora %s n�o pode ser exclu�da pois est� vinculada � um Vale Ticket
Public Const ERRO_LEITURA_FECHAMENTOBOLETOS = 13577
'Erro na leitura da tabela de Fechamento Boletos
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_FECHAMENTOBOLETOS = 13578 'Parametro:sCodigo
'A Administradora %s n�o pode ser exclu�da pois est� vinculada � um Fechamento Boletos
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_BOLETO = 13579 'Parametro:sCodigo
'A administradora %s n�o pode ser exclu�da pois est� vinculada � um Boleto
Public Const ERRO_LEITURA_BORDEROVALETICKET = 13580 'Parametro:sCodigo
'Erro na leitura da tabela de Bordero Vale Ticket
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_BORDEROVALETICKET = 13581 'Parametro:sCodigo
'A administradora n�o pode ser exclu�da pois est� vinculada � um Bordero Vale Ticket
Public Const ERRO_LEITURA_BORDEROBOLETO = 13582
'Erro na leitura da tabela de Bordero Boleto
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_BORDEROBOLETO = 13583 'Parametro:sCodigo
'A administradora n�o pode ser exclu�da pois est� vinculada � um Bordero Boleto
Public Const ERRO_ADMMEIOPAGTO_NAO_CADASTRADO = 13584 'Parametro: sCodigo
'A Administradora %s n�o est� cadastrada
Public Const ERRO_LOCK_ADMMEIOPAGTOCONDPAGTO = 13585 'Sem Par�metros
'Erro na tentativa de fazer 'lock' na tabela AdmMeioPagtoCondPagto.
Public Const ERRO_EXCLUSAO_ADMMEIOPAGTO = 13586 'Parametro:sCodigo
'Erro na exclus�o da Administradora %s.
Public Const ERRO_INSERCAO_ADMMEIOPAGTOCONDPAGTO = 13587 'Sem Parametros
'Erro na tentativa de inser��o de um registro da tabela AdmMeioPagtoCondPagto
Public Const ERRO_ATUALIZACAO_ADMINISTRADORA = 13588 'Parametro: sCodigo
'Erro na tentativa de atualizar administradora %s
Public Const ERRO_EXCLUSAO_ADMMEIOPAGTOCONDPAGTO = 13589
'Erro na tentativa de exclus�o de um registro da tabela AdmMeioPagtoCondPagto
Public Const ERRO_INSERCAO_ADMMEIOPAGTO = 13590 'Parametro:sCodigo
'Erro na tentativa de inser��o de um registro da tabela AdmMeioPagto
Public Const ERRO_SELECIONAR_CONDPAGTO_2VEZES = 13591 'Sem Parametro
'Uma condi��o de pagamento n�o pode ser selecionada mais de duas vezes
Public Const ERRO_SELECIONAR_CONDPAGTO = 13592 'Sem Parametro
'Selecionar na lista a condi��o de pagamento
Public Const ERRO_TIPOMEIOPAGTO_NAO_PREENCHIDO = 13593 'Sem Parametro
'Preenchimento do Meio pagto � obrigat�rio
Public Const ERRO_LOCK_ADMMEIOPAGTO = 13594 'Parametro: sCodigo
'Erro na tentativa de fazer 'lock' na tabela de AdmMeioPagto. Administradora %s
Public Const ERRO_LEITURA_ADMMEIOPAGTOCONDPAGTO = 13595
'Erro na leitura da tabela de AdmMeioPagtoCondPagto
Public Const ERRO_LEITURA_ADMMEIOPAGTO = 13596 'Parametro:sCodigo
'Erro na leitura da tabela de AdmMeioPagto. Administradora %s.
Public Const ERRO_LEITURA_ADMMEIOPAGTO1 = 13597 'Parametro:sCodigo
'Erro na leitura da tabela de AdmMeioPagto.
Public Const ERRO_LEITURA_TIPOMEIOPAGTOLOJA = 13598 'Sem Parametro
'Erro na leitura da tabela de TipoMeioPagtoLoja.
Public Const ERRO_LEITURA_FABRICANTEECF = 13599 'Par�metro: iCodigo
'Ocorreu um erro na leitura da tabela de Fabricantes (FabricanteECF). Fabricante = %s
Public Const ERRO_LEITURA_CUPOMFISCAL = 13600 'Sem Par�metros
'Ocorreu um erro na leitura da tabela de Cupons Fiscais (CupomFiscal).
Public Const ERRO_LEITURA_CHEQUE = 13601 'Sem par�metros
'Ocorreu um erro na leitura da tabela de Cheques (Cheque).
Public Const ERRO_LOCK_ECF = 13602 'Sem par�metros
'Ocorreu um erro ao tentar fazer o lock de um registro na tabela de ECF's (ECF).
Public Const ERRO_LOCK_FABRICANTEECF = 13603 'Sem par�metros
'Ocorreu um erro ao tentar fazer o lock de um registro da tabela de Fabricantes(FabricanteECF).
Public Const ERRO_FABRICANTE_ECF_NAO_SELECIONADO = 13604 'Sem par�metros
'O Fabricante da ECF n�o foi selecionado.
Public Const ERRO_FABRICANTE_ECF_NAO_CADASTRADO = 13605 'Par�metros: iCodigo
'O Fabricante %s n�o est� cadastrado.
Public Const ERRO_LEITURA_ECF1 = 13606 'Sem par�metros
'Ocorreu um erro na leitura da tabela de ECF'S (ECF).
Public Const ERRO_LEITURA_FABRICANTEECF1 = 13607 'Sem par�metros
'Ocorreu um erro na leitura da tabela de Fabricantes (FabricanteECF).
Public Const ERRO_POS_PERTENCE_OUTRA_FILIAL = 13608 'Par�metros: sCodigo, objPOS.iFilialEmpresa
'O POS com c�digo %s pertence � Filial %s e n�o pode ser usado nessa filial.
Public Const ERRO_CAIXA_ECF_NAO_SELECIONADO = 13609 'Sem par�metros
'O Caixa da ECF n�o foi selecionado.
Public Const ERRO_INSERCAO_ECF = 13610 'Par�metro: iCodigo
'Erro na inser��o do ECF %s.
Public Const ERRO_ATUALIZACAO_ECF = 13611 'Par�metro: iCodigo
'Erro na atualiza��o do ECF %s.
Public Const ERRO_EXCLUSAO_ECF = 13612 'Par�metro: iCodigo
'Ocorreu um na exclus�o do ECF %s.
Public Const ERRO_ECF_NAO_CADASTRADO = 13613 'Par�metro: iCodigo
'O ECF %s n�o est� cadastrado.
Public Const ERRO_ALTERACAO_ECF_OUTRA_FILIAL = 13614 'Par�metros: iCodigo, iFilialEmpresaECF
'O ECF %s n�o pode ser alterado, pois pertence � filial %s.
Public Const ERRO_EXCLUSAO_ECF_OUTRA_FILIAL = 13615 'Par�metros: iCodigo, iFilialEmpresaECF
'O ECF %s n�o pode ser exclui�do, pois pertence � filial %s.
Public Const ERRO_ECF_VINCULADO_VALETICKET = 13616 'Par�metro: iCodigo
'O ECF com c�digo %s n�o pode ser exclu�do, pois existem Tickets vinculados a ele.
Public Const ERRO_ECF_VINCULADO_BOLETO = 13617 'Par�metro: iCodigo
'O ECF com c�digo %s n�o pode ser exclu�do, pois existem Boletos vinculados a ele.
Public Const ERRO_ECF_VINCULADO_CUPOMFISCAL = 13618 'Par�metro: iCodigo
'O ECF com c�digo %s n�o pode ser exclu�do, pois existem Cupons Fiscais vinculados ele.
Public Const ERRO_ECF_VINCULADO_CHEQUE = 13619 'Par�metro: iCodigo
'O ECF com c�digo %s n�o pode ser exclu�do, pois existem Cheques vinculados a ele.
Public Const ERRO_LEITURA_CAIXAS1 = 13620 'Sem par�metros
'Ocorreu um erro na leitura da tabela de caixas(Caixa).
Public Const ERRO_LEITURA_CAIXAS2 = 13621 'Par�metro: sNomeReduzido
'Ocorreu um erro na leitura da tabela de caixas(Caixa). Caixa = %s



'C�DIGOS DE AVISO - Reservado de 15400 a 15499
Public Const AVISO_EXCLUIR_CAIXA = 15400 'Par�metro: iCodigo, sNomeReduzido
'Confirma exclus�o do Caixa %s - %s ?
Public Const AVISO_CONFIRMA_EXCLUSAO_OPERADOR = 15401 'Parametro: sCodUsuario
'Confirma a exclus�o do operador %s ?
Public Const AVISO_CONFIRMA_EXCLUSAO_GERENTE = 15402 'Parametro: sCodUsuario
'Confirma a exclus�o do gerente %s ?
Public Const AVISO_EXCLUIR_POS = 15403 'Parametro:sCodigo
'Confirma a exclus�o do POS com c�digo %s ?
Public Const AVISO_EXCLUIR_ADMINISTRADORA = 15404 'Parametro:sCodigo
'Confirma exclus�o da administradora com c�digo %s?
Public Const AVISO_DESEJA_CRIAR_CAIXA = 15405 'Par�metro: iCodigo ou sNomeReduzido
'O Caixa %s n�o est� cadastrado. Deseja cri�-lo agora?
Public Const AVISO_DESEJA_CRIAR_POS = 15406 'Par�metro: sCodigo
'O POS com c�digo %s n�o est� cadastrado. Deseja cri�-lo agora?
Public Const AVISO_EXCLUIR_ECF = 15407 'Par�metro: iCodigo
'Confirma exclus�o do ECF %s?


'******************** Code Start **************************
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Private Declare Function apiGetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function fOSMachineName() As String
'Returns the computername
Dim lngLen As Long, lngX As Long
Dim strCompName As String
    lngLen = 16
    strCompName = String$(lngLen, 0)
    lngX = apiGetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        fOSMachineName = left$(strCompName, lngLen)
    Else
        fOSMachineName = ""
    End If
End Function
'******************** Code End **************************



