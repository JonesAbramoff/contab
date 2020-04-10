Attribute VB_Name = "ErrosLoja"
Option Explicit

'Códigos de Erro - Reservado de 13500 a 13699
Public Const ERRO_LEITURA_ALIQUOTAICMS = 13500
'Erro na Leitura da tabela Aliquota ICMS.
Public Const ERRO_LEITURA_LOJACONFIG = 13501 'Parametro: sCodigo
'Erro na leitura da tabela Loja Config.
Public Const ERRO_ATUALIZACAO_LOJACONFIG = 13502 'Parametro: sCodigo
'Erro na atualização da tabela Loja Config.
'Public Const ERRO_CODIGO_NAO_EXISTE = 13503 'Parametro - sCodigo
'O %s não existe na tabela Loja Config - sCodigo
Public Const ERRO_REGISTRO_LOJA_CONFIG_NAO_ENCONTRADO = 13504 'Parametro: sCodigo
'Não foi encontrado o registro com o código %s na tabela LojaConfig
Public Const ERRO_ESPACOENTRELINHAS_NAO_PREENCHIDO = 13505
'O espaço entre linhas não foi informado.
Public Const ERRO_LINHASENTRECUPONS_NAO_PREENCHIDO = 13506
'O campo linhas entre cupons não foi informado.
Public Const ERRO_INSERCAO_ALIQUOTASICMS = 13507
'Erro na Inclusão de Registro na Tabela de AliquotasICMS
Public Const ERRO_EXCLUSAO_ALIQUOTASICMS = 13508
'Erro na exclusão de Registro na Tabela AliquotasICMS
Public Const ERRO_LEITURA_CAIXAS = 13509 'Parâmetro: iCodigo
'Ocorreu um erro na leitura da tabela de Caixas. Caixa= %s
Public Const ERRO_LEITURA_MOVIMENTOS_CAIXA = 13510 'Sem Parâmetros
'Ocorreu um erro na leitura da tabela de Movimentos de Caixa (MovimentosCaixa)
Public Const ERRO_LEITURA_ECF = 13511 'Sem parâmetros
'Ocorreu um erro na leitura da tabela de ECF's (ECF).
Public Const ERRO_LEITURA_SESSAO = 13512 'Sem parâmetros
'Ocorreu um erro na leitura da tabela de Sessões (Sessao).
Public Const ERRO_EXCLUSAO_CAIXA = 13513 'Parâmetros: iCodigo, sNomeReduzido
'Ocorreu um na exclusão do Caixa %s - %s.
Public Const ERRO_LOCK_CAIXA = 13514 'Sem parâmetros
'Ocorreu um erro ao tentar fazer o lock de um registro da tabela de Caixas(Caixa).
Public Const ERRO_CAIXA_NAO_CADASTRADO = 13515 'Parâmetro: iCodigo
'O caixa %s não está cadastrado.
Public Const ERRO_NOME_REDUZIDO_CAIXA_REPETIDO = 13516 'Parâmetros: iCodigo, sNomeReduzido
'O caixa %s já utiliza o nome reduzido %s
Public Const ERRO_DATAINICIALCAIXA_MAIOR_DATAMOVIMENTOCAIXA = 13517 'Parâmetros: dtDataInicial, dtDataMovimento
'A data de inicialização do Caixa (%s) não pode ser maior do que a data do primeiro movimento registrado para o mesmo Caixa (%s).
Public Const ERRO_PRIMEIRO_CAIXA_DEVE_SER_CENTRAL = 13518 'Sem parâmetros
'O primeiro Caixa de uma filial deve ser configurado como Caixa Central.
Public Const ERRO_CAIXA_CENTRAL_NAO_PODE_SER_ALTERADO = 13519 'Parâmetros: iCodigo, sNomeReduzido, giFilialEmpresa
'O Caixa %s - %s está configurado como Caixa Central da filial %s e não pode ser alterado para Caixa Comum.
Public Const ERRO_EXCLUSAO_CAIXA_CENTRAL = 13520 'Parâmetros: iCodigo, sNomeReduzido
'O Caixa %s - %s está configurado como Central e não pode ser excluído.
Public Const ERRO_INSERCAO_CAIXA = 13521 'Parâmetro: iCodigo, sNomeReduzido
'Erro na inserção do Caixa %s - %s
Public Const ERRO_ATUALIZACAO_CAIXA = 13522 'Parâmetro: iCodigo, sNomeReduzido
'Erro na atualização do Caixa %s - %s
Public Const ERRO_CAIXA_VINCULADO_MOVIMENTOCAIXA = 13523 'Parâmetro: iCodigo, sNomeReduzido
'O caixa %s - %s não pode ser excluído, pois existem movimentos de caixa vinculados a ele.
Public Const ERRO_CAIXA_VINCULADO_ECF = 13524 'Parâmetro: iCodigo, sNomeReduzido
'O caixa %s - %s não pode ser excluído, pois existem Emissoras de Cupom Fiscal vinculadas a ele.
Public Const ERRO_CAIXA_VINCULADO_SESSAO = 13525 'Parâmetros:iCodigo, sNomeReduzido
'O caixa %s - %s não pode ser excluído, pois existem Sessões vinculadas a ele.
Public Const ERRO_ALTERACAO_CAIXA_OUTRA_FILIAL = 13526 'Parâmetros: iCodigo, sNomeReduzido, iFilialEmpresa
'O Caixa %s - %s não pode ser alterado, pois pertence à filial %s.
Public Const ERRO_EXCLUSAO_CAIXA_OUTRA_FILIAL = 13527 'Parâmetros: iCodigo, sNomeReduzido, iFilialEmpresa
'O Caixa %s - %s não pode ser excluiído, pois pertence à filial %s.
Public Const ERRO_SIGLA_EXISTE = 13528 'SIGLA
'A sigla %s já existe
Public Const ERRO_ATUALIZACAO_PRODUTOSFILIAL1 = 13529 'Parâmetros: objLojaConfig.iTabelaPreco
'Erro na tentativa de atualizar registro na tabela ProdutosFilial com Tabela de Preço %s.
Public Const ERRO_OPERADOR_NAO_CADASTRADO = 13530 'Parametro: iCodigo
'O operador com o código %i não esta cadastrado
Public Const ERRO_LEITURA_OPERADOR2 = 13531
'Erro de leitura na tabela de Operadores.
Public Const ERRO_LEITURA_OPERADOR1 = 13532 'Parametro: sCodUsuario
'Erro de leitura do operador com o código de usuário %s.
Public Const ERRO_OPERADOR_USUARIO = 13533 'Parametros: iCodOperador,sCodUsuario
'O código de Operador %i correspondente ao Usuário %s no Bando de Dados não confere com o Operador %i da Tela.
Public Const ERRO_LOCK_OPERADOR = 13534
'Erro na tentativa de fazer 'lock' na tabela Operador.
Public Const ERRO_LEITURA_OPERADOR = 13535 'Parametro: iCodigo
'Erro de leitura do Operador com o código %i na tabela de operadores.
Public Const ERRO_ATUALIZACAO_OPERADOR = 13536 'Parametro: iCodOperador
'Erro na tentativa de atualizar o Operador %i a tabela Operador.
Public Const ERRO_INSERCAO_OPERADOR = 13537 'Parametro: iCodOperador
'Erro na tentiva de inserir o Operador %i na tabela Operador.
Public Const ERRO_USUARIO_OPERADOR_NAO_ALTERAVEL = 13538 'Parametro: iCodigo
'O Operador com o código %i nao esta cadastrado.
Public Const ERRO_EXCLUSAO_OPERADOR = 13539 ' Parametro: iCodOperador
'Erro na tentativa de excluir o operador %i na tabela Operador.
Public Const ERRO_OPERADOR_VINCULADO_BOLETO = 13540 'Parametro: icodigo
'Operador %i não pode ser excluído pois está vinculado a um registro na tabela Boleto.
Public Const ERRO_OPERADOR_VINCULADO_VALETICKET = 13541 'Parametro: icodigo
'Operador %i não pode ser excluído pois está vinculado a um registro na tabela ValeTicket.
Public Const ERRO_CATEGORIAPRODUTOITEM_EXISTE = 13542 'Parametro: CategoriaProdutoItem, CategoriaProduto
'O item %s e a categoria %s já foram definidas no grid.
Public Const ERRO_PRODUTO_CODBARRAS_NAO_PREENCHIDO = 13543
'O código de barras do produto não foi preenchido
Public Const ERRO_PRODUTO_REFERENCIA_NAO_PREENCHIDA = 13544
'A referência do produto não foi preenchida
Public Const ERRO_PRODUTO_SEM_TABELAPRECO_PADRAO = 13545
'Não existe preço cadastrado para o produto em questão
Public Const ERRO_PRECO_PRODUTO_NAO_CADASTRADO = 13546 'Parametro sCodProduto
'o produto %s não tem preço cadastrado
Public Const ERRO_GERENTE_NAO_CADASTRADO = 13547 'Parametro: iCodigo
'O gerente com o código %i não esta cadastrado.
Public Const ERRO_LEITURA_GERENTE2 = 13548
'Erro de leitura na tabela de Gerentes.
Public Const ERRO_LEITURA_GERENTE1 = 13549 'Parametro: sCodUsuario
'Erro de leitura do gerente com o código de usuário %s.
Public Const ERRO_GERENTE_USUARIO = 13550 'Parametros: iCodGerente,sCodUsuario
'O código de Gerente %i correspondente ao Usuário %s no Bando de Dados não confere com o Gerente %i da Tela.
Public Const ERRO_LOCK_GERENTE = 13551
'Erro na tentativa de fazer 'lock' na tabela Gerente.
Public Const ERRO_LEITURA_GERENTE = 13552 'Parametro: iCodigo
'Erro de leitura do Gerente com o código %i na tabela de gerentes.
Public Const ERRO_ATUALIZACAO_GERENTE = 13553 'Parametro: iCodGerente
'Erro na tentativa de atualizar o Gerente %i a tabela Gerente.
Public Const ERRO_INSERCAO_GERENTE = 13554 'Parametro: iCodGerente
'Erro na tentiva de inserir o Gerente %i na tabela Gerente.
Public Const ERRO_USUARIO_GERENTE_NAO_ALTERAVEL = 13555 'Parametro: iCodigo
'O Gerente com o código %i não esta cadastrado.
Public Const ERRO_EXCLUSAO_GERENTE = 13556 ' Parametro: iCodGerente
'Erro na tentativa de excluir o gerente %i na tabela Gerente.
Public Const ERRO_GERENTE_VINCULADO_SESSAO = 13557 'Parametro: icodigo
'Gerente %i não pode ser excluído pois está vinculado a um registro na tabela Sessão.
Public Const ERRO_GERENTE_VINCULADO_CUPOMFISCAL = 13558 'Parametro: sCodUsuario
'Gerente %i não pode ser excluído pois esta vinculado a um registro na tabela Cupom Fiscal.
Public Const ERRO_POS_NAO_CADASTRADO = 13559 'Parametro:sCodigo
'O POS %s não está cadastrado.
Public Const ERRO_LEITURA_POS = 13560 'Parametro:sCodigo
'Erro na leitura da tabela de POS. POS Código %s.
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
'O POS %s não pode ser excluído pois está vinculado à ECF
Public Const ERRO_LEITURA_BOLETO = 13568 'Sem Parâmetros
'Erro na leitura da tabela de Boleto.
Public Const ERRO_POS_VINCULADO_BOLETO = 13569 'Parametro:sCodigo
'O POS %s não pode ser excluído pois está vinculado à Boleto
Public Const ERRO_POS_OUTRA_FILIALEMPRESA = 13570 'Parametro:sCodigo, iFilialEmpresaPOS
'O POS código %s pertence a filial %i.
Public Const ERRO_LEITURA_REDE = 13571 'Parâmetro: iRede
'Erro na leitura da Rede %s na tabela de Redes
Public Const ERRO_REDE_NAO_ENCONTRADA = 13572 'Parâmetro: Rede
'A Rede %s não foi encontrada.
Public Const ERRO_INTERVALOVARIAVEL_EXISTE = 13573 'CondicoesPagto, IntervalosVariaveis
'A Condição de Pagamento %s com o intervalo %s já existem
Public Const ERRO_TAXA_NAO_PREENCHIDA = 13574
'Deve ser informada se a taxa de pagamento é à vista ou à prazo
Public Const ERRO_LEITURA_VALETICKET = 13575
'Erro na leitura da tabela de Vale Ticket
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_VALETICKET = 13576 'Parametro:sCodigo
'A Administradora %s não pode ser excluída pois está vinculada à um Vale Ticket
Public Const ERRO_LEITURA_FECHAMENTOBOLETOS = 13577
'Erro na leitura da tabela de Fechamento Boletos
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_FECHAMENTOBOLETOS = 13578 'Parametro:sCodigo
'A Administradora %s não pode ser excluída pois está vinculada à um Fechamento Boletos
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_BOLETO = 13579 'Parametro:sCodigo
'A administradora %s não pode ser excluída pois está vinculada à um Boleto
Public Const ERRO_LEITURA_BORDEROVALETICKET = 13580 'Parametro:sCodigo
'Erro na leitura da tabela de Bordero Vale Ticket
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_BORDEROVALETICKET = 13581 'Parametro:sCodigo
'A administradora não pode ser excluída pois está vinculada à um Bordero Vale Ticket
Public Const ERRO_LEITURA_BORDEROBOLETO = 13582
'Erro na leitura da tabela de Bordero Boleto
Public Const ERRO_ADMMEIOPAGTO_VINCULADO_BORDEROBOLETO = 13583 'Parametro:sCodigo
'A administradora não pode ser excluída pois está vinculada à um Bordero Boleto
Public Const ERRO_ADMMEIOPAGTO_NAO_CADASTRADO = 13584 'Parametro: sCodigo
'A Administradora %s não está cadastrada
Public Const ERRO_LOCK_ADMMEIOPAGTOCONDPAGTO = 13585 'Sem Parâmetros
'Erro na tentativa de fazer 'lock' na tabela AdmMeioPagtoCondPagto.
Public Const ERRO_EXCLUSAO_ADMMEIOPAGTO = 13586 'Parametro:sCodigo
'Erro na exclusão da Administradora %s.
Public Const ERRO_INSERCAO_ADMMEIOPAGTOCONDPAGTO = 13587 'Sem Parametros
'Erro na tentativa de inserção de um registro da tabela AdmMeioPagtoCondPagto
Public Const ERRO_ATUALIZACAO_ADMINISTRADORA = 13588 'Parametro: sCodigo
'Erro na tentativa de atualizar administradora %s
Public Const ERRO_EXCLUSAO_ADMMEIOPAGTOCONDPAGTO = 13589
'Erro na tentativa de exclusão de um registro da tabela AdmMeioPagtoCondPagto
Public Const ERRO_INSERCAO_ADMMEIOPAGTO = 13590 'Parametro:sCodigo
'Erro na tentativa de inserção de um registro da tabela AdmMeioPagto
Public Const ERRO_SELECIONAR_CONDPAGTO_2VEZES = 13591 'Sem Parametro
'Uma condição de pagamento não pode ser selecionada mais de duas vezes
Public Const ERRO_SELECIONAR_CONDPAGTO = 13592 'Sem Parametro
'Selecionar na lista a condição de pagamento
Public Const ERRO_TIPOMEIOPAGTO_NAO_PREENCHIDO = 13593 'Sem Parametro
'Preenchimento do Meio pagto é obrigatório
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
Public Const ERRO_LEITURA_FABRICANTEECF = 13599 'Parâmetro: iCodigo
'Ocorreu um erro na leitura da tabela de Fabricantes (FabricanteECF). Fabricante = %s
Public Const ERRO_LEITURA_CUPOMFISCAL = 13600 'Sem Parâmetros
'Ocorreu um erro na leitura da tabela de Cupons Fiscais (CupomFiscal).
Public Const ERRO_LEITURA_CHEQUE = 13601 'Sem parâmetros
'Ocorreu um erro na leitura da tabela de Cheques (Cheque).
Public Const ERRO_LOCK_ECF = 13602 'Sem parâmetros
'Ocorreu um erro ao tentar fazer o lock de um registro na tabela de ECF's (ECF).
Public Const ERRO_LOCK_FABRICANTEECF = 13603 'Sem parâmetros
'Ocorreu um erro ao tentar fazer o lock de um registro da tabela de Fabricantes(FabricanteECF).
Public Const ERRO_FABRICANTE_ECF_NAO_SELECIONADO = 13604 'Sem parâmetros
'O Fabricante da ECF não foi selecionado.
Public Const ERRO_FABRICANTE_ECF_NAO_CADASTRADO = 13605 'Parâmetros: iCodigo
'O Fabricante %s não está cadastrado.
Public Const ERRO_LEITURA_ECF1 = 13606 'Sem parâmetros
'Ocorreu um erro na leitura da tabela de ECF'S (ECF).
Public Const ERRO_LEITURA_FABRICANTEECF1 = 13607 'Sem parâmetros
'Ocorreu um erro na leitura da tabela de Fabricantes (FabricanteECF).
Public Const ERRO_POS_PERTENCE_OUTRA_FILIAL = 13608 'Parâmetros: sCodigo, objPOS.iFilialEmpresa
'O POS com código %s pertence à Filial %s e não pode ser usado nessa filial.
Public Const ERRO_CAIXA_ECF_NAO_SELECIONADO = 13609 'Sem parâmetros
'O Caixa da ECF não foi selecionado.
Public Const ERRO_INSERCAO_ECF = 13610 'Parâmetro: iCodigo
'Erro na inserção do ECF %s.
Public Const ERRO_ATUALIZACAO_ECF = 13611 'Parâmetro: iCodigo
'Erro na atualização do ECF %s.
Public Const ERRO_EXCLUSAO_ECF = 13612 'Parâmetro: iCodigo
'Ocorreu um na exclusão do ECF %s.
Public Const ERRO_ECF_NAO_CADASTRADO = 13613 'Parâmetro: iCodigo
'O ECF %s não está cadastrado.
Public Const ERRO_ALTERACAO_ECF_OUTRA_FILIAL = 13614 'Parâmetros: iCodigo, iFilialEmpresaECF
'O ECF %s não pode ser alterado, pois pertence à filial %s.
Public Const ERRO_EXCLUSAO_ECF_OUTRA_FILIAL = 13615 'Parâmetros: iCodigo, iFilialEmpresaECF
'O ECF %s não pode ser excluiído, pois pertence à filial %s.
Public Const ERRO_ECF_VINCULADO_VALETICKET = 13616 'Parâmetro: iCodigo
'O ECF com código %s não pode ser excluído, pois existem Tickets vinculados a ele.
Public Const ERRO_ECF_VINCULADO_BOLETO = 13617 'Parâmetro: iCodigo
'O ECF com código %s não pode ser excluído, pois existem Boletos vinculados a ele.
Public Const ERRO_ECF_VINCULADO_CUPOMFISCAL = 13618 'Parâmetro: iCodigo
'O ECF com código %s não pode ser excluído, pois existem Cupons Fiscais vinculados ele.
Public Const ERRO_ECF_VINCULADO_CHEQUE = 13619 'Parâmetro: iCodigo
'O ECF com código %s não pode ser excluído, pois existem Cheques vinculados a ele.
Public Const ERRO_LEITURA_CAIXAS1 = 13620 'Sem parâmetros
'Ocorreu um erro na leitura da tabela de caixas(Caixa).
Public Const ERRO_LEITURA_CAIXAS2 = 13621 'Parâmetro: sNomeReduzido
'Ocorreu um erro na leitura da tabela de caixas(Caixa). Caixa = %s



'CÓDIGOS DE AVISO - Reservado de 15400 a 15499
Public Const AVISO_EXCLUIR_CAIXA = 15400 'Parâmetro: iCodigo, sNomeReduzido
'Confirma exclusão do Caixa %s - %s ?
Public Const AVISO_CONFIRMA_EXCLUSAO_OPERADOR = 15401 'Parametro: sCodUsuario
'Confirma a exclusão do operador %s ?
Public Const AVISO_CONFIRMA_EXCLUSAO_GERENTE = 15402 'Parametro: sCodUsuario
'Confirma a exclusão do gerente %s ?
Public Const AVISO_EXCLUIR_POS = 15403 'Parametro:sCodigo
'Confirma a exclusão do POS com código %s ?
Public Const AVISO_EXCLUIR_ADMINISTRADORA = 15404 'Parametro:sCodigo
'Confirma exclusão da administradora com código %s?
Public Const AVISO_DESEJA_CRIAR_CAIXA = 15405 'Parâmetro: iCodigo ou sNomeReduzido
'O Caixa %s não está cadastrado. Deseja criá-lo agora?
Public Const AVISO_DESEJA_CRIAR_POS = 15406 'Parâmetro: sCodigo
'O POS com código %s não está cadastrado. Deseja criá-lo agora?
Public Const AVISO_EXCLUIR_ECF = 15407 'Parâmetro: iCodigo
'Confirma exclusão do ECF %s?


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



