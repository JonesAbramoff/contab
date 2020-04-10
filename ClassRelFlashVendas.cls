VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFlashVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long


'Declaração de um Obj Glogal ..
Dim gobjUltRelFlashVendas As New ClassFlashVendas

Function Obtem_NumVendas_IntervHorasAux(dNumVendas As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
'Função que Retorna o Numero de Vendas para uma determinada data e Horario

Dim lErro As Long
Dim objRelFlashVendas As ClassFlashVendas

On Error GoTo Erro_Obtem_NumVendas_IntervHorasAux

    'Verifica se os Valores guardados no Obj são diferentes aos passados por parâmetro
    If sProduto <> gobjUltRelFlashVendas.sProduto Or iFilialEmpresa <> gobjUltRelFlashVendas.iFilialEmpresa Or dtData <> gobjUltRelFlashVendas.dtData Or dHoraIni <> gobjUltRelFlashVendas.dHoraIni Or dHoraFim <> gobjUltRelFlashVendas.dHoraFim Or iCaixaIni <> gobjUltRelFlashVendas.iCaixaIni Or iCaixaFim <> gobjUltRelFlashVendas.iCaixaFim Then

        Set objRelFlashVendas = New ClassFlashVendas

        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
        objRelFlashVendas.dHoraFim = dHoraFim
        objRelFlashVendas.dHoraIni = dHoraIni
        objRelFlashVendas.dtData = dtData
        objRelFlashVendas.iCaixaFim = iCaixaFim
        objRelFlashVendas.iCaixaIni = iCaixaIni
        objRelFlashVendas.iFilialEmpresa = iFilialEmpresa
        objRelFlashVendas.iCaixaIni = iCaixaIni
        objRelFlashVendas.sProduto = sProduto

        'Função que lê as Estatisticas (Numero de vendas ) que serão utilizadas pelo Relatório
        lErro = RelFlashVendas_Le_Estatisticas_Hora(objRelFlashVendas)
        If lErro <> SUCESSO Then gError 113261

        'Aponta o Obj Global para o Obj Local a Função
        Set gobjUltRelFlashVendas = objRelFlashVendas

    End If

    'Guarda o Numero de Vandas para a Data e o Intervalo de Hora Especificado
    dNumVendas = gobjUltRelFlashVendas.lNumVendas

    Obtem_NumVendas_IntervHorasAux = SUCESSO

    Exit Function

Erro_Obtem_NumVendas_IntervHorasAux:

    Obtem_NumVendas_IntervHorasAux = gErr

        Select Case gErr

        Case 113261

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153148)

    End Select

    Exit Function

End Function


Function Obtem_QuantVend_IntervHorasAux(dQuantVendida As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
'Função que Retorna a Quantidade vendida para uma determinada data e Horario

Dim lErro As Long
Dim objRelFlashVendas As ClassFlashVendas

On Error GoTo Erro_Obtem_QuantVend_IntervHorasAux

    'Verifica se os Valores guardados no Obj são diferentes aos passados por parâmetro
    If sProduto <> gobjUltRelFlashVendas.sProduto Or iFilialEmpresa <> gobjUltRelFlashVendas.iFilialEmpresa Or dtData <> gobjUltRelFlashVendas.dtData Or dHoraIni <> gobjUltRelFlashVendas.dHoraIni Or dHoraFim <> gobjUltRelFlashVendas.dHoraFim Or iCaixaIni <> gobjUltRelFlashVendas.iCaixaIni Or iCaixaFim <> gobjUltRelFlashVendas.iCaixaFim Then

        Set objRelFlashVendas = New ClassFlashVendas

        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
        objRelFlashVendas.dHoraFim = dHoraFim
        objRelFlashVendas.dHoraIni = dHoraIni
        objRelFlashVendas.dtData = dtData
        objRelFlashVendas.iCaixaFim = iCaixaFim
        objRelFlashVendas.iCaixaIni = iCaixaIni
        objRelFlashVendas.iFilialEmpresa = iFilialEmpresa
        objRelFlashVendas.iCaixaIni = iCaixaIni
        objRelFlashVendas.sProduto = sProduto

        'Função que lê as Estatisticas (Numero de vendas ) que serão utilizadas pelo Relatório
        lErro = RelFlashVendas_Le_Estatisticas_Hora(objRelFlashVendas)
        If lErro <> SUCESSO Then gError 113285

        'Aponta o Obj Global para o Obj Local a Função
        Set gobjUltRelFlashVendas = objRelFlashVendas

    End If

    'Guarda a Quantidade de vendas
    dQuantVendida = gobjUltRelFlashVendas.dQuantVendida

    Obtem_QuantVend_IntervHorasAux = SUCESSO

    Exit Function

Erro_Obtem_QuantVend_IntervHorasAux:

    Obtem_QuantVend_IntervHorasAux = gErr

        Select Case gErr

        Case 113285

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153149)

    End Select

    Exit Function

End Function

Function Obtem_VlrVend_IntervHorasAux(dValorVendido As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
'Função que Retorna o Valor vendido  para uma determinada data e Horario

Dim lErro As Long
Dim objRelFlashVendas As ClassFlashVendas

On Error GoTo Erro_Obtem_VlrVend_IntervHorasAux

    'Verifica se os Valores guardados no Obj são diferentes aos passados por parâmetro
    If sProduto <> gobjUltRelFlashVendas.sProduto Or iFilialEmpresa <> gobjUltRelFlashVendas.iFilialEmpresa Or dtData <> gobjUltRelFlashVendas.dtData Or dHoraIni <> gobjUltRelFlashVendas.dHoraIni Or dHoraFim <> gobjUltRelFlashVendas.dHoraFim Or iCaixaIni <> gobjUltRelFlashVendas.iCaixaIni Or iCaixaFim <> gobjUltRelFlashVendas.iCaixaFim Then

        Set objRelFlashVendas = New ClassFlashVendas

        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
        objRelFlashVendas.dHoraFim = dHoraFim
        objRelFlashVendas.dHoraIni = dHoraIni
        objRelFlashVendas.dtData = dtData
        objRelFlashVendas.iCaixaFim = iCaixaFim
        objRelFlashVendas.iCaixaIni = iCaixaIni
        objRelFlashVendas.iFilialEmpresa = iFilialEmpresa
        objRelFlashVendas.iCaixaIni = iCaixaIni
        objRelFlashVendas.sProduto = sProduto

        'Função que lê as Estatisticas (Numero de vendas ) que serão utilizadas pelo Relatório
        lErro = RelFlashVendas_Le_Estatisticas_Hora(objRelFlashVendas)
        If lErro <> SUCESSO Then gError 113286

        'Aponta o Obj Global para o Obj Local a Função
        Set gobjUltRelFlashVendas = objRelFlashVendas

    End If

    'Guarda o Valor vendido
    dValorVendido = gobjUltRelFlashVendas.dValorVendido

    Obtem_VlrVend_IntervHorasAux = SUCESSO

    Exit Function

Erro_Obtem_VlrVend_IntervHorasAux:

    Obtem_VlrVend_IntervHorasAux = gErr

        Select Case gErr

        Case 113286

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153150)

    End Select

    Exit Function

End Function


Function RelFlashVendas_Le_Estatisticas_Hora(ByVal objRelFlashVendas As ClassFlashVendas) As Long
'Função que monta o Select que será executado posteriormente

Dim lErro As Long
Dim sSelect As String
Dim tRelFlashVendas As TypeRelFlashVendasVar
Dim lComando As Long

On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 113262

    'Função que Monta o select
    lErro = RelFlashVendas_Le_Estatisticas_Hora1(objRelFlashVendas, sSelect)
    If lErro <> SUCESSO Then gError 113263

    'Função que Prepara a Parte Fixa do Select
    lErro = RelFlashVendas_Le_Estatisticas_Hora2(lComando, sSelect, tRelFlashVendas)
    If lErro <> SUCESSO Then gError 113264

    'Prepara os parâmetros que variam de acordo com a seleção do usuário
    lErro = RelFlashVendas_Le_Estatisticas_Hora3(lComando, tRelFlashVendas, objRelFlashVendas, sSelect)
    If lErro <> SUCESSO Then gError 113265

    'Função que Processa os Registros Retormados Pelo Select
    lErro = RelFlashVendas_Le_Estatisticas_Hora4(lComando, tRelFlashVendas, objRelFlashVendas, sSelect)
    If lErro <> SUCESSO Then gError 113266

    'fecha o comando
    Call Comando_Fechar(lComando)

    RelFlashVendas_Le_Estatisticas_Hora = SUCESSO

    Exit Function

Erro_RelFlashVendas_Le_Estatisticas_Hora:

    RelFlashVendas_Le_Estatisticas_Hora = gErr

    Select Case gErr

        Case 113262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 113263 To 113266

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153151)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelFlashVendas_Le_Estatisticas_Hora1(ByVal objRelFlashVendas As ClassFlashVendas, sSelect As String) As Long
'Função que Guarda na String sSelect o comando que será utilizada para ler os registro em SldDiaFat

Dim lErro As Long
Dim sFrom As String
Dim sWhere As String

On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora1

    'select fixo
    sSelect = "SELECT  ItensCupomFiscal.UnidadeMed , COUNT(ItensCupomFiscal.UnidadeMed)AS NumUnidadeMed , SUM(ItensCupomFiscal.Quantidade)AS NumQuantidade , SUM(ItensCupomFiscal.Quantidade * ItensCupomFiscal.PrecoUnitario) AS Valor "

    'From Fixo
    sFrom = " FROM  CupomFiscal , ItensCupomFiscal"

    'Cláusula Where
    sWhere = " WHERE  CupomFiscal.NumIntDoc = ItensCupomFiscal.NumIntCupom AND ItensCupomFiscal.Tipo = 1 AND CupomFiscal.Tipo = 1 AND ItensCupomFiscal.Produto = ? AND CupomFiscal.HoraEmissao >= ? AND CupomFiscal.HoraEmissao <= ? AND CupomFiscal.DataEmissao = ? "

    'Verifica se o Filtro utilizado vai ser por filial ou não
    If objRelFlashVendas.iFilialEmpresa <> EMPRESA_TODA Then

        sWhere = sWhere & " AND CupomFiscal.FilialEmpresa = ItensCupomFiscal.FilialEmpresa AND ItensCupomFiscal.FilialEmpresa = ? "

    End If
    
'    'Verifica se o Filtro utilizado vai ser por filial ou não
'    If objRelFlashVendas.iFilialEmpresa <> EMPRESA_TODA Then
'
'        sWhere = sWhere & " AND CupomFiscal.FilialEmpresa = ItensCupomFiscal.FilialEmpresa AND ItensCupomFiscal.FilialEmpresa = ? "
'
'    End If
    
    'Verifica se o Filtro utilizado vai ser por Caixa DE ou Não
    If objRelFlashVendas.iCaixaIni <> CODIGO_NAO_PREENCHIDO Then

        sWhere = sWhere & " AND CupomFiscal.Caixa >= ? "

    End If
    
    
    'Verifica se o Filtro utilizado vai ser por Caixa ATE  ou Não
    If objRelFlashVendas.iCaixaFim <> CODIGO_NAO_PREENCHIDO Then

        sWhere = sWhere & " AND CupomFiscal.Caixa <= ? "

    End If
    
    'Adciona o group By
    sWhere = sWhere & "Group By (ItensCupomFiscal.Produto),( ItensCupomFiscal.UnidadeMed)"
    
    sSelect = sSelect & sFrom & sWhere

    RelFlashVendas_Le_Estatisticas_Hora1 = SUCESSO

    Exit Function

Erro_RelFlashVendas_Le_Estatisticas_Hora1:

    RelFlashVendas_Le_Estatisticas_Hora1 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153152)

    End Select

    Exit Function

End Function

Private Function RelFlashVendas_Le_Estatisticas_Hora2(ByVal lComando As Long, sSelect As String, tRelFlashVendas As TypeRelFlashVendasVar) As Long
'Função que Bind o as variáveis que serão recebidas pelo select

Dim lErro As Long

On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora2

    With tRelFlashVendas

        .vsUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        
        lErro = Comando_PrepararInt(lComando, sSelect)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113267

        .vsUMVenda = CStr(.vsUMVenda)
        lErro = Comando_BindVarInt(lComando, .vsUMVenda)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113268
    
        .vlNumVendas = CLng(.vlNumVendas)
        lErro = Comando_BindVarInt(lComando, .vlNumVendas)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113282

    
        .vdQuantVendas = CDbl(.vdQuantVendas)
        lErro = Comando_BindVarInt(lComando, .vdQuantVendas)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113269

        .vdValorVendido = CDbl(.vdValorVendido)
        lErro = Comando_BindVarInt(lComando, .vdValorVendido)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113270

        
    End With

    RelFlashVendas_Le_Estatisticas_Hora2 = SUCESSO

    Exit Function

Erro_RelFlashVendas_Le_Estatisticas_Hora2:

    RelFlashVendas_Le_Estatisticas_Hora2 = gErr

    Select Case gErr

        Case 113267 To 113270, 113282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153153)

    End Select

    Exit Function

End Function

Private Function RelFlashVendas_Le_Estatisticas_Hora3(ByVal lComando As Long, tRelFlashVendas As TypeRelFlashVendasVar, ByVal objRelFlashVendas As ClassFlashVendas, sSelect As String)
'Função que Bind os filtros passados pelo usuário

Dim lErro As Long

On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora3

    
    tRelFlashVendas.vsProduto = CStr(objRelFlashVendas.sProduto)
    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 113271
    
    tRelFlashVendas.vdHoraIni = CDbl(objRelFlashVendas.dHoraIni)
    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vdHoraIni)
    If (lErro <> AD_SQL_SUCESSO) Then gError 113272

    tRelFlashVendas.vdHoraFim = CDbl(objRelFlashVendas.dHoraFim)
    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vdHoraFim)
    If (lErro <> AD_SQL_SUCESSO) Then gError 113273

    tRelFlashVendas.vdtData = CDate(objRelFlashVendas.dtData)
    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vdtData)
    If (lErro <> AD_SQL_SUCESSO) Then gError 113274


    'Verifica se o filtro é por filial empresa
    If objRelFlashVendas.iFilialEmpresa <> EMPRESA_TODA Then

        tRelFlashVendas.viFilialEmpresa = CInt(giFilialEmpresa)
        lErro = Comando_BindVarInt(lComando, tRelFlashVendas.viFilialEmpresa)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113275


    End If

    'Verifica se o filtro também será por Caixa
    If objRelFlashVendas.iCaixaIni <> CODIGO_NAO_PREENCHIDO Then

        tRelFlashVendas.viCaixaIni = CInt(objRelFlashVendas.iCaixaIni)
        lErro = Comando_BindVarInt(lComando, tRelFlashVendas.viCaixaIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113276

    End If

    'Verifica se o filtro também será por Caixa
    If objRelFlashVendas.iCaixaFim <> CODIGO_NAO_PREENCHIDO Then

        tRelFlashVendas.viCaixaFim = CInt(objRelFlashVendas.iCaixaFim)
        lErro = Comando_BindVarInt(lComando, tRelFlashVendas.viCaixaFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 113277

    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 113278

    RelFlashVendas_Le_Estatisticas_Hora3 = SUCESSO

    Exit Function

Erro_RelFlashVendas_Le_Estatisticas_Hora3:

    RelFlashVendas_Le_Estatisticas_Hora3 = gErr

    Select Case gErr

        Case 113271 To 113278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153154)

    End Select

    Exit Function

End Function

Private Function RelFlashVendas_Le_Estatisticas_Hora4(ByVal lComando As Long, tRelFlashVendas As TypeRelFlashVendasVar, ByVal objRelFlashVendas As ClassFlashVendas, sSelect As String) As Long
'Busca no Banco de dados os Calculos Referentes a cada Produto vendido entre um determinado período de tempo

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFatorConv As Double

On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113279

    If lErro = AD_SQL_SEM_DADOS Then gError 113280
    
    'Atribui o Codigo do Produto em Questão para a Leitura no Banco de Dados
    objProduto.sCodigo = objRelFlashVendas.sProduto

    'Se Encontrou então Verifica qual é a unidade de venda do produto em questão
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 113281
    
    Do While lErro = AD_SQL_SUCESSO

        With objRelFlashVendas

            'acumular o Numero de Vendas
            .lNumVendas = .lNumVendas + tRelFlashVendas.vlNumVendas
            
            'Verifica se a unidade de venda do produto é diferente da unidade na Tabela de itens de Cupom
            If objProduto.sSiglaUMVenda <> tRelFlashVendas.vsUMVenda Then
                'Usar a Função que converte para unidade de venda
                lErro = CF("UM_CONVERSAO_REL", objProduto.iClasseUM, tRelFlashVendas.vsUMVenda, objProduto.sSiglaUMVenda, dFatorConv)
                If lErro <> SUCESSO Then gError 113283
                
                tRelFlashVendas.vdQuantVendas = tRelFlashVendas.vdQuantVendas * dFatorConv
            
            End If
            
            .sProduto = tRelFlashVendas.vsProduto
            .dQuantVendida = .dQuantVendida + tRelFlashVendas.vdQuantVendas
            .dValorVendido = .dValorVendido + tRelFlashVendas.vdValorVendido
        
        End With
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113284

    Loop

    RelFlashVendas_Le_Estatisticas_Hora4 = SUCESSO

    Exit Function

Erro_RelFlashVendas_Le_Estatisticas_Hora4:

    RelFlashVendas_Le_Estatisticas_Hora4 = gErr

    Select Case gErr

        Case 113279, 113284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)

        Case 113280, 113283
            'Só desvia o Código sem Msg
        
        Case 113280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr, sSelect)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153155)

    End Select

    Exit Function

End Function

