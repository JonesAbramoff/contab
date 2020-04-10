VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMesEstFil3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'criada para atender às chamadas de EXT_FUNC nos relatorios atraves de pesquisa, principalmente, nas tabelas SldMesEst, SldDiaEst
'existem grupos de funcoes:
'1)que precisa do saldo inicial no periodo

''IMPORTANTE: NAO ESQUECA DE TRATAR DATA_NULA COMO POSSIVEL DATA INICIAL OU FINAL
'
    
'Identificação
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date
'Private iUltAlmoxarifado As Integer
Private sUltProduto As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.
Private iUltApropriacaoCusto As Integer
Private iUltAnoEstApurado As Integer
Private iUltMesEstApurado As Integer
Private iUltFilialEmpresa As Integer
Private iUltNumEscaninho As Integer

'Valores armazenados para o último Produto/Filial pesquisado
Private dUltQuantInicial As Double
Private dUltValorInicial As Double

Private alComando(0 To 1) As Long

Private Sub Class_Terminate()

    If alComando(0) <> 0 Then Call Comando_Fechar(alComando(0))
    If alComando(1) <> 0 Then Call Comando_Fechar(alComando(1))
    
End Sub

Private Function ObterRelQtdSldIni(sProduto As String, iFilialEmpresa As Integer, dtDataInicial As Date, dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'Para obter os valores a partir de sldmesest/slddiaest e torná-los disponíveis para as outras funções

Dim lErro As Long
Dim dQuantInicial As Double
Dim dValorInicial As Double

On Error GoTo Erro_ObterRelQtdSldIni
    
    lErro = CF("Produto_ObterRelQtdSldIniFilTercEsc", sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, dQuantInicial, dValorInicial, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)
    If lErro <> SUCESSO Then gError 93743
    
    sUltProduto = sProduto
    iUltFilialEmpresa = iFilialEmpresa
    dtUltDataInicial = dtDataInicial
    dtUltDataFinal = dtDataFinal
    iUltApropriacaoCusto = iApropriacaoCusto
    iUltAnoEstApurado = iAnoEstApurado
    iUltMesEstApurado = iMesEstApurado
    iUltNumEscaninho = iNumEscaninho
    
    dUltQuantInicial = dQuantInicial
    dUltValorInicial = dValorInicial
    
    ObterRelQtdSldIni = SUCESSO
    
    Exit Function
    
Erro_ObterRelQtdSldIni:

    ObterRelQtdSldIni = gErr
    
    Select Case gErr

        Case 93743
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153262)

    End Select

    Exit Function
    
End Function

Public Function ObterQuantInicial1(sProduto As String, iFilialEmpresa As Integer, dtDataInicial As Date, dtDataFinal As Date, dQuantInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'Para obter a qtde inicial de um produto/filial até a datainicial

Dim lErro As Long

On Error GoTo Erro_ObterQuantInicial1
    
    If sProduto <> sUltProduto Or iFilialEmpresa <> iUltFilialEmpresa Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iApropriacaoCusto <> iUltApropriacaoCusto Or iAnoEstApurado <> iUltAnoEstApurado Or iMesEstApurado <> iUltMesEstApurado Or iNumEscaninho <> iUltNumEscaninho Then
        
        lErro = ObterRelQtdSldIni(sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)
        If lErro <> SUCESSO Then gError 93744
    
    End If
    
    dQuantInicial = dUltQuantInicial
    
    ObterQuantInicial1 = SUCESSO
    
    Exit Function
    
Erro_ObterQuantInicial1:

    ObterQuantInicial1 = gErr
    
    Select Case gErr

        Case 93744
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153263)

    End Select

    Exit Function
    
End Function

Public Function ObterValorInicial1(sProduto As String, iFilialEmpresa As Integer, dtDataInicial As Date, dtDataFinal As Date, dValorInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'Para obter o saldo (valor) inicial de um produto/filial até a datainicial

Dim lErro As Long

On Error GoTo Erro_ObterValorInicial1
    
    If sProduto <> sUltProduto Or iFilialEmpresa <> iUltFilialEmpresa Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iApropriacaoCusto <> iUltApropriacaoCusto Or iAnoEstApurado <> iUltAnoEstApurado Or iMesEstApurado <> iUltMesEstApurado Or iNumEscaninho <> iUltNumEscaninho Then
        
        lErro = ObterRelQtdSldIni(sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado, iNumEscaninho)
        If lErro <> SUCESSO Then gError 93745
    
    End If
    
    dValorInicial = dUltValorInicial
    
    ObterValorInicial1 = SUCESSO
    
    Exit Function
    
Erro_ObterValorInicial1:

    ObterValorInicial1 = gErr
    
    Select Case gErr

        Case 93745
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153264)

    End Select

    Exit Function
    
End Function

