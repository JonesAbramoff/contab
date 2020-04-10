VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMesEstAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'criada para atender às chamadas de EXT_FUNC nos relatorios atraves de pesquisa, principalmente, nas tabelas SldMesEstAlm, SldDiaEstAlm
'existem grupos de funcoes:
'1)que precisa do saldo inicial no periodo



''IMPORTANTE: NAO ESQUECA DE TRATAR DATA_NULA COMO POSSIVEL DATA INICIAL OU FINAL
'
    
'Identificação
Private dtUltDataInicial As Date
Private dtUltDataFinal As Date
Private iUltAlmoxarifado As Integer
Private sUltProduto As String 'último produto pesquisado. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" abaixo forem válidos.
Private iUltApropriacaoCusto As Integer
Private iUltAnoEstApurado As Integer
Private iUltMesEstApurado As Integer

'Valores armazenados para o último Produto/Almoxarifado pesquisado
Private dUltQuantInicial As Double
Private dUltValorInicial As Double

Private alComando(0 To 1) As Long

Private Sub Class_Terminate()

    If alComando(0) <> 0 Then Call Comando_Fechar(alComando(0))
    If alComando(1) <> 0 Then Call Comando_Fechar(alComando(1))
    
End Sub


Private Function ObterRelQtdSldIni(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'Para obter os valores a partir de sldmesestalm/slddiaestalm e torná-los disponíveis para as outras funções

Dim lErro As Long
Dim dQuantInicial As Double
Dim dValorInicial As Double

On Error GoTo Erro_ObterRelQtdSldIni

    lErro = CF("Produto_ObterRelQtdSldIni",sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, dQuantInicial, dValorInicial, alComando, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)
    If lErro <> SUCESSO Then Error 55070
    
    sUltProduto = sProduto
    iUltAlmoxarifado = iAlmoxarifado
    dtUltDataInicial = dtDataInicial
    dtUltDataFinal = dtDataFinal
    iUltApropriacaoCusto = iApropriacaoCusto
    iUltAnoEstApurado = iAnoEstApurado
    iUltMesEstApurado = iMesEstApurado
    
    dUltQuantInicial = dQuantInicial
    dUltValorInicial = dValorInicial
    
    ObterRelQtdSldIni = SUCESSO
    
    Exit Function
    
Erro_ObterRelQtdSldIni:

    ObterRelQtdSldIni = Err
    
    Select Case Err

        Case 55070
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153252)

    End Select

    Exit Function
    
End Function

Public Function ObterQuantInicial1(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, dQuantInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'Para obter a qtde inicial de um produto/almoxarifado até a datainicial

Dim lErro As Long

On Error GoTo Erro_ObterQuantInicial1
    
    If sProduto <> sUltProduto Or iAlmoxarifado <> iUltAlmoxarifado Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iApropriacaoCusto <> iUltApropriacaoCusto Or iAnoEstApurado <> iUltAnoEstApurado Or iMesEstApurado <> iUltMesEstApurado Then
        
        lErro = ObterRelQtdSldIni(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)
        If lErro <> SUCESSO Then Error 55077
    
    End If
    
    dQuantInicial = dUltQuantInicial
    
    ObterQuantInicial1 = SUCESSO
    
    Exit Function
    
Erro_ObterQuantInicial1:

    ObterQuantInicial1 = Err
    
    Select Case Err

        Case 55077
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153253)

    End Select

    Exit Function
    
End Function

Public Function ObterValorInicial1(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, dValorInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'Para obter o saldo (valor) inicial de um produto/almoxarifado até a datainicial

Dim lErro As Long

On Error GoTo Erro_ObterValorInicial1
    
    If sProduto <> sUltProduto Or iAlmoxarifado <> iUltAlmoxarifado Or dtDataInicial <> dtUltDataInicial Or dtDataFinal <> dtUltDataFinal Or iApropriacaoCusto <> iUltApropriacaoCusto Or iAnoEstApurado <> iUltAnoEstApurado Or iMesEstApurado <> iUltMesEstApurado Then
        
        lErro = ObterRelQtdSldIni(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)
        If lErro <> SUCESSO Then Error 55079
    
    End If
    
    dValorInicial = dUltValorInicial
    
    ObterValorInicial1 = SUCESSO
    
    Exit Function
    
Erro_ObterValorInicial1:

    ObterValorInicial1 = Err
    
    Select Case Err

        Case 55079
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153254)

    End Select

    Exit Function
    
End Function


