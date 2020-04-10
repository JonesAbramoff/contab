VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelInventarioTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"ClassRelInventarioTerc"
Option Explicit

Dim iFilialEmpresaAnt As Integer
Dim dQuantidadeTotalAnt As Double
Dim iTipoTercAnt As Integer
Dim lCodTercAnt  As Long
Dim iFilialTercAnt  As Integer
Dim sProdutoAnt  As String
Dim dtDataInicialAnt As Date
Dim iCodEscaninhoAnt  As Integer

Public Function ObterSaldoInicial(dQuantidadeTotal As Double, dtDataInicial As Date, iFilialEmpresa As Integer, iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer, sProduto As String, iCodEscaninho As Integer) As Long
'Obtem o saldo inicial

Dim lErro As Long
Dim dtDataInventario As Date
Dim sCampoEscaninho As String
Dim objEstTerc As New ClassEstoqueTerc

On Error GoTo Erro_ObterSaldoInicial

    'Se a chave é igual a anterior, a quantidade também é, logo não precisa calcular
    If dtDataInicial = dtDataInicialAnt And iFilialEmpresa = iFilialEmpresaAnt And iTipoTerc = iTipoTercAnt And lCodTerc = lCodTercAnt And iFilialTerc = iFilialTercAnt And sProduto = sProdutoAnt And iCodEscaninho = iCodEscaninhoAnt Then
        dQuantidadeTotal = dQuantidadeTotalAnt
        Exit Function
    End If
    
'    'Obtem qual campo é <> 0  em TiposDeMovimentosEstoque para codigo escaninho passado
'    lErro = ObterCampoEscaninho(sCampoEscaninho, iCodEscaninho)
'    If lErro <> SUCESSO Then gError 110260
'
'    lErro = CF("InventarioTercProd_Le_MaiorDataAnterior", dtDataInventario, dtDataInicial, iFilialEmpresa, iTipoTerc, lCodTerc, iFilialTerc, sProduto, iCodEscaninho)
'    If lErro <> SUCESSO Then gError 110261
'
'    lErro = CF("MovimentoEstoque_Le_SaldoInicial", dQuantidadeTotal, dtDataInventario, dtDataInicial, iFilialEmpresa, iTipoTerc, lCodTerc, iFilialTerc, sProduto, sCampoEscaninho, iCodEscaninho)
'    If lErro <> SUCESSO Then gError 110262

    objEstTerc.iEscaninho = iCodEscaninho
    objEstTerc.iFilial = iFilialTerc
    objEstTerc.iTipoCliForn = iTipoTerc
    objEstTerc.lCliForn = lCodTerc
    objEstTerc.sProduto = sProduto
    objEstTerc.dtData = dtDataInicial
    objEstTerc.iFilialEmpresa = iFilialEmpresa

    lErro = CF("EstoqueTerc_Le_Saldo", objEstTerc)
    If lErro <> SUCESSO Then gError 110262
    
    dQuantidadeTotalAnt = objEstTerc.dQuantidade
    dtDataInicialAnt = dtDataInicial
    iFilialEmpresaAnt = iFilialEmpresa
    iTipoTercAnt = iTipoTerc
    lCodTercAnt = lCodTerc
    iFilialTercAnt = iFilialTerc
    sProdutoAnt = sProduto
    iCodEscaninhoAnt = iCodEscaninho
    
    ObterSaldoInicial = SUCESSO

    Exit Function
    
Erro_ObterSaldoInicial:

    ObterSaldoInicial = gErr

    Select Case gErr
    
        Case 110260 To 110262

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153174)

    End Select
    
    Exit Function

End Function

Private Function ObterCampoEscaninho(sCampo As String, iCodEscaninho As Integer) As Long
'Obtem o campo <> 0 em TiposDeMovimentosEstoque para o código escaninho passado

On Error GoTo Erro_ObterCampoEscaninho

    Select Case iCodEscaninho
    
        Case ESCANINHO_CONSERTO_NOSSO
            sCampo = "AtualizaConserto"
        
        Case ESCANINHO_CONSIG_NOSSO
            sCampo = "AtualizaConsig"
        
        Case ESCANINHO_DEMO_NOSSO
            sCampo = "AtualizaDemo"
        
        Case ESCANINHO_OUTROS_NOSSO
            sCampo = "AtualizaOutras"
        
        Case ESCANINHO_BENEF_NOSSO
            sCampo = "AtualizaBenef"
    
        Case ESCANINHO_CONSERTO_3
            sCampo = "AtualizaCustoConserto3"
        
        Case ESCANINHO_CONSIG_3
            sCampo = "AtualizaCustoConsig3"
        
        Case ESCANINHO_DEMO_3
            sCampo = "AtualizaCustoDemo3"
        
        Case ESCANINHO_OUTROS_3
            sCampo = "AtualizaCustoOutras3"
        
        Case ESCANINHO_BENEF_3
            sCampo = "AtualizaCustoBenef3"
            
        Case Else
            gError 110280
    
    End Select

    ObterCampoEscaninho = SUCESSO

    Exit Function

Erro_ObterCampoEscaninho:

    ObterCampoEscaninho = gErr

    Select Case gErr
    
        Case 110280

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153175)

    End Select
    
    Exit Function

End Function

