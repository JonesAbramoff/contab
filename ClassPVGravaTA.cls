VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ItensOV_Grava_Customizado(ByVal objItensOV As ClassItemOV) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objItensOVTA As ClassItemOVTA
Dim lNumIntDoc As Long, sTabela As String

On Error GoTo Erro_ItensOV_Grava_Customizado

    If Not (objItensOV.objInfoUsu Is Nothing) Then
    
        Set objItensOVTA = objItensOV.objInfoUsu

        'Abre Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 210447
        Next
    
        If objItensOV.iHistorico <> MARCADO Then
            sTabela = "ItensOrcamentoVenda"
        Else
            sTabela = "ItensOrcamentoVendaHist"
        End If
        
        'Lê os itens do orçamento passado
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM " & sTabela & " WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItensOV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 210448
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 210449
    
        If lErro = AD_SQL_SEM_DADOS Then gError 210450
            
        'Altera a Observação
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE " & sTabela & " SET PrecoBase=?, PercAcresc=?, Acrescimo = ?", alComando(1), objItensOVTA.dPrecoBase, objItensOVTA.dPercAcresc, objItensOVTA.dAcrescimo)
        If lErro <> AD_SQL_SUCESSO Then gError 210451

        'Fecha comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next

    End If

    ItensOV_Grava_Customizado = SUCESSO

    Exit Function

Erro_ItensOV_Grava_Customizado:

    ItensOV_Grava_Customizado = gErr

    Select Case gErr

        Case 210447
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 210448, 210449
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr)

        Case 210450
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_VENDA_SEM_ITENS", gErr)
            
        Case 210451
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSOV", gErr, objItensOV.lCodOrcamento)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210452)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

