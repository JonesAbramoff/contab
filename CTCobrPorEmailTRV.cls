VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTCobrPorEmailTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get TipoDoc(ByVal objCT As Object) As Object
     Set TipoDoc = objCT.objUserControl.Controls("TipoDoc")
End Property

Function CE_Move_TabSelecao_Memoria(ByVal objCT As Object, ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objCobrSelTRV As New ClassCobrPorEmailSelTRV

On Error GoTo Erro_CE_Move_TabSelecao_Memoria

    Set objCobrancaEmailSel.objInfoUsu = objCobrSelTRV

    If objCT.giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Or objCT.giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
    
        For iLinha = 0 To TipoDoc(objCT).ListCount - 1
            
            If TipoDoc(objCT).Selected(iLinha) Then
                objCobrSelTRV.colTiposDoc.Add TipoDoc(objCT).ItemData(iLinha)
            End If
        
        Next
    
    End If
   
    CE_Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_CE_Move_TabSelecao_Memoria:

    CE_Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187027)

    End Select

    Exit Function

End Function

Function CE_BotaoEmail_Click(ByVal objCT As Object, iGeraRelac As Integer)

    If objCT.giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Or objCT.giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
        If glEmpresa < 3 Then iGeraRelac = MARCADO
    End If

End Function

Public Sub TipoDoc_Click(ByVal objCT As Object)
    objCT.iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Function CE_Calcula_Mnemonico_Cust(objMnemonicoValor As ClassMnemonicoValor, ByVal iTipoTela As Integer, ByVal objDoc As Object, sValorTexto As String) As Long

Dim lErro As Long
Dim sNumTitulo As String
Dim sConteudo As String
Dim lFornOver As Long

On Error GoTo Erro_CE_Calcula_Mnemonico_Cust

    lErro = CF("TRVConfig_Le", TRVCONFIG_CLIENTE_OVER, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lFornOver = StrParaLong(sConteudo)
     
     lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iTipoTela <> TIPOTELA_EMAIL_COBRANCA_FATURA And iTipoTela <> TIPOTELA_EMAIL_NFE Then
    
        If objDoc.lNumTitulo > 999999 Then
            sNumTitulo = String(8 - Len(CStr(objDoc.lNumTitulo)), "0") & CStr(objDoc.lNumTitulo)
        Else
            sNumTitulo = String(6 - Len(CStr(objDoc.lNumTitulo)), "0") & CStr(objDoc.lNumTitulo)
        End If
    Else
        sNumTitulo = String(6 - Len(CStr(objDoc.lNumNotaFiscal)), "0") & CStr(objDoc.lNumNotaFiscal)
    End If

    Select Case UCase(objMnemonicoValor.sMnemonico)

        Case "HTML_FAT_TRV"
            sValorTexto = sConteudo & gsEmpresaTRVHTML & sNumTitulo & ".html" & IIf(lFornOver = objDoc.lFornecedor, ";" & sConteudo & "xls\" & gsEmpresaTRVHTML & sNumTitulo & ".xls", "")
            objMnemonicoValor.colValor.Add sValorTexto
            
        Case "XLS_OVER_TRV"
            sValorTexto = sConteudo & "xls\" & gsEmpresaTRVHTML & sNumTitulo & ".xls"
            objMnemonicoValor.colValor.Add sValorTexto
            
        Case Else
            gError ERRO_SEM_MENSAGEM

    End Select

    CE_Calcula_Mnemonico_Cust = SUCESSO

    Exit Function

Erro_CE_Calcula_Mnemonico_Cust:

    CE_Calcula_Mnemonico_Cust = gErr

    Select Case gErr
       
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187164)

    End Select

    Exit Function

End Function

