VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFiscalItens_Grava_Customizado(ByVal objNFiscal As ClassNFiscal) As Long
    NFiscalItens_Grava_Customizado = CF("NFiscal_Grava_Customizado", objNFiscal)
End Function

'Incluído gravação dos campos comprimento, pesokg e vendapeso
Function NFiscal_Grava_Customizado(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim objNFiscalMgz As ClassNFiscalMgz
Dim lNumIntDoc As Long

On Error GoTo Erro_NFiscal_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141979
    Next
    
    If Not (objNFiscal.objInfoUsu Is Nothing) Then
    
        'Pesquisa os ítens vinculados à Nota Fiscal
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM NFiscal WHERE NumIntDoc = ? ", 0, lNumIntDoc, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141980

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141981

        If lErro = AD_SQL_SUCESSO Then
        
            Set objNFiscalMgz = objNFiscal.objInfoUsu
            
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE NFiscal SET Contrato = ?, Processo = ?, IDCorrespondencia = ? ", alComando(0), objNFiscalMgz.sContrato, objNFiscalMgz.sProcesso, objNFiscalMgz.sIDCorrespondencia)
            If lErro <> AD_SQL_SUCESSO Then gError 141982

        End If

    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscal_Grava_Customizado = SUCESSO

    Exit Function

Erro_NFiscal_Grava_Customizado:

    NFiscal_Grava_Customizado = gErr

    Select Case gErr

        Case 141979
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141980, 141981
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 141982
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_NFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141983)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

