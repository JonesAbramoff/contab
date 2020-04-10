VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotCargaCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Rotina_FTP_Recepcao_CC(objLoja As Object, Optional ByVal iUpDown As Integer = 0) As Long

Dim lErro As Long
Dim objTransfECF As New ClassTransfECF
Dim sRetorno As String
Dim iStatus As Integer
Dim lTamanho As Long
Dim objLojaConfig As New ClassLojaConfig

On Error GoTo Erro_Rotina_FTP_Recepcao_CC

    If gsNomePrinc = "" Then
    
        gsNomePrinc = "BatchTransfCC"
    
        lErro = Sistema_Abrir_Batch(objLoja.sNomeArqParam)
        If lErro <> SUCESSO Then gError 133435
        
    End If
    
    If MsgCC.gobjTransfECF Is Nothing Then

        objTransfECF.lIntervaloTrans = objLoja.lIntervaloTrans
        objTransfECF.sFTPDiretorio = objLoja.sFTPDiretorio
        objTransfECF.sFTPURL = objLoja.sFTPURL
        objTransfECF.sFTPUserName = objLoja.sFTPUserName
        objTransfECF.sFTPPassword = objLoja.sFTPPassword
    
        Set MsgCC.gobjTransfECF = objTransfECF
        
        MsgCC.iUpDown = iUpDown
        
        Set gcolModulo = objLoja.colModulo
        
        'se o diretorio nao for terminado por /  ===> acrescentar
        If Len(objTransfECF.sFTPDiretorio) > 0 And right(objTransfECF.sFTPDiretorio, 1) <> "/" Then objTransfECF.sFTPDiretorio = objTransfECF.sFTPDiretorio & "/"
    
        MsgCC.Show

    Else
    
        MsgCC.iUpDown = iUpDown
    
    End If

    Rotina_FTP_Recepcao_CC = SUCESSO

    Exit Function

Erro_Rotina_FTP_Recepcao_CC:

    Rotina_FTP_Recepcao_CC = gErr

    Select Case gErr

        Case 133435, 133506, 133599, 133613, 133618, 133619, 133620, 133628

        Case 133504
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_ENCONTRADO1", gErr, APLICACAO_LOJA, "CargaCC", NOME_ARQUIVO_ADM)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153384)

    End Select

    Exit Function

End Function



