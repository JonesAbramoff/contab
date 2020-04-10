VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotTransfECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Rotina_FTP_Envio_Caixa(objLojaECF As Object, Optional ByVal iDownload As Integer = 0) As Long

Dim lErro As Long
Dim objTransfECF As New ClassTransfECF

On Error GoTo Erro_Rotina_FTP_Envio_Caixa

    If MsgECF.gobjTransfECF Is Nothing Then

        gsNomePrinc = "BatchTransfECF"
    
        objTransfECF.lIntervaloTrans = objLojaECF.lIntervaloTrans
        objTransfECF.sFTPDiretorio = objLojaECF.sFTPDiretorio
        objTransfECF.sFTPURL = objLojaECF.sFTPURL
        objTransfECF.sFTPUserName = objLojaECF.sFTPUserName
        objTransfECF.sFTPPassword = objLojaECF.sFTPPassword
        
        gdtDataHoje = objLojaECF.dtData
    
        Set MsgECF.gobjTransfECF = objTransfECF
        
        
        'se o diretorio nao for terminado por /  ===> acrescentar
        If Len(objTransfECF.sFTPDiretorio) > 0 And right(objTransfECF.sFTPDiretorio, 1) <> "/" Then objTransfECF.sFTPDiretorio = objTransfECF.sFTPDiretorio & "/"
    
        MsgECF.Show

    End If

    If iDownload = 1 Then
    
        lErro = MsgECF.Download_DadosCC()
        If lErro <> SUCESSO Then gError 133577
        
    ElseIf iDownload = 2 Then
        
        lErro = MsgECF.Upload_Arq()
        If lErro <> SUCESSO Then gError 133620
        
    End If

    If MsgECF.gobjTransfECF.lIntervaloTrans = 0 Then
        Unload MsgECF
    End If

    Rotina_FTP_Envio_Caixa = SUCESSO

    Exit Function

Erro_Rotina_FTP_Envio_Caixa:

    Rotina_FTP_Envio_Caixa = gErr

    Select Case gErr

        Case 133577, 133620

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 153385)

    End Select

    Exit Function

End Function


