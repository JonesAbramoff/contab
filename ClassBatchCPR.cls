VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBatchCPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1

Function Rotina_Inicia_Batch_ICMS(ByVal sNomeArqParam As String, ByVal objGeracaoArqICMS As Object) As Long

On Error GoTo Erro_Rotina_Inicia_Batch_ICMS
    
    Set TelaAcompanhaBatchCPR.objGeracaoArqICMS = objGeracaoArqICMS
    TelaAcompanhaBatchCPR.Caption = TITULO_TELABATCH_GERACAO_ARQ_ICMS
    TelaAcompanhaBatchCPR.iRotinaBatch = ROTINA_BACH_GERACAO_ARQ_ICMS
    TelaAcompanhaBatchCPR.iCancelaBatch = 0
    TelaAcompanhaBatchCPR.sNomeArqParam = sNomeArqParam
    
    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchCPR.Show

    Rotina_Inicia_Batch_ICMS = SUCESSO
    
    Exit Function

Erro_Rotina_Inicia_Batch_ICMS:

    Rotina_Inicia_Batch_ICMS = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144648)

    End Select
    
    Exit Function

End Function

Function Rotina_Envia_Emails_Batch(ByVal sNomeArqParam As String, ByVal colEnvioEmail As Collection) As Long

On Error GoTo Erro_Rotina_Envia_Emails_Batch
    
    Set TelaAcompanhaBatchCPR2.gcolEnvioDeEmail = colEnvioEmail
    TelaAcompanhaBatchCPR2.Caption = "Envio de emails"
    TelaAcompanhaBatchCPR2.iCancelaBatch = 0
    TelaAcompanhaBatchCPR2.sNomeArqParam = sNomeArqParam
    
    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchCPR2.Show
    SetWindowPos TelaAcompanhaBatchCPR2.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    DoEvents

    Rotina_Envia_Emails_Batch = SUCESSO
    
    Exit Function

Erro_Rotina_Envia_Emails_Batch:

    Rotina_Envia_Emails_Batch = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196976)

    End Select
    
    Exit Function

End Function
