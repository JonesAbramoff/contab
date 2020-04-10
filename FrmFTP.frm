VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmFTP 
   Caption         =   "FTP"
   ClientHeight    =   795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   1740
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
End
Attribute VB_Name = "FrmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gsMsg As String
Public gsMsg1 As String
Public giUsouFTP As Integer
Public giTerminou As Integer
Public giErro As Integer

Const FTP_UPLOAD_ESPERA_EM_MIN = 30

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    giTerminou = DESMARCADO
    giErro = DESMARCADO
    giUsouFTP = DESMARCADO
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    
Dim sTeste As String
Dim sMsg1 As String
Dim lErro As Long

On Error GoTo Erro_Inet1_StateChanged

    Select Case State
    
        Case icResolvingHost
            gsMsg = "Pesquisando IP..."
              
        Case icHostResolved
            gsMsg = "IP encontrado"
        
        Case icReceivingResponse
            gsMsg = "Recebendo mensagem..."
            
        Case icResponseCompleted
            sTeste = "a"
            Do While Len(sTeste) > 0
                sTeste = Inet1.GetChunk(1000)
                sMsg1 = sMsg1 & sTeste
            Loop
            gsMsg1 = sMsg1
            gsMsg = "Mensagem completada"
            giTerminou = MARCADO
        
        Case icConnecting
            gsMsg = "Conectando..."
            
        Case icConnected
            giUsouFTP = 1
            gsMsg = "Conectado"
            
        Case icRequesting
            gsMsg = "Enviando pedido ao servidor..."
            
        Case icRequestSent
            gsMsg = "Pedido enviado ao servidor"
            
        Case icDisconnecting
            gsMsg = "Desconectando..."
            
        Case icDisconnected
            giUsouFTP = 0
            gsMsg = "Desconectado"
    
        Case icError
            gsMsg = "Erro de comunicação"
            giErro = MARCADO
    
        Case icResponseReceived
            gsMsg = "Mensagem recebida"
    
    End Select
    
    Exit Sub
    
Erro_Inet1_StateChanged:
    
    Select Case gErr

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211923)

    End Select

    Exit Sub
    
End Sub

Public Function Fazer_Upload(ByVal sOrigem As String, sDestino As String) As Long

Dim lErro As Long
Dim lTeste As Long, iIndice As Integer
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Fazer_Upload

    giErro = 0
    giTerminou = 0
    
    'If Inet1.StillExecuting Then Inet1.Cancel

    Inet1.Execute , "PUT " & sOrigem & " " & sDestino

    vbMsgBox = vbYes
    'Se não terminou, não deu erro, e o usuário pediu para esperar mais um pouco ou é a 1a tentativa
    Do While giTerminou <> MARCADO And vbMsgBox = vbYes And giErro <> MARCADO
        lTeste = 0
        Do While giTerminou <> MARCADO And lTeste < FTP_UPLOAD_ESPERA_EM_MIN And giErro <> MARCADO
            For iIndice = 1 To 60
                Call Sleep(1000)
            Next
            lTeste = lTeste + 1
            DoEvents
        Loop
        If giTerminou <> MARCADO And giErro <> MARCADO Then
            If UCase(gsUsuario) <> "BACKUP" Then vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_UPLOAD_ARQUIVO_FTP_SEM_RESPOSTA", sOrigem, FTP_UPLOAD_ESPERA_EM_MIN)
        End If
    Loop
    If giTerminou <> MARCADO Then gError ERRO_SEM_MENSAGEM 'Não conseguiu fazer o upload
    
    Fazer_Upload = SUCESSO

    Exit Function
    
Erro_Fazer_Upload:

    Fazer_Upload = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211924)

    End Select

    'If Inet1.StillExecuting Then Inet1.Cancel

    Exit Function
    
End Function

Public Function Apagar_Arquivo_FTP(ByVal sArquivo As String) As Long

Dim lErro As Long
Dim lTeste As Long, iIndice As Integer
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Apagar_Arquivo_FTP

    giErro = 0
    giTerminou = 0
    
    'If Inet1.StillExecuting Then Inet1.Cancel

    Inet1.Execute , "DELETE " & sArquivo

    lTeste = 0
    Do While giTerminou <> MARCADO And lTeste < FTP_UPLOAD_ESPERA_EM_MIN And giErro <> MARCADO
        For iIndice = 1 To 60
            Call Sleep(1000)
        Next
        lTeste = lTeste + 1
        DoEvents
    Loop
    If giTerminou <> MARCADO Then gError ERRO_SEM_MENSAGEM 'Não conseguiu excluir
    
    Apagar_Arquivo_FTP = SUCESSO

    Exit Function
    
Erro_Apagar_Arquivo_FTP:

    Apagar_Arquivo_FTP = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211924)

    End Select

    'If Inet1.StillExecuting Then Inet1.Cancel

    Exit Function
    
End Function

Public Function Fazer_Download(ByVal sOrigem As String, sDestino As String) As Long

Dim lErro As Long
Dim lTeste As Long, iIndice As Integer
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Fazer_Download

    'If Inet1.StillExecuting Then Inet1.Cancel
    
    giErro = 0
    giTerminou = 0
    
    Inet1.Execute , "GET " & sOrigem & " " & sDestino

    vbMsgBox = vbYes
    'Se não terminou, não deu erro, e o usuário pediu para esperar mais um pouco ou é a 1a tentativa
    Do While giTerminou <> MARCADO And vbMsgBox = vbYes And giErro <> MARCADO
        lTeste = 0
        Do While giTerminou <> MARCADO And lTeste < FTP_UPLOAD_ESPERA_EM_MIN And giErro <> MARCADO
            For iIndice = 1 To 60
                Call Sleep(1000)
            Next
            lTeste = lTeste + 1
            DoEvents
        Loop
        If giTerminou <> MARCADO And giErro <> MARCADO Then
            vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_DOWNLOAD_ARQUIVO_FTP_SEM_RESPOSTA", sOrigem, FTP_UPLOAD_ESPERA_EM_MIN)
        End If
    Loop
    If giTerminou <> MARCADO Then gError ERRO_SEM_MENSAGEM 'Não conseguiu fazer o upload
    
    Fazer_Download = SUCESSO

    Exit Function
    
Erro_Fazer_Download:

    Fazer_Download = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211924)

    End Select

    'If Inet1.StillExecuting Then Inet1.Cancel

    Exit Function
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    
On Error GoTo Erro_Form_UnLoad

    If giUsouFTP = 1 Then Inet1.Execute , "QUIT"
    
    Exit Sub
    
Erro_Form_UnLoad:
    
    Select Case gErr

        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211925)

    End Select

    Exit Sub
    
End Sub
