Attribute VB_Name = "SingleLogon"
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, ByRef pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
End Type
   
Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

' GetQueueStatus flags
Public Const QS_KEY = &H1
Public Const QS_MOUSEMOVE = &H2
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_POSTMESSAGE = &H8
Public Const QS_TIMER = &H10
Public Const QS_PAINT = &H20
Public Const QS_SENDMESSAGE = &H40
Public Const QS_HOTKEY = &H80

Public Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)

Public Const QS_INPUT = (QS_MOUSE Or QS_KEY)

Public Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

Public Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

'possible returns from MsgWaitForMultipleObjects
Public Const WAIT_OBJECT_0 = 0&
Public Const WAIT_TIMEOUT = &H102&
Public Const WAIT_ABANDONED_0 = &H80&

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_MINIMIZE = 6

Private Function EnvironmentVariable(ByVal sVariavel As String) As String
' Returns the network login name
Dim lngLen As Long, lngX As Long
Dim strEnvironmentVariable As String
    strEnvironmentVariable = String$(254, 0)
    lngLen = 255
    lngX = GetEnvironmentVariable(sVariavel, strEnvironmentVariable, lngLen)
    If (lngX > 0) Then
        EnvironmentVariable = left$(strEnvironmentVariable, lngLen - 1)
    Else
        EnvironmentVariable = vbNullString
    End If
    EnvironmentVariable = Replace(EnvironmentVariable, Chr(0), "")
End Function

Private Function ComputerName() As String
'Returns the computername
Dim lngLen As Long, lngX As Long
Dim strCompName As String
    lngLen = 16
    strCompName = String$(lngLen, 0)
    lngX = GetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        ComputerName = left$(strCompName, lngLen)
    Else
        ComputerName = ""
    End If
End Function

Private Function UserName() As String
' Returns the network login name
Dim lngLen As Long, lngX As Long
Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = GetUserName(strUserName, lngLen)
    If (lngX > 0) Then
        UserName = left$(strUserName, lngLen - 1)
    Else
        UserName = vbNullString
    End If
End Function

Private Function ExecCmd(cmdline As String) As Long
'This function is used to execute a command line function and cause the VB program
'to wait until the command has completed.

Dim Proc As PROCESS_INFORMATION
Dim Start As STARTUPINFO
Dim hProc As Long
Dim ret As Long
Dim OpenForms As Integer
Dim i As Long

On Error GoTo Erro_ExecCmd

    Start.dwFlags = STARTF_USESHOWWINDOW
    Start.wShowWindow = SW_MINIMIZE

   ' Initialize the STARTUPINFO structure:
   Start.cb = Len(Start)

   ' Start the shelled application:
   ret = CreateProcessA(0&, cmdline, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, 0&, Start, Proc)

   ' Wait for the shelled application to finish.  Note that the setup does some
   ' posting of messages to the desktop.  If we simply waited for the setup to finish
   ' via WaitForSingleObject, the VB application would not be processing these messages
   ' and the setup would hang.  So we use MsgWaitForMultipleObjects and the DoEvents()
   ' function to allow the VB app to process these messages and allow setup to finish.
   
   i = 0
 Do
     ret = MsgWaitForMultipleObjects(1&, Proc.hProcess, 0&, INFINITE, _
         (QS_POSTMESSAGE Or QS_SENDMESSAGE))
     If ret = (WAIT_OBJECT_0) Then Exit Do   'The process ended.
     OpenForms = DoEvents()
     
     '  Cut off the process if it does not respond.
     '  The higher the number, the more tolerant the installation.
     If i = 99999 Then
         CloseHandle (Proc.hProcess)
         gError 187470
     End If
     i = i + 1
 
Loop
    ret = CloseHandle(Proc.hProcess)
   
    ExecCmd = SUCESSO
   
    Exit Function
   
Erro_ExecCmd:

    ExecCmd = gErr

    Select Case gErr

        Case 187470

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187471)

    End Select

    Exit Function

End Function

Public Function Single_Logon() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iCount As Integer
Dim sNomeArquivo As String
Dim sRegistro As String
Dim iPosUserName As Integer
Dim bArquivoAberto As Boolean
Dim bAbrindoArq As Boolean
Dim iNumTentativas As Integer
Dim sUserName As String
Dim bForca As Boolean
Dim dtData As Date, bReabrir As Boolean
Dim objUsuario As New ClassDicUsuario
Dim sComputador As String

On Error GoTo Erro_Single_Logon

    'If ucase(ComputerName) = "WAGNER" Or Ucase(ComputerName) = "JONES" Then Exit Function
    
    'O usuário demo100 pode se conectar várias vezes
    If UCase(UserName) = "DEMO100" Or UCase(UserName) = "DEMO101" Or UCase(UserName) = "ADMINISTRADOR" Or left(UCase(UserName), 4) = "DEMO" Then Exit Function

    'Nome do arquivo que guarda os usuários conectados
    sNomeArquivo = App.Path & "\qwinsta\qwinsta.txt"

    'Cria o arquivo que guarda os usuários conectados
    'Call Shell("Bat_qwinsta.bat", vbMinimizedNoFocus)
    lErro = ExecCmd("Bat_qwinsta.bat")
    If lErro <> SUCESSO And lErro <> 187470 Then gError 187472
    
    If lErro <> SUCESSO Then gError 187473
        
    bAbrindoArq = True
    
Abre_o_Arquivo_denovo:

    DoEvents
    
    'Se tentar 20 vezes a validação e não conseguir -> Erro
    iNumTentativas = iNumTentativas + 1
    If iNumTentativas > 20 Then gError 187463
    
    'Pega a data de geração do arquivo
    dtData = FileDateTime(sNomeArquivo)
    
    'Se já tem mais de 20 segundos considera como arquivo antigo e espera mais um pouco
    If DateDiff("s", dtData, Now) > 20 Then GoTo Abre_o_Arquivo_denovo
    
    'Tenta abrir o arquivo
    bReabrir = False
    Open sNomeArquivo For Input As #1
    If bReabrir Then GoTo Abre_o_Arquivo_denovo
    
    bAbrindoArq = False
    bArquivoAberto = True
    
    'Busca o primeiro registro do arquivo
    Line Input #1, sRegistro
    
    'Se está em branco é porque ainda está sendo gerado ... fecha e manda abrir de novo
    If Len(Trim(sRegistro)) = 0 Then
        Close #1
        GoTo Abre_o_Arquivo_denovo
    End If
    
    'Busca a posição do nome do usuário
    iPosUserName = InStr(1, sRegistro, "USERNAME")
    If iPosUserName = 0 Then iPosUserName = InStr(1, sRegistro, "NOMEUTILIZADOR")

    'Enquanto existirem usuários conectados
    bForca = False
    Do While Not EOF(1) Or bForca

        'Pega o usuário
        sUserName = Trim(Mid(sRegistro, iPosUserName, 20))
        
        'Conta quantas vezes ele está conectado
        If UCase(sUserName) = UCase(UserName) Then
            iCount = iCount + 1
        End If
        
        'Se já tratou a última linha sai do loop
        If bForca Then Exit Do
        
        'Busca o próximo registro
        Line Input #1, sRegistro

        'Força a leitura da última linha
        If EOF(1) Then bForca = True

    Loop
    
    'Fecha o arquivo
    Close #1
    bArquivoAberto = False
    
    'Se existirem mais de 1 usuário conectado com o mesmo login -> erro
    If iCount > 1 Then
    
        lErro = Usuario_Le_Login(UserName, sComputador, objUsuario)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187469
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then sComputador = ""
        
        gError 187464
    
    End If
    
    Single_Logon = SUCESSO
       
    Exit Function
    
Erro_Single_Logon:
    
    Single_Logon = gErr
    
    Select Case gErr
    
        Case 53
            bReabrir = True
            Resume Next
            
        Case 187463
'            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_TENTATIVAS_EXCEDIDO", gErr)

        Case 187464
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_JA_LOGADO", gErr, UserName, sComputador, objUsuario.sNome)
            
        Case 187469, 187472
        
        Case 76, 187473
'            Call Rotina_Erro(vbOKOnly, "ERRO_FALTANDO_ARQ_BAT_QWINSTA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165222)

    End Select
    
    If bArquivoAberto Then Close #1

    Exit Function
    
End Function

