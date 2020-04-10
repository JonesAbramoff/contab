Attribute VB_Name = "ZipCompacta"
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const SW_HIDE = 0
Private Const STARTF_USESHOWWINDOW = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

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

Private Declare Function addZIP Lib "azip32.dll" () As Integer
Private Declare Function addZIP_ArchiveName Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_ClearAttributes Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_Comment Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_Delete Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_DeleteComment Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_DisplayComment Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_Encrypt Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_Exclude Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_ExcludeListFile Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_GetLastError Lib "azip32.dll" () As Integer
Private Declare Function addZIP_GetLastWarning Lib "azip32.dll" () As Integer
Private Declare Function addZIP_Include Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_IncludeArchive Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Private Declare Function addZIP_IncludeDirectoryEntries Lib "azip32.dll" (ByVal flag As Integer) As Integer
Private Declare Function addZIP_IncludeFilesNewer Lib "azip32.dll" (ByVal DateVal As String) As Integer
Private Declare Function addZIP_IncludeFilesOlder Lib "azip32.dll" (ByVal DateVal As String) As Integer
Private Declare Function addZIP_IncludeHidden Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Private Declare Function addZIP_IncludeListFile Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_IncludeReadOnly Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Private Declare Function addZIP_IncludeSystem Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Private Declare Sub addZIP_Initialise Lib "azip32.dll" ()
Private Declare Function addZIP_InstallCallback Lib "azip32.dll" (ByVal cbFunction As Long) As Integer
Private Declare Function addZIP_Overwrite Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_Recurse Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_Register Lib "azip32.dll" (ByVal lpStr As String, ByVal Uint32 As Long) As Integer
Private Declare Function addZIP_SaveAttributes Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_SaveStructure Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_SetArchiveDate Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_SetCompressionLevel Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_SetParentWindowHandle Lib "azip32.dll" (ByVal hWnd As Long) As Integer
Private Declare Function addZIP_SetTempDrive Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_SetWindowHandle Lib "azip32.dll" (ByVal hWnd As Long) As Integer
Private Declare Function addZIP_Span Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_Store Lib "azip32.dll" (ByVal lpStr As String) As Integer
Private Declare Function addZIP_UseLFN Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Private Declare Function addZIP_View Lib "azip32.dll" (ByVal Int16 As Integer) As Integer

'  constants for addZIP_SetCompressionLevel(...)
Private Const azCOMPRESSION_MAXIMUM = &H3
Private Const azCOMPRESSION_MINIMUM = &H1
Private Const azCOMPRESSION_NONE = &H0
Private Const azCOMPRESSION_NORMAL = &H2

' constants for addZIP_SaveStructure(...)
Private Const azSTRUCTURE_ABSOLUTE = &H2
Private Const azSTRUCTURE_NONE = &H0
Private Const azSTRUCTURE_RELATIVE = &H1

' constants for addZIP_Overwrite(...)
Private Const azOVERWRITE_ALL = &HB
Private Const azOVERWRITE_NONE = &HC
Private Const azOVERWRITE_QUERY = &HA

' constants for addZIP_SetArchiveDate()
Private Const DATE_NEWEST = &H3
Private Const DATE_OLDEST = &H2
Private Const DATE_ORIGINAL = &H0
Private Const DATE_TODAY = &H1

' constants for addZIP_IncludeXXX attribute functions
Private Const azNEVER = &H0       ' files must never have this attribute set
Private Const azALWAYS = &HFF ' files may or may not have this attribute set
Private Const azYES = &H1         ' files must always have this attribute set

'  constants for addZIP_ClearAttributes(...)
Private Const azATTR_NONE = 0
Private Const azATTR_READONLY = 1
Private Const azATTR_HIDDEN = 2
Private Const azATTR_SYSTEM = 4
Private Const azATTR_ARCHIVE = 32
Private Const azATTR_ALL = 39

' constants used in messages to identify library
Private Const azLIBRARY_ADDZIP = 0

' 'messages' used to provide information to the calling program
Private Const AM_SEARCHING = &HA
Private Const AM_ZIPCOMMENT = &HB
Private Const AM_ZIPPING = &HC
Private Const AM_ZIPPED = &HD
Private Const AM_UNZIPPING = &HE
Private Const AM_UNZIPPED = &HF
Private Const AM_TESTING = &H10
Private Const AM_TESTED = &H11
Private Const AM_DELETING = &H12
Private Const AM_DELETED = &H13
Private Const AM_DISKCHANGE = &H14
Private Const AM_VIEW = &H15
Private Const AM_ERROR = &H16
Private Const AM_WARNING = &H17
Private Const AM_QUERYOVERWRITE = &H18
Private Const AM_COPYING = &H19
Private Const AM_COPIED = &H1A
Private Const AM_ABORT = &HFF

' Constants for whether file is encrypted or not in AM_VIEW
Private Const azFT_ENCRYPTED = &H1
Private Const azFT_NOT_ENCRYPTED = &H0

' Constants for whether file is text or binary in AM_VIEW
Private Const azFT_BINARY = &H1
Private Const azFT_TEXT = &H0

' Constants for compression method in AM_VIEW
Private Const azCM_DEFLATED_FAST = &H52
Private Const azCM_DEFLATED_MAXIMUM = &H51
Private Const azCM_DEFLATED_NORMAL = &H50
Private Const azCM_DEFLATED_SUPERFAST = &H53
Private Const azCM_IMPLODED = &H3C
Private Const azCM_NONE = &H0
Private Const azCM_REDUCED_1 = &H14
Private Const azCM_REDUCED_2 = &H1E
Private Const azCM_REDUCED_3 = &H28
Private Const azCM_REDUCED_4 = &H32
Private Const azCM_SHRUNK = &HA
Private Const azCM_TOKENISED = &H46
Private Const azCM_UNKNOWN = &HFF

' Constants used in returning from a AM_QUERYOVERWRITE message
Private Const azOW_NO = &H2
Private Const azOW_NO_TO_ALL = &H3
Private Const azOW_YES = &H0
Private Const azOW_YES_TO_ALL = &H1

Public Function Zip_Compacta(cArqCompactado As String, cArq As String) As Long

Dim lErro As Long
Dim sDiretorio As String
Dim lRetorno As Long
Dim sRetVal As String

On Error GoTo Erro_Zip_Compacta

    If bExisteFrmWrk Then
    
        sDiretorio = String(255, 0)
        lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
        sDiretorio = left(sDiretorio, lRetorno)
        
        Call ExecCmd(sDiretorio & "Corporator_Zip " & Replace(cArq, "*.*", "") & ", " & cArqCompactado)
    
    Else

        Call addZIP_Initialise

        lErro = addZIP_Register("UBS, INC.", 600365060)
        
        'Compacta um ou mais arquivos no formato WinZip
        lErro = addZIP_SetCompressionLevel(azCOMPRESSION_MAXIMUM)
        lErro = addZIP_SaveStructure(azSTRUCTURE_NONE) 'StoreFullPathName - azSTRUCTURE_ABSOLUTE
        
        lErro = addZIP_ArchiveName(cArqCompactado)
        lErro = addZIP_Include(cArq)
        lErro = addZIP
        
    End If
    
    sRetVal = Dir(cArqCompactado)
    If Len(Trim(sRetVal)) = 0 Then gError 209255
    
    Exit Function
    
Erro_Zip_Compacta:

    Zip_Compacta = gErr
     
    Select Case gErr
    
        Case 209255 'ERRO_PREPARACAO_ARQUIVO_TEMP
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_ARQUIVO_TEMP", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209255)
     
    End Select
     
    Exit Function
    
End Function

Public Function Zip_Copia_Arquivo(ByVal sDir As String, ByVal sDirDestino As String) As Long

On Error GoTo Erro_Zip_Copia_Arquivo

    FileCopy sDir, sDirDestino

    Exit Function
    
Erro_Zip_Copia_Arquivo:

    Zip_Copia_Arquivo = gErr
     
    Select Case gErr
          
        Case 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209250)
     
    End Select
     
    Exit Function

End Function

Public Function Zip_Cria_Diretorio(ByVal sDir As String) As Long

On Error GoTo Erro_Zip_Cria_Diretorio

    'Se o diretório já existe primeiro exclui para depois recriá-lo
    If Len(Trim(Dir(sDir, vbDirectory))) <> 0 Then
        lErro = Zip_Exclui_Diretorio(sDir)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    MkDir sDir 'Cria o diretório temporário

    Exit Function
    
Erro_Zip_Cria_Diretorio:

    Zip_Cria_Diretorio = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209251)
     
    End Select
     
    Exit Function

End Function

Public Function Zip_Exclui_Diretorio(ByVal sDir As String) As Long

On Error GoTo Erro_Zip_Exclui_Diretorio

    Call Zip_Exclui_Diretorio1(sDir) 'Apaga os arquivo da pasta se tiver
    RmDir sDir 'Só consegue excluir a pasta se não tiver arquivos

    Exit Function
    
Erro_Zip_Exclui_Diretorio:

    Zip_Exclui_Diretorio = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209252)
     
    End Select
     
    Exit Function

End Function

Public Sub Zip_Exclui_Diretorio1(ByVal sDir As String)

On Error GoTo Erro_Zip_Exclui_Diretorio1

    Kill sDir & "*.*" 'Exclui todos os arquivos do diretório

    Exit Sub
    
Erro_Zip_Exclui_Diretorio1:
     
    Exit Sub

End Sub

Public Function Zip_Exclui_Arquivo(ByVal sArquivo As String) As Long

On Error GoTo Erro_Zip_Exclui_Arquivo

    Kill sArquivo 'Exclui um arquivo

    Exit Function
    
Erro_Zip_Exclui_Arquivo:

    Zip_Exclui_Arquivo = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209253)
     
    End Select
     
    Exit Function

End Function

Function Zip_Verifica_Existencia_Arquivo(ByVal sArquivo As String) As Long
'Testa para ver se os arquivos existem

Dim lErro As Long
Dim iPos As Integer
Dim iPosAnt As Integer
Dim sArqAux As String
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Zip_Verifica_Existencia_Arquivo

    If Len(Trim(sArquivo)) > 0 Then

        sArqAux = sArquivo
        iPosAnt = 0
        iPos = InStr(iPosAnt + 1, sArquivo, ";")
        Do While iPos <> 0
            sArqAux = Mid(sArquivo, iPosAnt + 1, iPos - (iPosAnt + 1))
            Open Trim(sArqAux) For Input As #1
            Close #1
            iPosAnt = iPos
            iPos = InStr(iPosAnt + 1, sArquivo, ";")
        Loop
        If iPosAnt <> 0 Then
            sArqAux = Mid(sArquivo, iPosAnt + 1)
        End If
        Open Trim(sArqAux) For Input As #1
        Close #1
        
    End If

    Zip_Verifica_Existencia_Arquivo = SUCESSO

    Exit Function

Erro_Zip_Verifica_Existencia_Arquivo:

    Zip_Verifica_Existencia_Arquivo = gErr

    Select Case gErr
    
        Case 53, 52
            'Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_FTP_NAO_ENCONTRADO", gErr, sArqAux)
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_ARQUIVO_FTP_NAO_ENCONTRADO", sArqAux)
            If vbResult = vbYes Then Zip_Verifica_Existencia_Arquivo = ERRO_ARQUIVO_NAO_ENCONTRADO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209254)

    End Select

    Exit Function

End Function

Private Function ExecCmd(cmdline$)

Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim RET&
    
    start.dwFlags = STARTF_USESHOWWINDOW
    start.wShowWindow = SW_HIDE
    
    start.cb = Len(start)
    
    RET& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

    RET& = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, RET&)
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    ExecCmd = RET&
End Function

Private Function bExisteFrmWrk() As Boolean

Dim sDir As String
Dim sUsaFrmWrk As String

On Error GoTo Erro_bExisteFrmWrk

    'Incluído para evitar testes em instalações de Framwork problemáticas
    sUsaFrmWrk = String(255, 0)
    lRetorno = GetPrivateProfileString("Geral", "UsaFrmWrk", "", sUsaFrmWrk, 128, "ADM100.INI")
    sUsaFrmWrk = left(sUsaFrmWrk, lRetorno)

    If sUsaFrmWrk = "0" Then
        bExisteFrmWrk = False
    Else
        sDir = GetWinDir & "Microsoft.NET\Framework"
        
        If Len(Trim(Dir(sDir, vbDirectory))) = 0 Then
            bExisteFrmWrk = False
        Else
            bExisteFrmWrk = True
        End If
    End If

    Exit Function
    
Erro_bExisteFrmWrk:

    bExisteFrmWrk = False
    
    Exit Function

End Function

Private Function GetWinDir() As String
  Dim strFolder As String
  Dim lngResult As Long
  strFolder = String(255, 0)
  lngResult = GetWindowsDirectory(strFolder, 255)
  If lngResult <> 0 Then
    If right(left(strFolder, lngResult), 1) = "\" Then
      GetWinDir = left(strFolder, lngResult)
    Else
      GetWinDir = left(strFolder, lngResult) & "\"
    End If
  Else
    GetWinDir = ""
  End If
End Function
