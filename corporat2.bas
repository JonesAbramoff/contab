Attribute VB_Name = "PrincCorporat2"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Const SUCESSO = 0
Public Const NOME_ARQUIVO_ADM = "ADM100.INI"
Public Const DLL_A_REGISTRAR = "Não há DLLs a Registrar"  'Inserir nome da DLL que deve ser registrada

Public iParaTudo As Integer
Public colOutrosArquivos As New Collection

Sub Main()

Dim lErro As Long
Dim sDirOrigem As String
Dim sDirDestino As String
Dim sNomeArq As String
Dim dDataArqOrigem As Date
Dim dDataArqDestino As Date
Dim bCopiaArq As Boolean
Dim sAtualiza As String

On Error GoTo Erro_Principal

   sAtualiza = LeArqINI("Geral", "AutoAtualiza", NOME_ARQUIVO_ADM)
   If Len(sAtualiza) = 0 Then Error 5001

   If sAtualiza = "1" Then
              
        sDirDestino = App.Path
        If Right(sDirDestino, 1) <> "\" Then
           sDirDestino = sDirDestino & "\"
        End If
        
        sDirOrigem = LeArqINI("Geral", "DirProgram", NOME_ARQUIVO_ADM)
        If Len(sDirOrigem) = 0 Then Error 5007
        
        If Right(sDirOrigem, 1) <> "\" Then
           sDirOrigem = sDirOrigem & "\"
        End If

        sNomeArq = "Corporator.exe"
        
        If Len(Dir(sDirOrigem & sNomeArq)) = 0 Then Error 5008
        
        bCopiaArq = False
        dDataArqOrigem = FileDateTime(sDirOrigem & sNomeArq)
         
        If Len(Dir(sDirDestino & sNomeArq)) = 0 Then
            
           bCopiaArq = True
         
        Else
         
           dDataArqDestino = FileDateTime(sDirDestino & sNomeArq)
         
           If dDataArqOrigem > dDataArqDestino Then
              bCopiaArq = True
           End If
         
        End If
         
        If bCopiaArq Then
            
           FileCopy sDirOrigem & sNomeArq, sDirDestino & sNomeArq
            
        End If
   
    End If

    Shell "Corporator.exe", vbMaximized
   
    End

Erro_Principal:

    Beep
   
    Select Case Err.Number
   
        Case 5001
            MsgBox "Não há informação sobre o atualização automática no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
 
        Case 5007
            MsgBox "Não há informação sobre os arquivos de atualização no " & NOME_ARQUIVO_ADM, vbCritical, "Atenção!"
      
        Case 5008
         
            If MsgBox("O arquivo " & Chr(34) & sDirOrigem & sNomeArq & Chr(34) & " não existe. Deseja continuar mesmo assim?", vbYesNo + vbQuestion + vbDefaultButton2, "Continua?") = vbYes Then
                Resume Next
            End If
            
      Case Else
         MsgBox "Erro do VB! (" & Err.Number & " - " & Err.Description & ").", vbCritical, "Atenção!"
      
   End Select
   
   End

End Sub

Public Function LeArqINI(sSecao As String, sItem As String, sArqIni As String) As String
'Lê valor string no arquivo ADM100.INI
Dim sConteudo As String * 128
Dim lRetorno As Long
   
On Error GoTo Erro_LeArqINI
   
   lRetorno = GetPrivateProfileString(sSecao, sItem, "", sConteudo, 128, sArqIni)
   
   LeArqINI = Left(sConteudo, lRetorno)

   Exit Function
   
Erro_LeArqINI:

   LeArqINI = ""

   Exit Function

End Function

Public Function Extrai_Diretorio(sNomeArq As String) As String
'Extrai o nome do diretório do nome do arquivo
Dim iUltimaPos As Integer
Dim iBarraPos As Integer

On Error GoTo Erro_Extrai_Diretorio

   iUltimaPos = 2                     'depois do c:
   
   iBarraPos = InStr(sNomeArq, "\")
   
   Do While iBarraPos > 0
      
      iUltimaPos = iBarraPos
      iBarraPos = InStr(iBarraPos + 1, sNomeArq, "\")
   
   Loop
   
   Extrai_Diretorio = Mid(sNomeArq, 1, iUltimaPos)
   
   Exit Function

Erro_Extrai_Diretorio:

   Extrai_Diretorio = ""
   
   Exit Function
   
End Function
