VERSION 5.00
Begin VB.Form TestaTempFile 
   Caption         =   "Teste de criação de Arquivo temporario"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CLIQUE AQUI !"
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   540
      Width           =   2880
   End
End
Attribute VB_Name = "TestaTempFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit

      Private Declare Function GetTempPath Lib "kernel32" _
         Alias "GetTempPathA" (ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

      Private Declare Function GetTempFileName Lib "kernel32" _
         Alias "GetTempFileNameA" (ByVal lpszPath As String, _
         ByVal lpPrefixString As String, ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long

      Private Function CreateTempFile(sPrefix As String) As String
         Dim sTmpPath As String * 512
         Dim sTmpName As String * 576
         Dim nRet As Long

         nRet = GetTempPath(512, sTmpPath)
         If (nRet > 0 And nRet < 512) Then
            nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
            If nRet <> 0 Then
               CreateTempFile = Left$(sTmpName, _
                  InStr(sTmpName, vbNullChar) - 1)
            End If
         End If
      End Function

      Private Sub Command1_Click()
         Dim sTmpFile As String
         Dim sMsg As String
         Dim hFile As Long

         sTmpFile = CreateTempFile("mail")
         hFile = FreeFile

         Open sTmpFile For Binary As hFile
            Put #hFile, , "This is a test. 1234"
         Close hFile

         sMsg = "Temp FileName: " & sTmpFile & vbCrLf
         sMsg = sMsg & "File Length: " & FileLen(sTmpFile) & vbCrLf
         sMsg = sMsg & "Time Created: " & _
            Format$(FileDateTime(sTmpFile), "long time") & vbCrLf

         MsgBox sMsg, vbInformation, "TempFile"

         Kill sTmpFile
      End Sub

