VERSION 5.00
Begin VB.Form FormTesteFTP 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormTesteFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim objFTP As Object
Dim sErro As String
Dim lErro As Long

    Set objFTP = CreateObject("SGEUtil.FTPCli")
    
    lErro = objFTP.Upload_Arquivo("ftp.corporator.com.br", "corporator", "abc123...", "c:\sge\ecf\dadoscc\1_1_dadoscc.ccc", "demo_ecf_mario/1_1_dadosccjones.ccc", sErro)
    If lErro <> 0 Then
        MsgBox (sErro)
    Else
        MsgBox ("sucesso")
    End If

End Sub
