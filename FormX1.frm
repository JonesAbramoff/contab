VERSION 5.00
Begin VB.Form Form1 
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sCOOInicial As String, sCOOFinal As String
Dim iStatus As Integer
Dim iStatus1 As Integer
Dim lCOOIni As Long
Dim lCOOFim As Long

            sCOOInicial = Space(6)
            sCOOFinal = Space(6)

            iStatus = rRetornarInformacao_ECF_Daruma("27", sCOOInicial)
            iStatus1 = rRetornarInformacao_ECF_Daruma("26", sCOOFinal)
            
            lCOOIni = StrParaLong(sCOOInicial)
            lCOOFim = StrParaLong(sCOOFinal)
            


End Sub


Function StrParaLong(sTexto As String) As Long
'retorna sTexto como Long

    If Len(Trim(sTexto)) = 0 Then
        StrParaLong = 0
    Else
        StrParaLong = CLng(sTexto)
    End If

End Function

