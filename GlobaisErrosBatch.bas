Attribute VB_Name = "GlobaisErrosBatch"
Option Explicit

'Sub Main(ByVal sDescricao As String, ByVal sEmpresa As String, ByVal sFilial As String, ByVal sUsuario As String, ByVal sTexto As String)

Public Sub Main()

Dim X As New FormMsgErroBatch2
Dim sParam As String

    sParam = Command$
            
    Call X.Inicializa(sParam)
    X.Show
    X.SetFocus
        
End Sub
